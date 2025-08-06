// ==UserScript==
// @name         SCOPSR文章保存器
// @namespace    http://tampermonkey.net/
// @version      3.5
// @description  一键保存SCOPSR网站的文章为DOCX格式，支持多页批量保存，保留原有格式
// @author       You
// @match        https://www.scopsr.gov.cn/was5/web/search*
// @match        https://www.scopsr.gov.cn/xwzx/*
// @match        http://www.scopsr.gov.cn/was5/web/search*
// @match        http://www.scopsr.gov.cn/xwzx/*
// @grant        GM_download
// @grant        GM_addStyle
// @grant        GM_xmlhttpRequest
// ==/UserScript==

(function() {
    'use strict';

    // 添加保存按钮样式
    GM_addStyle(`
        .save-articles-btn {
            position: fixed;
            top: 100px;
            right: 20px;
            z-index: 9999;
            background-color: #4CAF50;
            color: white;
            padding: 10px 20px;
            border: none;
            border-radius: 5px;
            cursor: pointer;
            font-size: 16px;
            box-shadow: 0 4px 6px rgba(0,0,0,0.1);
        }
        .save-articles-btn:hover {
            background-color: #45a049;
        }
        .save-articles-btn:disabled {
            background-color: #cccccc;
            cursor: not-allowed;
        }
    `);

    // 判断当前页面类型
    const isSearchPage = window.location.href.includes('/was5/web/search');
    const isArticlePage = window.location.href.includes('/xwzx/');

    // 创建Word兼容的HTML文档
    function createWordHTML(articles) {
        const html = `<!DOCTYPE html>
<html xmlns:o='urn:schemas-microsoft-com:office:office' 
      xmlns:w='urn:schemas-microsoft-com:office:word' 
      xmlns='http://www.w3.org/TR/REC-html40'>
<head>
    <meta charset='utf-8'>
    <title>SCOPSR文章集</title>
    <!--[if gte mso 9]>
    <xml>
        <w:WordDocument>
            <w:View>Print</w:View>
            <w:Zoom>100</w:Zoom>
            <w:DoNotOptimizeForBrowser/>
        </w:WordDocument>
    </xml>
    <![endif]-->
    <style>
        @page {
            size: A4;
            margin: 2.54cm;
            mso-page-orientation: portrait;
        }
        
        body {
            font-family: '宋体', SimSun, serif;
            font-size: 12pt;
            line-height: 1.8;
            color: #000;
            background: white;
        }
        
        h1 {
            font-size: 22pt;
            font-weight: bold;
            text-align: center;
            margin: 20pt 0;
            color: #000;
            mso-pagination: none;
        }
        
        h2 {
            font-size: 16pt;
            font-weight: bold;
            margin: 15pt 0 10pt 0;
            color: #000;
            page-break-before: always;
            mso-pagination: none;
        }
        
        .article-info {
            text-align: center;
            color: #666;
            font-size: 10pt;
            margin: 10pt 0;
        }
        
        .article-content p {
            text-indent: 2em;
            margin: 10pt 0;
            line-height: 1.8;
            text-align: justify;
        }
        
        .page-break {
            page-break-after: always;
            mso-special-character: line-break;
            page-break-before: always;
        }
    </style>
</head>
<body>
    <div class='Section1'>
        <h1>SCOPSR文章集</h1>
        <p style='text-align: center; color: #666;'>
            生成日期：${new Date().toLocaleDateString('zh-CN')}
        </p>
        <p style='text-align: center; color: #666;'>
            共收录 ${articles.length} 篇文章${articles.some(a => a.pageNumber) ? ' (来自' + Math.max(...articles.filter(a => a.pageNumber).map(a => a.pageNumber)) + '页)' : ''}
        </p>
        
        ${articles.map((article, index) => `
            <div class='article' ${index > 0 ? "style='page-break-before: always;'" : ""}>
                <h2>${escapeHtml(article.title)}</h2>
                <div class='article-info'>
                    <p>发布时间：${article.date}</p>
                    <p>来源：${article.url}</p>
                </div>
                <div class='article-content'>
                    ${formatContent(article.content)}
                </div>
            </div>
        `).join('')}
    </div>
</body>
</html>`;
        
        return html;
    }

    // HTML转义
    function escapeHtml(text) {
        const map = {
            '&': '&amp;',
            '<': '&lt;',
            '>': '&gt;',
            '"': '&quot;',
            "'": '&#039;'
        };
        return text.replace(/[&<>"']/g, m => map[m]);
    }

    // 格式化文章内容，保留段落和换行
    function formatContent(content) {
        if (!content || content === '获取失败' || content === '无法获取文章内容') {
            return '<p style="color: red;">文章内容获取失败</p>';
        }
        
        // 按换行分割，每行作为一个段落
        const lines = content.split('\n').filter(line => line.trim());
        
        return lines.map(line => {
            const escapedLine = escapeHtml(line.trim());
            // 如果是空行或只有空格，返回空段落
            if (!escapedLine) {
                return '<p>&nbsp;</p>';
            }
            return `<p>${escapedLine}</p>`;
        }).join('\n');
    }

    // 保存为DOCX文件（实际是HTML，但Word可以打开）
    function saveAsDocx(html, filename) {
        const blob = new Blob(['\ufeff', html], {
            type: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document'
        });
        
        const url = URL.createObjectURL(blob);
        GM_download({
            url: url,
            name: filename,
            onload: () => {
                URL.revokeObjectURL(url);
            },
            onerror: (error) => {
                console.error('下载失败:', error);
                // 备用下载方式
                const a = document.createElement('a');
                a.href = url;
                a.download = filename;
                a.click();
                URL.revokeObjectURL(url);
            }
        });
    }

    if (isSearchPage) {
        // 搜索结果页面：只添加批量保存按钮（可处理单页或多页）
        const multiPageBtn = document.createElement('button');
        multiPageBtn.textContent = '批量保存';
        multiPageBtn.className = 'save-articles-btn';
        multiPageBtn.onclick = saveMultiplePages;
        document.body.appendChild(multiPageBtn);


        // 多页保存功能
        async function saveMultiplePages() {
            const currentPage = getCurrentPageNumber();
            const totalAvailable = getTotalAvailablePages();
            const pageCount = prompt(
                `当前在第${currentPage}页，最多可保存到第${totalAvailable}页\n请输入要保存的页数（从当前页开始计算）:`, 
                '1'
            );
            
            if (!pageCount || isNaN(pageCount) || parseInt(pageCount) <= 0) {
                alert('请输入有效的页数');
                return;
            }

            const requestedPages = parseInt(pageCount);
            const startPage = currentPage;
            let endPage = startPage + requestedPages - 1;
            
            // 检查是否超出可用页数
            if (endPage > totalAvailable) {
                const actualPages = totalAvailable - startPage + 1;
                if (actualPages <= 0) {
                    alert(`当前已在第${currentPage}页，没有更多页面可保存`);
                    return;
                }
                endPage = totalAvailable;
                const confirmMsg = `请求保存${requestedPages}页，但只有${actualPages}页可用（第${startPage}页到第${endPage}页）\n是否继续？`;
                if (!confirm(confirmMsg)) {
                    return;
                }
            }

            const multiPageBtn = document.querySelector('[onclick="saveMultiplePages"]') || 
                                document.evaluate('//button[text()="批量保存"]', document, null, XPathResult.FIRST_ORDERED_NODE_TYPE, null).singleNodeValue;
            
            multiPageBtn.disabled = true;
            multiPageBtn.textContent = '正在收集文章...';

            try {
                const allArticles = [];
                const allSuccessfulArticles = [];
                const allFailedArticles = [];
                const currentUrl = new URL(window.location.href);
                const searchWord = getSearchWordFromUrl();
                let actualPagesProcessed = 0;
                let actualEndPage = startPage;

                for (let page = startPage; page <= endPage; page++) {
                    multiPageBtn.textContent = `正在处理第 ${page}页 (${page - startPage + 1}/${endPage - startPage + 1})...`;
                    
                    let pageArticles = [];
                    if (page === currentPage) {
                        // 当前页面直接获取文章
                        pageArticles = collectArticleLinks();
                    } else {
                        // 其他页面需要通过请求获取
                        const pageUrl = new URL(currentUrl);
                        pageUrl.searchParams.set('page', page);
                        pageArticles = await fetchPageArticles(pageUrl.toString());
                    }
                    
                    if (pageArticles.length === 0) {
                        console.log(`第 ${page} 页没有找到文章，可能已达到最后一页`);
                        break;
                    }

                    actualPagesProcessed++;
                    actualEndPage = page;

                    // 获取每篇文章的内容
                    for (let i = 0; i < pageArticles.length; i++) {
                        multiPageBtn.textContent = `第${page}页: 处理 ${i + 1}/${pageArticles.length} (总成功:${allSuccessfulArticles.length} 总失败:${allFailedArticles.length})`;
                        const content = await fetchArticleContent(pageArticles[i].url);
                        
                        const articleWithPage = `[第${page}页] ${pageArticles[i].title}`;
                        
                        if (content && content !== '获取失败' && content !== '解析失败' && content !== '无法获取文章内容') {
                            allSuccessfulArticles.push(articleWithPage);
                        } else {
                            allFailedArticles.push(articleWithPage);
                        }
                        
                        allArticles.push({
                            ...pageArticles[i],
                            content: content,
                            pageNumber: page
                        });
                        
                        // 添加延迟避免请求过快
                        await new Promise(resolve => setTimeout(resolve, 500));
                    }
                    
                    console.log(`第 ${page} 页处理完成`);
                }

                if (allArticles.length === 0) {
                    alert('未找到任何文章');
                    return;
                }

                // 生成Word文档，使用实际处理的页面范围
                const wordHTML = createWordHTML(allArticles);
                const filename = generateFilename(searchWord, startPage, actualEndPage, actualPagesProcessed);
                
                saveAsDocx(wordHTML, filename);
                
                // 显示详细的多页保存结果
                showDetailedMultiPageResult(allSuccessfulArticles, allFailedArticles, allArticles.length, actualPagesProcessed, startPage, actualEndPage, filename);

            } catch (error) {
                console.error('多页保存时出错:', error);
                alert('多页保存失败，请查看控制台错误信息');
            } finally {
                multiPageBtn.disabled = false;
                multiPageBtn.textContent = '批量保存';
            }
        }

        // 显示详细的保存结果
        function showDetailedResult(successfulArticles, failedArticles, totalCount, filename) {
            let message = `保存完成！\n文件名：${filename}\n\n总计: ${totalCount} 篇\n成功: ${successfulArticles.length} 篇\n失败: ${failedArticles.length} 篇\n\n`;
            
            // 优先显示失败的文章
            if (failedArticles.length > 0) {
                message += '❌ 保存失败的文章:\n';
                failedArticles.forEach((title, index) => {
                    message += `${index + 1}. ${title}\n`;
                });
                message += '\n';
            }
            
            // 显示所有成功的文章
            if (successfulArticles.length > 0) {
                message += '📄 成功保存的文章:\n';
                successfulArticles.forEach((title, index) => {
                    message += `${index + 1}. ${title}\n`;
                });
            }
            
            // 如果内容太长，使用更好的显示方式
            if (message.length > 1000) {
                // 创建一个模态框来显示详细信息
                showDetailedModal(message);
            } else {
                alert(message);
            }
        }
        
        // 显示详细的多页保存结果
        function showDetailedMultiPageResult(successfulArticles, failedArticles, totalCount, actualPages, startPage, endPage, filename) {
            const pageRangeText = startPage === endPage ? `第${startPage}页` : `第${startPage}页到第${endPage}页`;
            let message = `批量保存完成！\n文件名：${filename}\n\n页面范围: ${pageRangeText} (共${actualPages}页)\n总计文章: ${totalCount} 篇\n成功: ${successfulArticles.length} 篇\n失败: ${failedArticles.length} 篇\n\n`;
            
            // 优先显示失败的文章
            if (failedArticles.length > 0) {
                message += '❌ 保存失败的文章:\n';
                failedArticles.forEach((title, index) => {
                    message += `${index + 1}. ${title}\n`;
                });
                message += '\n';
            }
            
            // 显示所有成功的文章
            if (successfulArticles.length > 0) {
                message += '📄 成功保存的文章:\n';
                successfulArticles.forEach((title, index) => {
                    message += `${index + 1}. ${title}\n`;
                });
            }
            
            // 如果内容太长，使用更好的显示方式
            if (message.length > 1000) {
                // 创建一个模态框来显示详细信息
                showDetailedModal(message);
            } else {
                alert(message);
            }
        }
        
        // 显示详细信息的模态框
        function showDetailedModal(message) {
            // 创建模态框HTML
            const modal = document.createElement('div');
            modal.style.cssText = `
                position: fixed;
                top: 0;
                left: 0;
                width: 100%;
                height: 100%;
                background-color: rgba(0, 0, 0, 0.5);
                z-index: 10000;
                display: flex;
                justify-content: center;
                align-items: center;
            `;
            
            const modalContent = document.createElement('div');
            modalContent.style.cssText = `
                background: white;
                padding: 20px;
                border-radius: 8px;
                max-width: 600px;
                max-height: 80%;
                overflow-y: auto;
                box-shadow: 0 4px 20px rgba(0, 0, 0, 0.3);
                position: relative;
            `;
            
            const closeButton = document.createElement('button');
            closeButton.textContent = '关闭';
            closeButton.style.cssText = `
                position: absolute;
                top: 10px;
                right: 15px;
                background: #f44336;
                color: white;
                border: none;
                padding: 5px 10px;
                border-radius: 4px;
                cursor: pointer;
            `;
            closeButton.onclick = () => document.body.removeChild(modal);
            
            const messageElement = document.createElement('pre');
            messageElement.textContent = message;
            messageElement.style.cssText = `
                white-space: pre-wrap;
                word-wrap: break-word;
                font-family: '微软雅黑', sans-serif;
                font-size: 14px;
                line-height: 1.5;
                margin: 0;
                padding-right: 60px;
            `;
            
            modalContent.appendChild(closeButton);
            modalContent.appendChild(messageElement);
            modal.appendChild(modalContent);
            document.body.appendChild(modal);
            
            // 点击背景关闭模态框
            modal.onclick = (e) => {
                if (e.target === modal) {
                    document.body.removeChild(modal);
                }
            };
        }

        // 从URL中提取搜索词
        function getSearchWordFromUrl() {
            const urlParams = new URLSearchParams(window.location.search);
            const searchWord = urlParams.get('searchword');
            return searchWord ? decodeURIComponent(searchWord) : null;
        }

        // 从URL或页面中获取当前页码
        function getCurrentPageNumber() {
            // 首先从URL参数获取
            const urlParams = new URLSearchParams(window.location.search);
            const pageParam = urlParams.get('page');
            if (pageParam && !isNaN(parseInt(pageParam))) {
                return parseInt(pageParam);
            }

            // 如果URL没有页码参数，从分页组件获取
            const paginationElements = document.querySelectorAll('.ui-paging .ui-paging-current, .page-current, .current');
            if (paginationElements.length > 0) {
                const currentPageText = paginationElements[0].textContent.trim();
                const currentPage = parseInt(currentPageText);
                if (!isNaN(currentPage)) {
                    return currentPage;
                }
            }

            // 默认返回第1页
            return 1;
        }

        // 检测总可用页数
        function getTotalAvailablePages() {
            console.log('开始检测总页数...');
            let maxPage = 1;

            // 查找包含分页的td.t4元素，使用更具体的选择器
            // 首先尝试找到包含页码链接的td.t4
            let paginationContainer = null;
            const t4Elements = document.querySelectorAll('td.t4');
            
            // 遍历所有td.t4元素，找到包含页码链接的那个
            for (const element of t4Elements) {
                // 检查是否包含page=参数的链接或者数字页码
                const hasPageLinks = element.querySelector('a[href*="page="]') || 
                                   Array.from(element.querySelectorAll('a')).some(a => /^\d+$/.test(a.textContent.trim()));
                if (hasPageLinks) {
                    paginationContainer = element;
                    console.log('找到分页容器:', paginationContainer.innerHTML.substring(0, 200) + '...');
                    break;
                }
            }
            
            if (paginationContainer) {
                
                // 首先查找"尾页"链接，这通常包含最大页码
                const lastPageLink = paginationContainer.querySelector('a.last-page, a[href*="page="]:last-of-type');
                if (lastPageLink) {
                    const href = lastPageLink.getAttribute('href');
                    if (href) {
                        const pageMatch = href.match(/page=(\d+)/);
                        if (pageMatch) {
                            const lastPage = parseInt(pageMatch[1]);
                            console.log('从"尾页"链接检测到最大页码:', lastPage);
                            if (!isNaN(lastPage)) {
                                maxPage = Math.max(maxPage, lastPage);
                            }
                        }
                    }
                }
                
                // 如果没有找到"尾页"链接，检查所有包含"尾页"文本的链接
                const tailPageLinks = Array.from(paginationContainer.querySelectorAll('a')).filter(a => a.textContent.includes('尾页'));
                tailPageLinks.forEach(link => {
                    const href = link.getAttribute('href');
                    if (href) {
                        const pageMatch = href.match(/page=(\d+)/);
                        if (pageMatch) {
                            const lastPage = parseInt(pageMatch[1]);
                            console.log('从"尾页"文本链接检测到最大页码:', lastPage);
                            if (!isNaN(lastPage)) {
                                maxPage = Math.max(maxPage, lastPage);
                            }
                        }
                    }
                });
                
                // 查找所有数字页码链接
                const pageLinks = paginationContainer.querySelectorAll('a[href*="page="]');
                pageLinks.forEach(element => {
                    const pageText = element.textContent.trim();
                    const pageNum = parseInt(pageText);
                    if (!isNaN(pageNum)) {
                        console.log('从页码链接检测到页码:', pageNum);
                        maxPage = Math.max(maxPage, pageNum);
                    }
                    
                    // 同时检查链接URL中的页码
                    const href = element.getAttribute('href');
                    if (href) {
                        const pageMatch = href.match(/page=(\d+)/);
                        if (pageMatch) {
                            const urlPage = parseInt(pageMatch[1]);
                            if (!isNaN(urlPage)) {
                                console.log('从链接URL检测到页码:', urlPage);
                                maxPage = Math.max(maxPage, urlPage);
                            }
                        }
                    }
                });
            } else {
                console.log('未找到包含页码链接的分页容器 td.t4，尝试其他选择器...');
                
                // 备用选择器：查找其他可能的分页组件
                const fallbackSelectors = [
                    '.paging', 
                    '.pagination', 
                    '.page-nav',
                    'td[align="center"]', // 分页容器通常居中对齐
                    '.ui-paging a', 
                    '.page-link', 
                    '.pagination a', 
                    '.page a', 
                    'a[href*="page="]'
                ];
                
                fallbackSelectors.forEach(selector => {
                    if (paginationContainer) return; // 如果已找到就跳过
                    
                    const elements = document.querySelectorAll(selector);
                    if (elements.length > 0) {
                        console.log(`尝试备用选择器 ${selector} 找到 ${elements.length} 个元素`);
                        
                        // 对于容器类选择器，检查是否包含页码链接
                        if (selector.includes('paging') || selector.includes('pagination') || selector.includes('page-nav') || selector.includes('td[align')) {
                            for (const element of elements) {
                                const hasPageLinks = element.querySelector('a[href*="page="]') || 
                                                   Array.from(element.querySelectorAll('a')).some(a => /^\d+$/.test(a.textContent.trim()));
                                if (hasPageLinks) {
                                    paginationContainer = element;
                                    console.log(`使用备用选择器 ${selector} 找到分页容器:`, element.innerHTML.substring(0, 200) + '...');
                                    break;
                                }
                            }
                        } else {
                            // 对于链接选择器，直接处理页码
                            elements.forEach(element => {
                                const pageText = element.textContent.trim();
                                const pageNum = parseInt(pageText);
                                if (!isNaN(pageNum) && pageNum > maxPage) {
                                    maxPage = pageNum;
                                    console.log('从备用选择器检测到页码:', pageNum);
                                }
                            });
                        }
                    }
                });
                
                // 如果通过备用选择器找到了容器，处理其中的页码
                if (paginationContainer) {
                    // 重新处理找到的容器中的页码
                    const pageLinks = paginationContainer.querySelectorAll('a[href*="page="]');
                    pageLinks.forEach(element => {
                        const pageText = element.textContent.trim();
                        const pageNum = parseInt(pageText);
                        if (!isNaN(pageNum)) {
                            console.log('从备用容器中的页码链接检测到页码:', pageNum);
                            maxPage = Math.max(maxPage, pageNum);
                        }
                        
                        // 同时检查链接URL中的页码
                        const href = element.getAttribute('href');
                        if (href) {
                            const pageMatch = href.match(/page=(\d+)/);
                            if (pageMatch) {
                                const urlPage = parseInt(pageMatch[1]);
                                if (!isNaN(urlPage)) {
                                    console.log('从备用容器链接URL检测到页码:', urlPage);
                                    maxPage = Math.max(maxPage, urlPage);
                                }
                            }
                        }
                    });
                }
            }

            // 查找"共X页"或类似的文本
            const totalPageText = document.body.textContent.match(/共\s*(\d+)\s*页/);
            if (totalPageText) {
                const totalFromText = parseInt(totalPageText[1]);
                console.log('从"共X页"文本检测到总页数:', totalFromText);
                if (totalFromText > maxPage) {
                    maxPage = totalFromText;
                }
            }

            console.log('最终检测到的总页数:', maxPage);
            return maxPage;
        }

        // 生成文件名（包含页码范围）
        function generateFilename(searchWord, startPage, endPage, actualPages) {
            const now = new Date();
            const dateStr = now.toISOString().slice(0, 10);
            const timeStr = now.toTimeString().slice(0, 8).replace(/:/g, '-');
            const baseFilename = searchWord || 'SCOPSR文章集';

            if (startPage === endPage || actualPages === 1) {
                // 单页
                return `${baseFilename}_第${startPage}页_${dateStr}_${timeStr}.docx`;
            } else {
                // 多页范围
                return `${baseFilename}_第${startPage}页到第${endPage}页_${dateStr}_${timeStr}.docx`;
            }
        }

        // 获取指定页面的文章列表
        function fetchPageArticles(pageUrl) {
            return new Promise((resolve) => {
                GM_xmlhttpRequest({
                    method: 'GET',
                    url: pageUrl,
                    onload: function(response) {
                        try {
                            const parser = new DOMParser();
                            const doc = parser.parseFromString(response.responseText, 'text/html');
                            const articles = [];
                            
                            // 获取搜索结果容器中的所有文章
                            const container = doc.querySelector('#searchresult, .jiansuo-container-result');
                            if (!container) {
                                console.log(`页面 ${pageUrl} 未找到搜索结果容器`);
                                resolve(articles);
                                return;
                            }
                            
                            // 每个ul包含一篇文章
                            const articleElements = container.querySelectorAll('ul');
                            
                            articleElements.forEach((ul, index) => {
                                // 获取标题和链接
                                const titleLink = ul.querySelector('li h2 a');
                                if (!titleLink) {
                                    return;
                                }
                                
                                // 清理标题，移除HTML标签
                                const title = titleLink.textContent.replace(/<[^>]*>/g, '').trim();
                                const url = titleLink.href;
                                
                                // 获取日期
                                const dateSpan = ul.querySelector('li div.jiansuo-result-link span');
                                const date = dateSpan ? dateSpan.textContent.trim() : '';
                                
                                // 获取摘要
                                const summary = ul.querySelector('li p');
                                const summaryText = summary ? summary.innerHTML.replace(/<[^>]*>/g, '').trim() : '';
                                
                                if (url && titleLink.href) {
                                    articles.push({
                                        title: title,
                                        url: url,
                                        date: date,
                                        summary: summaryText
                                    });
                                }
                            });

                            console.log(`从页面 ${pageUrl} 收集到 ${articles.length} 篇文章`);
                            resolve(articles);
                        } catch (error) {
                            console.error('解析页面失败:', pageUrl, error);
                            resolve([]);
                        }
                    },
                    onerror: function(error) {
                        console.error('获取页面失败:', pageUrl, error);
                        resolve([]);
                    }
                });
            });
        }

        function collectArticleLinks() {
            const articles = [];
            
            // 获取搜索结果容器中的所有文章
            const container = document.querySelector('#searchresult, .jiansuo-container-result');
            if (!container) {
                console.log('未找到搜索结果容器');
                return articles;
            }
            
            // 每个ul包含一篇文章
            const articleElements = container.querySelectorAll('ul');
            console.log(`找到 ${articleElements.length} 个ul元素`);
            
            articleElements.forEach((ul, index) => {
                // 获取标题和链接
                const titleLink = ul.querySelector('li h2 a');
                if (!titleLink) {
                    console.log(`第 ${index + 1} 个ul没有找到标题链接`);
                    return;
                }
                
                // 清理标题，移除HTML标签
                const title = titleLink.textContent.replace(/<[^>]*>/g, '').trim();
                const url = titleLink.href;
                
                // 获取日期
                const dateSpan = ul.querySelector('li div.jiansuo-result-link span');
                const date = dateSpan ? dateSpan.textContent.trim() : '';
                
                // 获取摘要
                const summary = ul.querySelector('li p');
                const summaryText = summary ? summary.innerHTML.replace(/<[^>]*>/g, '').trim() : '';
                
                // 不再限制URL必须包含/xwzx/，因为有些文章可能在其他栏目
                if (url && titleLink.href) {
                    articles.push({
                        title: title,
                        url: url,
                        date: date,
                        summary: summaryText
                    });
                    console.log(`收集第 ${articles.length} 篇文章: ${title}`);
                }
            });

            console.log(`总共收集到 ${articles.length} 篇文章`);
            return articles;
        }

        function fetchArticleContent(url) {
            return new Promise((resolve) => {
                GM_xmlhttpRequest({
                    method: 'GET',
                    url: url,
                    onload: function(response) {
                        try {
                            const parser = new DOMParser();
                            const doc = parser.parseFromString(response.responseText, 'text/html');
                            
                            // 主要选择器：文章内容在id="Zoom"的元素中
                            const zoomElement = doc.querySelector('#Zoom');
                            if (zoomElement) {
                                resolve(zoomElement.textContent.trim());
                                return;
                            }
                            
                            // 备用选择器
                            const possibleSelectors = [
                                'td#Zoom',
                                '.TRS_Editor',
                                '.Custom_UnionStyle',
                                '.hui12#Zoom',
                                'td.hui12[id="Zoom"]'
                            ];
                            
                            for (const selector of possibleSelectors) {
                                const element = doc.querySelector(selector);
                                if (element && element.textContent.length > 100) {
                                    resolve(element.textContent.trim());
                                    return;
                                }
                            }
                            
                            resolve('无法获取文章内容');
                        } catch (error) {
                            console.error('解析文章内容失败:', url, error);
                            resolve('解析失败');
                        }
                    },
                    onerror: function(error) {
                        console.error('获取文章内容失败:', url, error);
                        resolve('获取失败');
                    }
                });
            });
        }
    }

    if (isArticlePage) {
        // 文章详情页面：添加单独保存按钮
        const saveBtn = document.createElement('button');
        saveBtn.textContent = '保存为DOCX文档';
        saveBtn.className = 'save-articles-btn';
        saveBtn.onclick = saveCurrentArticle;
        document.body.appendChild(saveBtn);

        function saveCurrentArticle() {
            // 使用正确的选择器获取文章内容
            const zoomElement = document.querySelector('#Zoom');
            let content = '';
            let title = '';
            let date = '';
            
            if (zoomElement) {
                content = zoomElement.textContent.trim();
            } else {
                // 备用选择器
                const possibleSelectors = [
                    'td#Zoom',
                    '.TRS_Editor',
                    '.Custom_UnionStyle',
                    '.hui12#Zoom'
                ];
                for (const selector of possibleSelectors) {
                    const element = document.querySelector(selector);
                    if (element && element.textContent.length > 100) {
                        content = element.textContent.trim();
                        break;
                    }
                }
            }

            if (!content) {
                alert('无法找到文章内容');
                return;
            }
            
            // 获取文章标题
            const titleElement = document.querySelector('.hui14c');
            title = titleElement ? titleElement.textContent.trim() : document.title;
            
            // 获取发布时间
            const timeElement = document.querySelector('.hui14');
            if (timeElement) {
                const timeMatch = timeElement.textContent.match(/时间：(\d{4}-\d{2}-\d{2})/);
                date = timeMatch ? timeMatch[1] : new Date().toISOString().slice(0, 10);
            } else {
                date = new Date().toISOString().slice(0, 10);
            }

            const article = {
                title: title,
                url: window.location.href,
                date: date,
                content: content
            };

            // 生成单篇文章的Word文档
            const wordHTML = createWordHTML([article]);
            // 文件名包含日期和时间（精确到秒）
            const now = new Date();
            const timeStr = now.toTimeString().slice(0, 8).replace(/:/g, '-');
            const filename = `${title}_${date}_${timeStr}.docx`;
            
            saveAsDocx(wordHTML, filename);
            
            alert(`文章已保存为DOCX文档！\n文件名：${filename}`);
        }
    }
})();