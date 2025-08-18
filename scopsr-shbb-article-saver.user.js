// ==UserScript==
// @name         SCOPSR & SHBB 文章保存器
// @namespace    http://tampermonkey.net/
// @version      5.0-site-identification
// @description  一键保存SCOPSR和SHBB网站的文章为DOCX格式，支持多页批量保存，保留原始段落结构、标题、列表和空段落格式
// @author       You
// @match        https://www.scopsr.gov.cn/was5/web/search*
// @match        https://www.scopsr.gov.cn/xwzx/*
// @match        http://www.scopsr.gov.cn/was5/web/search*
// @match        http://www.scopsr.gov.cn/xwzx/*
// @match        https://www.shbb.gov.cn/search*
// @match        https://www.shbb.gov.cn/*/*.jhtml
// @match        https://www.shbb.gov.cn/*/*/*.jhtml
// @match        https://www.shbb.gov.cn/*/*/*/*.jhtml
// @match        http://www.shbb.gov.cn/search*
// @match        http://www.shbb.gov.cn/*/*.jhtml
// @match        http://www.shbb.gov.cn/*/*/*.jhtml
// @match        http://www.shbb.gov.cn/*/*/*/*.jhtml
// @grant        GM_download
// @grant        GM_addStyle
// @grant        GM_xmlhttpRequest
// ==/UserScript==

(function() {
    'use strict';

    // Debug logging function
    function debugLog(message, data = null) {
        console.log(`[SCOPSR-SHBB Script] ${message}`, data || '');
    }

    // Visual indicator for script loading
    function showScriptLoadedIndicator() {
        const indicator = document.createElement('div');
        indicator.style.cssText = `
            position: fixed;
            top: 10px;
            right: 10px;
            z-index: 10000;
            background: #4CAF50;
            color: white;
            padding: 8px 16px;
            border-radius: 4px;
            font-size: 12px;
            box-shadow: 0 2px 8px rgba(0,0,0,0.2);
            opacity: 0;
            transition: opacity 0.3s ease;
        `;
        indicator.textContent = 'SCOPSR/SHBB脚本已加载';
        document.body.appendChild(indicator);
        
        // Fade in
        setTimeout(() => indicator.style.opacity = '1', 100);
        
        // Fade out after 3 seconds
        setTimeout(() => {
            indicator.style.opacity = '0';
            setTimeout(() => {
                if (indicator.parentNode) {
                    indicator.parentNode.removeChild(indicator);
                }
            }, 300);
        }, 3000);
    }

    // Test function to verify selectors
    function testSelectors() {
        debugLog('开始测试选择器...');
        
        const site = detectSite();
        if (!site) {
            debugLog('无法检测到支持的网站');
            return false;
        }
        
        debugLog('当前网站配置:', site.domain);
        
        // Test search page selectors
        if (site.searchPagePattern.test(window.location.pathname)) {
            debugLog('当前是搜索页面，测试搜索页面选择器...');
            
            const container = document.querySelector(site.selectors.searchContainer);
            debugLog('搜索容器元素:', container ? '找到' : '未找到');
            
            if (container) {
                const items = container.querySelectorAll(site.selectors.articleItems);
                debugLog(`文章项目数量: ${items.length}`);
                
                if (items.length > 0) {
                    const firstItem = items[0];
                    const titleLink = firstItem.querySelector(site.selectors.titleLink);
                    const dateElement = firstItem.querySelector(site.selectors.dateElement);
                    
                    debugLog('第一个文章项目:', {
                        '标题链接': titleLink ? titleLink.textContent.trim() : '未找到',
                        '日期元素': dateElement ? dateElement.textContent.trim() : '未找到',
                        'HTML结构': firstItem.innerHTML.substring(0, 200) + '...'
                    });
                }
                
                const pagination = document.querySelector(site.selectors.paginationContainer);
                debugLog('分页容器:', pagination ? '找到' : '未找到');
            }
        }
        
        // Test article page selectors
        if (site.articlePagePattern.test(window.location.pathname)) {
            debugLog('当前是文章页面，测试文章页面选择器...');
            
            const content = document.querySelector(site.selectors.articleContent);
            debugLog('文章内容元素:', content ? `找到，长度: ${content.textContent.length}` : '未找到');
            
            if (!content) {
                debugLog('尝试备用选择器...');
                for (const selector of site.selectors.articleContentBackup) {
                    const backupElement = document.querySelector(selector);
                    if (backupElement) {
                        debugLog(`备用选择器 "${selector}" 找到，长度: ${backupElement.textContent.length}`);
                        break;
                    }
                }
            }
            
            const title = document.querySelector(site.selectors.articleTitle);
            const time = document.querySelector(site.selectors.articleTime);
            
            debugLog('文章标题元素:', title ? title.textContent.trim() : '未找到');
            debugLog('文章时间元素:', time ? time.textContent.trim() : '未找到');
        }
        
        return true;
    }

    debugLog('脚本开始执行...');

    // 网站配置
    const SITE_CONFIGS = {
        SCOPSR: {
            domain: 'scopsr.gov.cn',
            searchPagePattern: /\/was5\/web\/search/,
            articlePagePattern: /\/xwzx\//,
            selectors: {
                searchContainer: '#searchresult, .jiansuo-container-result',
                articleItems: 'ul',
                titleLink: 'li h2 a',
                dateElement: 'li div.jiansuo-result-link span',
                summaryElement: 'li p',
                paginationContainer: 'td.t4',
                pageLinks: 'a[href*="page="]',
                articleContent: '#Zoom',
                articleContentBackup: ['td#Zoom', '.TRS_Editor', '.Custom_UnionStyle', '.hui12#Zoom', 'td.hui12[id="Zoom"]'],
                articleTitle: '.hui14c',
                articleTime: '.hui14'
            },
            pagePattern: 'page',
            pagination: {
                getPageFromUrl: (url) => {
                    const urlParams = new URLSearchParams(new URL(url).search);
                    return parseInt(urlParams.get('page')) || 1;
                },
                buildPageUrl: (baseUrl, page) => {
                    const url = new URL(baseUrl);
                    url.searchParams.set('page', page);
                    return url.toString();
                }
            }
        },
        SHBB: {
            domain: 'shbb.gov.cn',
            searchPagePattern: /\/search/,
            articlePagePattern: /\.jhtml$/,
            selectors: {
                searchContainer: '#conView .jiansuoResult',
                articleItems: 'tr',
                titleLink: 'th a',
                dateElement: 'td:last-child',
                summaryElement: 'td[colspan="2"]',
                paginationContainer: '#conView .page',
                pageLinks: 'a.Num',
                articleContent: '#conView .cvbody',
                articleContentBackup: ['#conView blockquote .cvbody', '#conView .con .cvbody'],
                articleTitle: '#conView h2',
                articleTime: '#conView .details label'
            },
            pagePattern: 'search',
            pagination: {
                getPageFromUrl: (url) => {
                    const match = url.match(/search_(\d+)\.jspx/);
                    return match ? parseInt(match[1]) : 1;
                },
                buildPageUrl: (baseUrl, page) => {
                    const url = new URL(baseUrl);
                    const searchParams = url.searchParams;
                    const q = searchParams.get('q') || '';
                    
                    if (page === 1) {
                        return `${url.origin}/search.jspx?q=${encodeURIComponent(q)}`;
                    } else {
                        return `${url.origin}/search_${page}.jspx?q=${encodeURIComponent(q)}`;
                    }
                }
            }
        }
    };

    // 检测当前网站
    function detectSite() {
        const hostname = window.location.hostname;
        if (hostname.includes('scopsr.gov.cn')) {
            return SITE_CONFIGS.SCOPSR;
        } else if (hostname.includes('shbb.gov.cn')) {
            return SITE_CONFIGS.SHBB;
        }
        return null;
    }

    const currentSite = detectSite();
    if (!currentSite) {
        debugLog('不支持的网站:', window.location.hostname);
        return;
    }

    debugLog('检测到网站:', currentSite.domain);
    debugLog('当前URL:', window.location.href);
    
    // 测试URL模式匹配
    if (currentSite.domain === 'shbb.gov.cn') {
        const testUrls = [
            'https://www.shbb.gov.cn/bzglyj202404/9601.jhtml',
            'https://www.shbb.gov.cn/search.jspx?q=test',
            'https://www.shbb.gov.cn/category/123.jhtml'
        ];
        testUrls.forEach(url => {
            const pathname = new URL(url).pathname;
            debugLog(`URL测试: ${pathname}`, {
                '搜索页面': currentSite.searchPagePattern.test(pathname),
                '文章页面': currentSite.articlePagePattern.test(pathname)
            });
        });
    }
    
    // Show script loaded indicator
    showScriptLoadedIndicator();
    
    // Test selectors after page load
    setTimeout(() => {
        testSelectors();
    }, 1000);

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
    const isSearchPage = currentSite.searchPagePattern.test(window.location.pathname);
    const isArticlePage = currentSite.articlePagePattern.test(window.location.pathname);
    
    debugLog('页面类型检测:', {
        '搜索页面': isSearchPage,
        '文章页面': isArticlePage,
        '当前路径': window.location.pathname,
        '搜索页面正则': currentSite.searchPagePattern.toString(),
        '文章页面正则': currentSite.articlePagePattern.toString()
    });

    // 创建Word兼容的HTML文档
    function createWordHTML(articles) {
        const html = `<!DOCTYPE html>
<html xmlns:o='urn:schemas-microsoft-com:office:office' 
      xmlns:w='urn:schemas-microsoft-com:office:word' 
      xmlns='http://www.w3.org/TR/REC-html40'>
<head>
    <meta charset='utf-8'>
    <title>${currentSite.domain}文章集</title>
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
            line-height: 1.8;
            margin: 12pt 0;
            padding: 0;
            font-weight: normal;
        }
        
        .article-content p br {
            mso-data-placement: same-cell;
        }
        
        .article-content p[style*="text-align: center"] {
            text-align: center;
            font-weight: bold;
            margin: 18pt 0 12pt 0;
            text-indent: 0;
            font-size: 14pt;
        }
        
        .article-content p[style*="text-indent: 2em"] {
            text-indent: 2em;
            margin: 12pt 0;
            text-align: justify;
            line-height: 1.8;
            font-weight: normal;
        }
        
        .article-content p[style*="text-indent: 0"] {
            text-indent: 0;
            margin: 12pt 0;
            text-align: justify;
            line-height: 1.8;
            font-weight: normal;
        }
        
        .article-content p:empty,
        .article-content p[style*="&nbsp;"] {
            margin: 12pt 0;
            line-height: 1;
            height: 12pt;
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
        <h1>${currentSite.domain}文章集</h1>
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
        
        // 按双换行分割段落，保留段落结构
        const paragraphs = content.split('\n\n');
        
        return paragraphs.map((paragraph, index) => {
            const trimmedParagraph = paragraph.trim();
            
            // 如果是空段落，返回空段落用于分隔
            if (!trimmedParagraph) {
                return '<p style="margin: 12pt 0;">&nbsp;</p>';
            }
            
            // 处理单行内的换行（保持在同一段落内）
            const lines = trimmedParagraph.split('\n');
            const processedLines = lines.map(line => line.trim()).filter(line => line.length > 0);
            
            if (processedLines.length === 0) {
                return '<p style="margin: 12pt 0;">&nbsp;</p>';
            }
            
            const firstLine = processedLines[0];
            
            // NOTE: Heading detection is now handled in extractContentWithFormatting
            // This function now only processes pre-formatted structured content
            
            // 检测是否为列表项（数字编号或括号编号）
            const isListItem = /^\d+\s*[\.、]\s*/.test(firstLine) || 
                              /^（[一二三四五六七八九十\d]+）/.test(firstLine) ||
                              /^[一二三四五六七八九十]+[\.、]\s*/.test(firstLine);
            
            // 如果段落内有多行，使用<br>标签连接
            const escapedContent = processedLines.length > 1 
                ? processedLines.map(line => escapeHtml(line)).join('<br/>') 
                : escapeHtml(firstLine);
            
            // 应用样式（标题检测已在HTML层面完成）
            if (isListItem) {
                return `<p style="margin: 12pt 0; text-indent: 0; text-align: justify; font-weight: normal;">${escapedContent}</p>`;
            } else {
                // 普通段落，确保有明确的段落间距且不加粗
                return `<p style="margin: 12pt 0; text-indent: 2em; text-align: justify; line-height: 1.8; font-weight: normal;">${escapedContent}</p>`;
            }
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

    // 收集当前页面的文章链接（增强调试）
    function collectArticleLinks() {
        const articles = [];
        debugLog('开始收集当前页面文章链接...');
        
        // 获取搜索结果容器
        const container = document.querySelector(currentSite.selectors.searchContainer);
        if (!container) {
            debugLog('未找到搜索结果容器，使用的选择器:', currentSite.selectors.searchContainer);
            debugLog('页面HTML结构预览:', document.body.innerHTML.substring(0, 500) + '...');
            return articles;
        }
        
        debugLog('找到搜索结果容器');
        
        // 获取文章元素
        const articleElements = container.querySelectorAll(currentSite.selectors.articleItems);
        debugLog(`找到 ${articleElements.length} 个文章元素，使用选择器: ${currentSite.selectors.articleItems}`);
        
        if (articleElements.length === 0) {
            debugLog('容器HTML结构:', container.innerHTML.substring(0, 500) + '...');
        }
        
        articleElements.forEach((element, index) => {
            // 跳过表头行（针对SHBB网站）
            if (currentSite.domain === 'shbb.gov.cn' && element.querySelector('th') && !element.querySelector('th a')) {
                return;
            }
            
            // 获取标题和链接
            const titleLink = element.querySelector(currentSite.selectors.titleLink);
            if (!titleLink) {
                console.log(`第 ${index + 1} 个元素没有找到标题链接`);
                return;
            }
            
            // 清理标题，移除HTML标签
            const title = titleLink.textContent.replace(/<[^>]*>/g, '').trim();
            const url = titleLink.href;
            
            if (!title || !url) {
                return;
            }
            
            // 获取日期
            let date = '';
            if (currentSite.domain === 'scopsr.gov.cn') {
                const dateSpan = element.querySelector(currentSite.selectors.dateElement);
                date = dateSpan ? dateSpan.textContent.trim() : '';
            } else if (currentSite.domain === 'shbb.gov.cn') {
                // SHBB网站的日期在同一行的最后一个td
                const dateTd = element.querySelector(currentSite.selectors.dateElement);
                date = dateTd ? dateTd.textContent.trim() : '';
            }
            
            // 获取摘要
            let summaryText = '';
            if (currentSite.domain === 'scopsr.gov.cn') {
                const summary = element.querySelector(currentSite.selectors.summaryElement);
                summaryText = summary ? summary.innerHTML.replace(/<[^>]*>/g, '').trim() : '';
            } else if (currentSite.domain === 'shbb.gov.cn') {
                // SHBB网站的摘要在下一行的合并单元格中
                const nextRow = element.nextElementSibling;
                if (nextRow) {
                    const summaryTd = nextRow.querySelector(currentSite.selectors.summaryElement);
                    summaryText = summaryTd ? summaryTd.textContent.trim() : '';
                }
            }
            
            articles.push({
                title: title,
                url: url,
                date: date,
                summary: summaryText
            });
            debugLog(`收集第 ${articles.length} 篇文章:`, { title, url, date });
        });

        debugLog(`总共收集到 ${articles.length} 篇文章`);
        return articles;
    }

    // 获取文章内容（增强错误处理和调试）
    function fetchArticleContent(url) {
        return new Promise((resolve) => {
            debugLog(`开始获取文章内容: ${url}`);
            
            GM_xmlhttpRequest({
                method: 'GET',
                url: url,
                timeout: 30000, // 30秒超时
                onload: function(response) {
                    try {
                        debugLog(`文章请求成功，状态码: ${response.status}`);
                        
                        if (response.status !== 200) {
                            debugLog(`HTTP错误状态码: ${response.status}`);
                            resolve(`HTTP错误: ${response.status}`);
                            return;
                        }
                        
                        const parser = new DOMParser();
                        const doc = parser.parseFromString(response.responseText, 'text/html');
                        
                        // 主要选择器
                        const contentElement = doc.querySelector(currentSite.selectors.articleContent);
                        if (contentElement) {
                            const content = extractContentWithFormatting(contentElement);
                            debugLog(`主选择器成功，内容长度: ${content.length}`);
                            resolve(content);
                            return;
                        }
                        
                        debugLog('主选择器未找到内容，尝试备用选择器...');
                        
                        // 备用选择器
                        for (let i = 0; i < currentSite.selectors.articleContentBackup.length; i++) {
                            const selector = currentSite.selectors.articleContentBackup[i];
                            const element = doc.querySelector(selector);
                            if (element && element.textContent.length > 100) {
                                const content = extractContentWithFormatting(element);
                                debugLog(`备用选择器${i + 1} "${selector}" 成功，内容长度: ${content.length}`);
                                resolve(content);
                                return;
                            }
                        }
                        
                        debugLog('所有选择器都未找到有效内容');
                        resolve('无法获取文章内容');
                    } catch (error) {
                        debugLog('解析文章内容失败:', error.message);
                        resolve('解析失败');
                    }
                },
                onerror: function(error) {
                    debugLog('网络请求失败:', error);
                    resolve('网络请求失败');
                },
                ontimeout: function() {
                    debugLog('请求超时:', url);
                    resolve('请求超时');
                }
            });
        });
    }

    // 提取内容并保留格式
    // 增强版段落处理：区分标题、列表、普通段落和空段落，保留原有的段落结构和换行
    function extractContentWithFormatting(element) {
        // 通用函数：提取元素内容并保留换行
        function extractTextWithLineBreaks(elem, preserveVisualLayout = false) {
            // 首先处理<br>标签，将其替换为换行符
            const clonedElem = elem.cloneNode(true);
            const brElements = clonedElem.querySelectorAll('br');
            brElements.forEach(br => {
                br.replaceWith('\n');
            });
            
            // 获取文本内容，这时<br>已经被替换为换行符
            let text = clonedElem.textContent || clonedElem.innerText || '';
            
            // 如果是SHBB网站且需要保持视觉布局，特殊处理
            if (preserveVisualLayout && currentSite.domain === 'shbb.gov.cn') {
                // 对于SHBB网站，模拟CSS white-space: normal的效果
                // 1. 将换行符替换为空格（除非是段落之间的换行）
                text = text.replace(/([^\n])\n([^\n])/g, '$1 $2');
                // 2. 合并多个空格为单个空格
                text = text.replace(/[ \t]+/g, ' ');
                // 3. 保留段落之间的换行（连续两个换行）
                text = text.replace(/\n\n+/g, '\n\n');
                // 4. 去掉行首行尾的空格
                text = text.split('\n').map(line => line.trim()).join('\n');
            } else {
                // 原有的清理逻辑
                text = text.replace(/[ \t]+/g, ' '); // 将多个空格和制表符替换为单个空格
                text = text.replace(/\n[ \t]+/g, '\n'); // 去掉换行后的空格
                text = text.replace(/[ \t]+\n/g, '\n'); // 去掉换行前的空格
                text = text.replace(/\n{3,}/g, '\n\n'); // 将3个或更多换行替换为2个
            }
            
            return text.trim();
        }
        
        // 通用段落处理函数（适用于SHBB和SCOPSR）
        function processParagraphStructure(containerElement, siteName) {
            const paragraphs = containerElement.querySelectorAll('p');
            if (paragraphs.length > 0) {
                const processedParagraphs = [];
                
                paragraphs.forEach((p, index) => {
                    // 使用新的文本提取函数，对SHBB网站使用视觉保留模式
                    const preserveVisual = (siteName === 'shbb.gov.cn');
                    const textContent = extractTextWithLineBreaks(p, preserveVisual);
                    const style = p.getAttribute('style') || '';
                    
                    // 处理空段落（只包含&nbsp;或空白）
                    if (!textContent || textContent === '\u00A0' || /^\s*$/.test(textContent)) {
                        // 保留空段落作为段落分隔
                        processedParagraphs.push({ type: 'empty', content: '' });
                        return;
                    }
                    
                    // 检测是否为标题段落 - 基于HTML结构和CSS样式，而不是文本内容
                    let isHeading = false;
                    
                    // 检查HTML标签是否为标题标签
                    if (p.tagName && /^H[1-6]$/.test(p.tagName)) {
                        isHeading = true;
                        debugLog(`检测到HTML标题标签 ${p.tagName}:`, textContent.substring(0, 50));
                    }
                    
                    if (!isHeading && siteName === 'SHBB') {
                        // 对于SHBB，检查font-family是否为黑体（标题字体）
                        let hasHeadingFont = false;
                        
                        // 检查直接style属性
                        if (style.includes('font-family')) {
                            const fontMatch = style.match(/font-family\s*:\s*([^;]+)/);
                            if (fontMatch && fontMatch[1].includes('黑体')) {
                                hasHeadingFont = true;
                            }
                        }
                        
                        // 检查内部span元素的font-family
                        if (!hasHeadingFont) {
                            const spans = p.querySelectorAll('span');
                            for (const span of spans) {
                                const spanStyle = span.getAttribute('style') || '';
                                if (spanStyle.includes('font-family')) {
                                    const spanFontMatch = spanStyle.match(/font-family\s*:\s*([^;]+)/);
                                    if (spanFontMatch && spanFontMatch[1].includes('黑体')) {
                                        hasHeadingFont = true;
                                        break;
                                    }
                                }
                            }
                        }
                        
                        // 检查是否居中对齐
                        const isCentered = style.includes('text-align:center') || style.includes('text-align: center');
                        
                        // SHBB标题标准：黑体字体或居中对齐
                        if (hasHeadingFont || isCentered) {
                            isHeading = true;
                            debugLog(`SHBB检测到标题段落 (黑体:${hasHeadingFont}, 居中:${isCentered}):`, textContent.substring(0, 50));
                        }
                    } else if (!isHeading && siteName === 'SCOPSR') {
                        // 对于SCOPSR，检查是否居中对齐
                        const isCentered = style.includes('text-align:center') || style.includes('text-align: center');
                        if (isCentered) {
                            isHeading = true;
                            debugLog(`SCOPSR检测到居中标题段落:`, textContent.substring(0, 50));
                        }
                    }
                    
                    // 检测是否为有缩进的段落
                    const hasIndent = style.includes('text-indent') && !style.includes('text-indent:0');
                    
                    // 检测字体样式
                    let fontFamily = '';
                    const fontMatch = style.match(/font-family\s*:\s*([^;]+)/);
                    if (fontMatch) {
                        fontFamily = fontMatch[1].trim();
                    }
                    
                    // 检测是否为列表项（数字开头或（一）、（二）等格式）
                    const isListItem = /^\d+\s*[\.、]\s*/.test(textContent) || /^（[一二三四五六七八九十\d]+）/.test(textContent);
                    
                    // 对于SCOPSR，检测中文全角空格开头的段落（通常是正文段落）
                    const hasChineseIndent = textContent.startsWith('　　');
                    
                    // 直接处理每个段落，保持原有结构，不进行任何自动拆分
                    processedParagraphs.push({
                        type: isHeading ? 'heading' : isListItem ? 'list' : 'paragraph',
                        content: textContent,
                        hasIndent: hasIndent || hasChineseIndent,
                        fontFamily: fontFamily,
                        originalStyle: style
                    });
                });
                
                // 构建结构化的文本内容 - 确保每个段落都有明确的分隔
                const structuredContent = processedParagraphs.map((para, index) => {
                    switch (para.type) {
                        case 'empty':
                            return ''; // 空段落
                        case 'heading':
                            return para.content; // 标题段落
                        case 'list':
                            return para.content; // 列表项段落
                        case 'paragraph':
                        default:
                            return para.content; // 普通段落
                    }
                }).filter(content => content !== ''); // 移除空字符串但保留实际的空段落
                
                // 使用双换行符分隔所有段落，确保Word中有明确的段落分隔
                const finalContent = structuredContent.join('\n\n');
                
                debugLog(`${siteName}段落提取结果: ${finalContent.length}个字符, ${paragraphs.length}个原始段落, ${structuredContent.length}个处理后段落`);
                debugLog('段落结构分析:', {
                    '总段落数': processedParagraphs.length,
                    '标题段落': processedParagraphs.filter(p => p.type === 'heading').length,
                    '列表段落': processedParagraphs.filter(p => p.type === 'list').length,
                    '普通段落': processedParagraphs.filter(p => p.type === 'paragraph').length,
                    '空段落': processedParagraphs.filter(p => p.type === 'empty').length,
                    '最终段落数': structuredContent.length
                });
                
                // 输出前几个段落的内容用于调试
                if (structuredContent.length > 0) {
                    debugLog('前3个段落预览:', {
                        '段落1': structuredContent[0]?.substring(0, 100) + '...',
                        '段落2': structuredContent[1]?.substring(0, 100) + '...',
                        '段落3': structuredContent[2]?.substring(0, 100) + '...'
                    });
                }
                
                return finalContent;
            }
            return null;
        }
        
        if (currentSite.domain === 'shbb.gov.cn') {
            // 对于SHBB网站，直接处理段落结构
            const result = processParagraphStructure(element, 'shbb.gov.cn');
            if (result) return result;
            // 如果没有找到p标签，回退到带换行的文本提取（使用视觉保留模式）
            return extractTextWithLineBreaks(element, true);
        } else if (currentSite.domain === 'scopsr.gov.cn') {
            // 对于SCOPSR网站，需要查找嵌套的段落结构
            debugLog('SCOPSR内容提取开始，查找嵌套段落结构...');
            
            // 首先尝试在.TRS_Editor或.Custom_UnionStyle中查找段落
            let paragraphContainer = element.querySelector('.TRS_Editor .Custom_UnionStyle');
            if (!paragraphContainer) {
                paragraphContainer = element.querySelector('.TRS_Editor');
            }
            if (!paragraphContainer) {
                paragraphContainer = element.querySelector('.Custom_UnionStyle');
            }
            
            if (paragraphContainer) {
                debugLog('找到SCOPSR段落容器:', paragraphContainer.className || 'no class');
                const result = processParagraphStructure(paragraphContainer, 'SCOPSR');
                if (result) return result;
            } else {
                debugLog('未找到SCOPSR段落容器，尝试直接在主元素中查找段落');
                const result = processParagraphStructure(element, 'SCOPSR');
                if (result) return result;
            }
            
            // 如果没有找到p标签，回退到带换行的文本提取
            debugLog('SCOPSR回退到通用文本提取');
            return extractTextWithLineBreaks(element);
        } else {
            // 其他网站使用通用处理
            return extractTextWithLineBreaks(element);
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
                        
                        // 获取搜索结果容器
                        const container = doc.querySelector(currentSite.selectors.searchContainer);
                        if (!container) {
                            console.log(`页面 ${pageUrl} 未找到搜索结果容器`);
                            resolve(articles);
                            return;
                        }
                        
                        // 获取文章元素
                        const articleElements = container.querySelectorAll(currentSite.selectors.articleItems);
                        
                        articleElements.forEach((element, index) => {
                            // 跳过表头行（针对SHBB网站）
                            if (currentSite.domain === 'shbb.gov.cn' && element.querySelector('th') && !element.querySelector('th a')) {
                                return;
                            }
                            
                            // 获取标题和链接
                            const titleLink = element.querySelector(currentSite.selectors.titleLink);
                            if (!titleLink) {
                                return;
                            }
                            
                            // 清理标题，移除HTML标签
                            const title = titleLink.textContent.replace(/<[^>]*>/g, '').trim();
                            const url = titleLink.href;
                            
                            if (!title || !url) {
                                return;
                            }
                            
                            // 获取日期
                            let date = '';
                            if (currentSite.domain === 'scopsr.gov.cn') {
                                const dateSpan = element.querySelector(currentSite.selectors.dateElement);
                                date = dateSpan ? dateSpan.textContent.trim() : '';
                            } else if (currentSite.domain === 'shbb.gov.cn') {
                                const dateTd = element.querySelector(currentSite.selectors.dateElement);
                                date = dateTd ? dateTd.textContent.trim() : '';
                            }
                            
                            // 获取摘要
                            let summaryText = '';
                            if (currentSite.domain === 'scopsr.gov.cn') {
                                const summary = element.querySelector(currentSite.selectors.summaryElement);
                                summaryText = summary ? summary.innerHTML.replace(/<[^>]*>/g, '').trim() : '';
                            } else if (currentSite.domain === 'shbb.gov.cn') {
                                const nextRow = element.nextElementSibling;
                                if (nextRow) {
                                    const summaryTd = nextRow.querySelector(currentSite.selectors.summaryElement);
                                    summaryText = summaryTd ? summaryTd.textContent.trim() : '';
                                }
                            }
                            
                            articles.push({
                                title: title,
                                url: url,
                                date: date,
                                summary: summaryText
                            });
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

    // 从URL中提取搜索词
    function getSearchWordFromUrl() {
        const urlParams = new URLSearchParams(window.location.search);
        if (currentSite.domain === 'scopsr.gov.cn') {
            const searchWord = urlParams.get('searchword');
            return searchWord ? decodeURIComponent(searchWord) : null;
        } else if (currentSite.domain === 'shbb.gov.cn') {
            const q = urlParams.get('q');
            return q ? decodeURIComponent(q) : null;
        }
        return null;
    }

    // 获取当前页码
    function getCurrentPageNumber() {
        return currentSite.pagination.getPageFromUrl(window.location.href);
    }

    // 检测总可用页数
    function getTotalAvailablePages() {
        console.log('开始检测总页数...');
        let maxPage = 1;

        // 查找分页容器
        const paginationContainer = document.querySelector(currentSite.selectors.paginationContainer);
        if (!paginationContainer) {
            console.log('未找到分页容器');
            return maxPage;
        }

        console.log('找到分页容器:', paginationContainer.innerHTML.substring(0, 200) + '...');
        
        if (currentSite.domain === 'scopsr.gov.cn') {
            // SCOPSR网站的分页处理
            // 查找"尾页"链接
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
            
            // 查找所有页码链接
            const pageLinks = paginationContainer.querySelectorAll(currentSite.selectors.pageLinks);
            pageLinks.forEach(element => {
                const pageText = element.textContent.trim();
                const pageNum = parseInt(pageText);
                if (!isNaN(pageNum)) {
                    console.log('从页码链接检测到页码:', pageNum);
                    maxPage = Math.max(maxPage, pageNum);
                }
                
                // 检查链接URL中的页码
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
        } else if (currentSite.domain === 'shbb.gov.cn') {
            // SHBB网站的分页处理
            const pageLinks = paginationContainer.querySelectorAll(currentSite.selectors.pageLinks);
            pageLinks.forEach(element => {
                const onclick = element.getAttribute('onclick');
                if (onclick) {
                    // 解析onclick中的页码
                    const pageMatch = onclick.match(/search_(\d+)\.jspx/);
                    if (pageMatch) {
                        const pageNum = parseInt(pageMatch[1]);
                        if (!isNaN(pageNum)) {
                            console.log('从SHBB页码链接检测到页码:', pageNum);
                            maxPage = Math.max(maxPage, pageNum);
                        }
                    } else if (onclick.includes('search.jspx')) {
                        // 第一页
                        maxPage = Math.max(maxPage, 1);
                    }
                }
                
                // 也检查链接文本
                const pageText = element.textContent.trim();
                if (pageText === '尾页') {
                    // 从总记录数计算页数
                    const totalRecordsElement = paginationContainer.querySelector('label span');
                    if (totalRecordsElement) {
                        const totalRecords = parseInt(totalRecordsElement.textContent);
                        if (!isNaN(totalRecords)) {
                            // 假设每页显示10条记录
                            const estimatedPages = Math.ceil(totalRecords / 10);
                            console.log('从总记录数估算页数:', estimatedPages);
                            maxPage = Math.max(maxPage, estimatedPages);
                        }
                    }
                } else {
                    const pageNum = parseInt(pageText);
                    if (!isNaN(pageNum)) {
                        maxPage = Math.max(maxPage, pageNum);
                    }
                }
            });
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

    // 生成文件名（包含页码范围和网站标识）
    function generateFilename(searchWord, startPage, endPage, actualPages) {
        const now = new Date();
        const dateStr = now.toISOString().slice(0, 10);
        const timeStr = now.toTimeString().slice(0, 8).replace(/:/g, '-');
        
        // 获取网站标识
        const sitePrefix = currentSite.domain === 'scopsr.gov.cn' ? 'SCOPSR' : 'SHBB';
        const baseFilename = searchWord || `文章集`;

        if (startPage === endPage || actualPages === 1) {
            // 单页
            return `${sitePrefix}_${baseFilename}_第${startPage}页_${dateStr}_${timeStr}.docx`;
        } else {
            // 多页范围
            return `${sitePrefix}_${baseFilename}_第${startPage}页到第${endPage}页_${dateStr}_${timeStr}.docx`;
        }
    }

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

        const multiPageBtn = document.querySelector('.save-articles-btn');
        multiPageBtn.disabled = true;
        multiPageBtn.textContent = '正在收集文章...';

        try {
            const allArticles = [];
            const allSuccessfulArticles = [];
            const allFailedArticles = [];
            const searchWord = getSearchWordFromUrl();
            let actualPagesProcessed = 0;
            let actualEndPage = startPage;
            
            debugLog(`开始多页保存，页面范围: ${startPage} 到 ${endPage}`);
            debugLog(`搜索关键词: ${searchWord || '无'}`);

            for (let page = startPage; page <= endPage; page++) {
                multiPageBtn.textContent = `正在处理第 ${page}页 (${page - startPage + 1}/${endPage - startPage + 1})...`;
                
                let pageArticles = [];
                if (page === currentPage) {
                    // 当前页面直接获取文章
                    pageArticles = collectArticleLinks();
                } else {
                    // 其他页面需要通过请求获取
                    const pageUrl = currentSite.pagination.buildPageUrl(window.location.href, page);
                    pageArticles = await fetchPageArticles(pageUrl);
                }
                
                if (pageArticles.length === 0) {
                    debugLog(`第 ${page} 页没有找到文章，可能已达到最后一页`);
                    break;
                }
                
                debugLog(`第 ${page} 页收集到 ${pageArticles.length} 篇文章`);

                actualPagesProcessed++;
                actualEndPage = page;

                // 获取每篇文章的内容
                for (let i = 0; i < pageArticles.length; i++) {
                    multiPageBtn.textContent = `第${page}页: 处理 ${i + 1}/${pageArticles.length} (总成功:${allSuccessfulArticles.length} 总失败:${allFailedArticles.length})`;
                    debugLog(`开始获取第${page}页第${i + 1}篇文章内容: ${pageArticles[i].title}`);
                    
                    const content = await fetchArticleContent(pageArticles[i].url);
                    
                    const articleWithPage = `[第${page}页] ${pageArticles[i].title}`;
                    
                    if (content && content !== '获取失败' && content !== '解析失败' && content !== '无法获取文章内容' && content !== 'HTTP错误: 404' && content !== '网络请求失败' && content !== '请求超时') {
                        allSuccessfulArticles.push(articleWithPage);
                        debugLog(`文章内容获取成功，长度: ${content.length}`);
                    } else {
                        allFailedArticles.push(articleWithPage);
                        debugLog(`文章内容获取失败: ${content}`);
                    }
                    
                    allArticles.push({
                        ...pageArticles[i],
                        content: content,
                        pageNumber: page
                    });
                    
                    // 添加延迟避免请求过快
                    await new Promise(resolve => setTimeout(resolve, 800));
                }
                
                console.log(`第 ${page} 页处理完成`);
            }

            if (allArticles.length === 0) {
                alert('未找到任何文章');
                return;
            }

            // 生成Word文档
            const wordHTML = createWordHTML(allArticles);
            const filename = generateFilename(searchWord, startPage, actualEndPage, actualPagesProcessed);
            
            saveAsDocx(wordHTML, filename);
            
            // 显示详细的多页保存结果
            showDetailedMultiPageResult(allSuccessfulArticles, allFailedArticles, allArticles.length, actualPagesProcessed, startPage, actualEndPage, filename);

        } catch (error) {
            debugLog('多页保存时出错:', error.message);
            console.error('多页保存详细错误:', error);
            alert(`多页保存失败: ${error.message}\n请查看控制台获取详细错误信息`);
        } finally {
            multiPageBtn.disabled = false;
            multiPageBtn.textContent = '批量保存';
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
            showDetailedModal(message);
        } else {
            alert(message);
        }
    }

    // 显示详细信息的模态框
    function showDetailedModal(message) {
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

    // 保存当前文章
    function saveCurrentArticle() {
        let content = '';
        let title = '';
        let date = '';
        
        // 获取文章内容
        const contentElement = document.querySelector(currentSite.selectors.articleContent);
        if (contentElement) {
            content = extractContentWithFormatting(contentElement);
        } else {
            // 备用选择器
            for (const selector of currentSite.selectors.articleContentBackup) {
                const element = document.querySelector(selector);
                if (element && element.textContent.length > 100) {
                    content = extractContentWithFormatting(element);
                    break;
                }
            }
        }

        if (!content) {
            alert('无法找到文章内容');
            return;
        }
        
        // 获取文章标题
        const titleElement = document.querySelector(currentSite.selectors.articleTitle);
        title = titleElement ? titleElement.textContent.trim() : document.title;
        
        // 获取发布时间
        const timeElement = document.querySelector(currentSite.selectors.articleTime);
        if (timeElement) {
            if (currentSite.domain === 'scopsr.gov.cn') {
                const timeMatch = timeElement.textContent.match(/时间：(\d{4}-\d{2}-\d{2})/);
                date = timeMatch ? timeMatch[1] : new Date().toISOString().slice(0, 10);
            } else if (currentSite.domain === 'shbb.gov.cn') {
                const timeMatch = timeElement.textContent.match(/(\d{4}-\d{2}-\d{2})/);
                date = timeMatch ? timeMatch[1] : new Date().toISOString().slice(0, 10);
            }
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
        const now = new Date();
        const timeStr = now.toTimeString().slice(0, 8).replace(/:/g, '-');
        
        // 获取网站标识
        const sitePrefix = currentSite.domain === 'scopsr.gov.cn' ? 'SCOPSR' : 'SHBB';
        const filename = `${sitePrefix}_${title}_${date}_${timeStr}.docx`;
        
        saveAsDocx(wordHTML, filename);
        
        alert(`文章已保存为DOCX文档！\n文件名：${filename}`);
    }

    // 等待页面加载完成后添加按钮
    function addButtons() {
        // 根据页面类型添加相应的按钮
        if (isSearchPage) {
            // 搜索结果页面：添加批量保存按钮
            const multiPageBtn = document.createElement('button');
            multiPageBtn.textContent = '批量保存文章';
            multiPageBtn.className = 'save-articles-btn';
            multiPageBtn.title = '批量保存搜索结果中的文章为DOCX格式';
            multiPageBtn.onclick = saveMultiplePages;
            document.body.appendChild(multiPageBtn);
            debugLog('已添加批量保存按钮到搜索页面');
        }

        if (isArticlePage) {
            // 文章详情页面：添加单独保存按钮
            const saveBtn = document.createElement('button');
            saveBtn.textContent = '保存本文为DOCX';
            saveBtn.className = 'save-articles-btn';
            saveBtn.title = '保存当前文章为DOCX格式';
            saveBtn.onclick = saveCurrentArticle;
            document.body.appendChild(saveBtn);
            debugLog('已添加单篇保存按钮到文章页面');
        }
    }

    // 跨平台兼容的按钮加载逻辑
    function addButtonsWithRetry() {
        let retryCount = 0;
        const maxRetries = 5;
        const retryInterval = 800;
        
        function attemptAddButtons() {
            retryCount++;
            debugLog(`第${retryCount}次尝试添加按钮...`);
            
            // 确保DOM完全准备就绪
            if (document.readyState !== 'complete' && retryCount < maxRetries) {
                debugLog('DOM未完全加载，延迟重试');
                setTimeout(attemptAddButtons, retryInterval);
                return;
            }
            
            // 确保body元素存在
            if (!document.body) {
                debugLog('document.body不存在，延迟重试');
                if (retryCount < maxRetries) {
                    setTimeout(attemptAddButtons, retryInterval);
                }
                return;
            }
            
            // 检查是否已存在按钮（避免重复创建）
            const existingButton = document.querySelector('.save-articles-btn');
            if (existingButton) {
                debugLog('按钮已存在，跳过创建');
                return;
            }
            
            try {
                // 执行原有的按钮添加逻辑
                addButtons();
                
                // 验证按钮是否创建成功
                setTimeout(() => {
                    const createdButton = document.querySelector('.save-articles-btn');
                    if (!createdButton && retryCount < maxRetries) {
                        debugLog(`按钮创建验证失败，准备第${retryCount + 1}次重试`);
                        setTimeout(attemptAddButtons, retryInterval);
                    } else if (createdButton) {
                        debugLog('按钮创建并验证成功');
                    } else {
                        debugLog('已达到最大重试次数，按钮创建失败');
                    }
                }, 200);
                
            } catch (error) {
                debugLog('创建按钮时出错:', error.message);
                if (retryCount < maxRetries) {
                    setTimeout(attemptAddButtons, retryInterval);
                }
            }
        }
        
        // 立即尝试一次
        attemptAddButtons();
        
        // 额外的保险机制：页面完全加载后再试一次
        if (document.readyState !== 'complete') {
            window.addEventListener('load', () => {
                setTimeout(() => {
                    const button = document.querySelector('.save-articles-btn');
                    if (!button) {
                        debugLog('页面load事件触发后补充尝试');
                        attemptAddButtons();
                    }
                }, 1000);
            });
        }
    }
    
    // 使用新的重试机制
    addButtonsWithRetry();
})();