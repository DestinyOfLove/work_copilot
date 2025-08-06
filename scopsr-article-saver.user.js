// ==UserScript==
// @name         SCOPSR文章保存器
// @namespace    http://tampermonkey.net/
// @version      3.2
// @description  一键保存SCOPSR网站的文章为DOCX格式，保留原有格式
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
            共收录 ${articles.length} 篇文章
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
        // 搜索结果页面：添加批量保存按钮
        const saveBtn = document.createElement('button');
        saveBtn.textContent = '保存为DOCX文档';
        saveBtn.className = 'save-articles-btn';
        saveBtn.onclick = saveAllArticles;
        document.body.appendChild(saveBtn);

        async function saveAllArticles() {
            saveBtn.disabled = true;
            saveBtn.textContent = '正在收集文章...';

            try {
                // 获取所有文章链接
                const articles = collectArticleLinks();
                
                if (articles.length === 0) {
                    alert('未找到文章列表');
                    return;
                }

                saveBtn.textContent = `正在保存 ${articles.length} 篇文章...`;

                // 获取每篇文章的内容
                const articlesWithContent = [];
                let successCount = 0;
                let failCount = 0;
                
                for (let i = 0; i < articles.length; i++) {
                    saveBtn.textContent = `正在处理 ${i + 1}/${articles.length} (成功:${successCount} 失败:${failCount})`;
                    const content = await fetchArticleContent(articles[i].url);
                    
                    if (content && content !== '获取失败' && content !== '解析失败' && content !== '无法获取文章内容') {
                        successCount++;
                    } else {
                        failCount++;
                    }
                    
                    articlesWithContent.push({
                        ...articles[i],
                        content: content
                    });
                    // 添加延迟避免请求过快
                    await new Promise(resolve => setTimeout(resolve, 500));
                }

                // 生成Word文档
                const wordHTML = createWordHTML(articlesWithContent);
                // 文件名包含日期和时间（精确到秒）
                const now = new Date();
                const dateStr = now.toISOString().slice(0, 10);
                const timeStr = now.toTimeString().slice(0, 8).replace(/:/g, '-');
                const filename = `SCOPSR文章集_${dateStr}_${timeStr}.docx`;
                
                saveAsDocx(wordHTML, filename);
                
                alert(`保存完成！\n总计: ${articlesWithContent.length} 篇\n成功: ${successCount} 篇\n失败: ${failCount} 篇`);

            } catch (error) {
                console.error('保存文章时出错:', error);
                alert('保存失败，请查看控制台错误信息');
            } finally {
                saveBtn.disabled = false;
                saveBtn.textContent = '保存为DOCX文档';
            }
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
            
            alert('文章已保存为DOCX文档！');
        }
    }
})();