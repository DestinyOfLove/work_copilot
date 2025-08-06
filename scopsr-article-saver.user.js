// ==UserScript==
// @name         SCOPSR文章保存器
// @namespace    http://tampermonkey.net/
// @version      1.1
// @description  一键保存SCOPSR网站的文章列表和内容（支持跨域请求）
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

    if (isSearchPage) {
        // 搜索结果页面：添加批量保存按钮
        const saveBtn = document.createElement('button');
        saveBtn.textContent = '保存所有文章';
        saveBtn.className = 'save-articles-btn';
        saveBtn.onclick = saveAllArticles;
        document.body.appendChild(saveBtn);

        async function saveAllArticles() {
            saveBtn.disabled = true;
            saveBtn.textContent = '正在收集文章...';

            try {
                // 获取所有文章链接
                const articles = await collectArticleLinks();
                
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

                // 保存为JSON文件
                const jsonData = JSON.stringify(articlesWithContent, null, 2);
                const blob = new Blob([jsonData], { type: 'application/json' });
                const url = URL.createObjectURL(blob);
                const filename = `scopsr_articles_${new Date().toISOString().slice(0, 10)}.json`;
                
                GM_download({
                    url: url,
                    name: filename,
                    onload: () => {
                        alert(`保存完成！\n总计: ${articlesWithContent.length} 篇\n成功: ${successCount} 篇\n失败: ${failCount} 篇`);
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

            } catch (error) {
                console.error('保存文章时出错:', error);
                alert('保存失败，请查看控制台错误信息');
            } finally {
                saveBtn.disabled = false;
                saveBtn.textContent = '保存所有文章';
            }
        }

        function collectArticleLinks() {
            const articles = [];
            
            // 获取搜索结果容器中的所有文章
            const container = document.querySelector('#searchresult, .jiansuo-container-result');
            if (!container) return articles;
            
            // 每个ul包含一篇文章
            const articleElements = container.querySelectorAll('ul');
            
            articleElements.forEach(ul => {
                // 获取标题和链接
                const titleLink = ul.querySelector('li h2 a');
                if (!titleLink) return;
                
                const title = titleLink.textContent.trim();
                const url = titleLink.href;
                
                // 获取日期
                const dateSpan = ul.querySelector('li div.jiansuo-result-link span');
                const date = dateSpan ? dateSpan.textContent.trim() : '';
                
                // 获取摘要
                const summary = ul.querySelector('li p');
                const summaryText = summary ? summary.textContent.trim() : '';
                
                if (url && url.includes('/xwzx/')) {
                    articles.push({
                        title: title,
                        url: url,
                        date: date,
                        summary: summaryText
                    });
                }
            });

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
        saveBtn.textContent = '保存本文';
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

            const jsonData = JSON.stringify(article, null, 2);
            const blob = new Blob([jsonData], { type: 'application/json' });
            const url = URL.createObjectURL(blob);
            const filename = `article_${date}_${new Date().getTime()}.json`;
            
            const a = document.createElement('a');
            a.href = url;
            a.download = filename;
            a.click();
            URL.revokeObjectURL(url);
            
            alert('文章已保存！');
        }
    }
})();