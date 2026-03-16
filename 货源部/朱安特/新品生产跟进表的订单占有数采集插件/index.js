// ==UserScript==
// @name         订单占有数采集
// @namespace    http://tampermonkey.net/
// @version      1.0
// @match        https://bi.erp321.com/*
// @grant        GM_xmlhttpRequest
// @connect      ivxubxfk0fm.feishu.cn
// @updateURL    https://raw.githubusercontent.com/hzlop/-/main/货源部/朱安特/新品生产跟进表的订单占有数采集/index.js
// @downloadURL  https://raw.githubusercontent.com/hzlop/-/main/货源部/朱安特/新品生产跟进表的订单占有数采集/index.js
// ==/UserScript==

(function() {
    let Iid = [];
    let send_data = [];
    // 核心查询函数
    async function querySalesData(iids) {
        const viewState = window.document.querySelector('#__VIEWSTATE').value;
        const viewStateGen = window.document.querySelector('#__VIEWSTATEGENERATOR').value;
        if (!viewState || !viewStateGen) {
            alert('错误：关键元素缺失，停止执行');
            return null;
        }

        // 3. 构建查询条件 (Search)
        const searchParams = [
            {"k":"sales_qty_30","v":"0","c":">=","t":""},
            {"k":"A.seller_flag","v":"-5","c":"@=","t":""},
            {"k":"A.status","v":"WAITCONFIRM,WAITDELIVER,DELIVERING,SENT,QUESTION,WAITOUTERSENT,CANCELLED","c":"@=","t":""},
            {"k":"C.sent_flag","v":"1","c":"@=","t":""},
            {"k":"is_filter_empty","v":"true","c":"@=","t":""}
        ];

        if (iids && iids.length > 0) {
            searchParams.unshift({"k":"D.i_id","v":iids,"c":"@=","t":""});
        }

        const callbackParam = {"Method":"LoadDataToJSON","Args":["1","{\"fld\":\"qty\",\"type\":\"desc\"}",""],"CallControl":"{page}"};

        const eventvalidation = window.document.querySelector('#__EVENTVALIDATION').value;
        // 5. 组装 Form Data
        const formData = new URLSearchParams();
        formData.append('__VIEWSTATE', viewState);
        formData.append('__VIEWSTATEGENERATOR', viewStateGen);
        formData.append('search', JSON.stringify(searchParams));
        formData.append('dataPageCount', '');
        formData.append('style', 'normal');
        formData.append('style_flds', 'normal');
        formData.append('__CALLBACKID', 'ACall1');
        formData.append('__CALLBACKPARAM', JSON.stringify(callbackParam));
        formData.append('__EVENTVALIDATION', eventvalidation);

        // 6. 发送请求
        try {
            const timestamp = new Date().getTime();
            const url = document.querySelector("#form1").action + `&ts___=${timestamp}&am___=LoadDataToJSON`;
            const response = await fetch(url, {
                method: "POST",
                headers: {
                    "content-type": "application/x-www-form-urlencoded; charset=UTF-8",
                    "x-requested-with": "XMLHttpRequest"
                },
                body: formData
            });

            const text = await response.text();
            let jsonObjStartIndex = text.indexOf('{');
            let realJsonStr = text.substring(jsonObjStartIndex);
            try {
                    const data = JSON.parse(JSON.parse(realJsonStr)['ReturnValue'])['datas'];
                    if(!data || data.length === 0){
                        alert("聚水潭没有返回数据");
                        return [];
                    };
                    // 处理数据：排除最后一行（总计行）
                    const processedData = data.slice(0, data.length - 1).map(item => ({
                        "item": item.i_id,
                        "Occupancy": item.order_lock,
                    }));
                    return processedData;
                } catch (parseErr) {
                    console.error("解析 JSON 失败。");
                    console.error(parseErr);
                    return null;
                }
        } catch (err) {
            console.error("请求发送失败", err);
            return null;
        }
    }

    // UI 构建与事件绑定
    function initUI() {
        // 注入样式
        const style = document.createElement('style');
        style.innerHTML = `
            #sales-query-panel {
                position: fixed;
                top: 50%;
                right: -320px;
                transform: translateY(-50%);
                width: 320px;
                height: auto;
                max-height: 90vh;
                background-color: #fff;
                box-shadow: -4px 0 16px rgba(0,0,0,0.08);
                z-index: 100000;
                transition: right 0.3s cubic-bezier(0.4, 0, 0.2, 1);
                padding: 24px;
                box-sizing: border-box;
                display: flex;
                flex-direction: column;
                border-radius: 8px 0 0 8px;
                font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, "Helvetica Neue", Arial, sans-serif;
            }
            #sales-query-panel.open {
                right: 0;
            }
            /* 侧边按钮 (FAB) - 挂在面板左侧外 */
            #sales-query-fab {
                position: absolute;
                left: -20px;
                top: 50%;
                transform: translateY(-50%);
                width: 20px;
                height: 100px;
                background-color: #1890ff;
                color: white;
                border: none;
                border-radius: 4px 0 0 4px;
                cursor: pointer;
                box-shadow: -2px 0 6px rgba(0,0,0,0.1);
                display: flex;
                align-items: center;
                justify-content: center;
                font-size: 14px;
                line-height: 1;
                outline: none;
                transition: background-color 0.3s;
            }
            #sales-query-fab:hover {
                background-color: #40a9ff;
            }

            .panel-header {
                display: flex;
                justify-content: space-between;
                align-items: center;
                margin-bottom: 20px;
                border-bottom: 1px solid #f0f0f0;
                padding-bottom: 12px;
            }
            .panel-title {
                font-size: 16px;
                font-weight: 600;
                color: #262626;
            }

            #tags-container {
                width: 100%;
                height: 300px;
                margin-bottom: 16px;
                padding: 8px 12px;
                border: 1px solid #d9d9d9;
                border-radius: 4px;
                background: #fafafa;
                overflow-y: auto;
                display: flex;
                flex-wrap: wrap;
                align-content: flex-start;
                gap: 8px;
                cursor: text;
                transition: all 0.3s;
            }
            #tags-container:focus-within {
                border-color: #40a9ff;
                background: #fff;
                box-shadow: 0 0 0 2px rgba(24, 144, 255, 0.2);
            }
            .tag {
                background-color: #e6f7ff;
                border: 1px solid #91d5ff;
                border-radius: 2px;
                padding: 2px 8px;
                font-size: 12px;
                color: #1890ff;
                display: flex;
                align-items: center;
                height: 24px;
                box-sizing: border-box;
                white-space: nowrap;
                animation: fadeIn 0.2s ease-in-out;
            }
            @keyframes fadeIn {
                from { opacity: 0; transform: scale(0.9); }
                to { opacity: 1; transform: scale(1); }
            }
            .tag-close {
                margin-left: 6px;
                cursor: pointer;
                font-size: 14px;
                color: #1890ff;
                opacity: 0.6;
                transition: opacity 0.2s;
            }
            .tag-close:hover {
                opacity: 1;
            }
            #tag-input {
                border: none;
                outline: none;
                flex-grow: 1;
                min-width: 100px;
                height: 24px;
                font-size: 13px;
                background: transparent;
                color: #333;
            }

            .btn-group {
                display: flex;
                gap: 12px;
            }
            .action-btn {
                flex: 1;
                padding: 8px 16px;
                background-color: #1890ff;
                color: white;
                border: none;
                border-radius: 4px;
                cursor: pointer;
                font-size: 14px;
                transition: all 0.3s;
                font-weight: 500;
            }
            .action-btn:hover {
                background-color: #40a9ff;
                box-shadow: 0 2px 8px rgba(24, 144, 255, 0.3);
            }
            .secondary-btn {
                background-color: #fff;
                color: #666;
                border: 1px solid #d9d9d9;
                flex: 0 0 80px;
            }
            .secondary-btn:hover {
                background-color: #fff;
                color: #ff4d4f;
                border-color: #ff4d4f;
                box-shadow: 0 2px 8px rgba(255, 77, 79, 0.2);
            }

            .action-btn.loading {
                opacity: 0.7;
                cursor: wait;
            }
            .action-btn.loading::after {
                content: "";
                position: absolute;
                top: 50%;
                left: 50%;
                width: 14px;
                height: 14px;
                margin-top: -7px;
                margin-left: -7px;
                border: 2px solid #fff;
                border-top-color: transparent;
                border-radius: 50%;
                animation: spin 0.8s linear infinite;
            }
            .disabled {
                opacity: 0.5;
                cursor: not-allowed;
            }
            @keyframes spin {
                to { transform: rotate(360deg); }
            }
            .loading{
                opacity: 0.7;
                wait-cursor: wait;
            }
        `;
        document.head.appendChild(style);

        // 创建侧边面板 (包含FAB)
        const panel = document.createElement('div');
        panel.id = 'sales-query-panel';
        panel.innerHTML = `
            <button id="sales-query-fab"><</button>
            <div class="panel-header">
                <span class="panel-title">面板</span>
            </div>
            <div id="tags-container">
                <input type="text" id="tag-input" placeholder="输入款式编码，回车确认" />
            </div>
            <div class="btn-group">
                <button class="action-btn secondary-btn" id="clear-tags-btn">清空</button>
                <button class="action-btn" id="do-search-btn">抓取数据</button>
            </div>
        `;
        document.body.appendChild(panel);

        // 获取元素
        const fab = panel.querySelector('#sales-query-fab');
        const tagsContainer = panel.querySelector('#tags-container');
        const tagInput = panel.querySelector('#tag-input');
        const searchBtn = panel.querySelector('#do-search-btn');
        const clearBtn = panel.querySelector('#clear-tags-btn');

        //发送webhook（优先使用 GM_xmlhttpRequest，回退到 fetch）
        async function sendWebhook(data) {
            return new Promise((resolve, reject) => {
                try {
                    GM_xmlhttpRequest({
                        method: 'POST',
                        url: `https://ivxubxfk0fm.feishu.cn/base/automation/webhook/event/NBa8aPcNaw5NFShYu3ZcHMhyn6p`,
                        headers: {
                                "Authorization": "Bearer 8WwY9Y9pyn0Rog4MhSAyZLIq",
                                "Content-Type": "application/json"
                            },
                        data: JSON.stringify(data),
                        onload: function(res) {
                            if (res.status >= 200 && res.status < 300) {
                                resolve(res);
                            } else {
                                console.error('Webhook发送失败 (GM_xmlhttpRequest)', res);
                                reject(res);
                            }
                        },
                        onerror: function(err) {
                            console.error('GM_xmlhttpRequest 错误:', err);
                            reject(err);
                        }
                    });
                } catch (e) {
                    console.error('发送webhook出现错误:', e);
                }
            });
        }

        //发送按钮事件：每条间隔0.3s，每4条额外间隔0.5s，发送完为止
        async function send(send_data) {
            if(!send_data || send_data.length === 0) {
                alert('没有可发送的数据');
                return;
            }
            try {
                for (let i = 0; i < send_data.length; i++) {
                    const item = send_data[i];
                    try {
                        await sendWebhook(item);
                    } catch (e) {
                        console.error('单条发送出错，继续下一条:', e);
                    }

                    // 每条间隔 300ms
                    await new Promise(res => setTimeout(res, 300));

                    // 每发送 4 条再额外等待 500ms（i+1 对应第几条，从 1 开始计数）
                    if ((i + 1) % 4 === 0) {
                        await new Promise(res => setTimeout(res, 500));
                    }
                }
            } catch (e) {
                console.error('发送数据出错:', e);
                alert('发送数据出错，请检查控制台日志');
            } finally {
            }
        };

        // 标签管理函数
        function createTag(text) {
            const tag = document.createElement('div');
            tag.className = 'tag';
            tag.innerHTML = `
                <span>${text}</span>
                <span class="tag-close">×</span>
            `;
            // 删除事件
            tag.querySelector('.tag-close').onclick = (e) => {
                e.stopPropagation();
                tag.remove();
            };
            return tag;
        }

        function addTags(texts) {
            const fragment = document.createDocumentFragment();
            texts.forEach(text => {
                if (text && text.trim()) {
                    fragment.appendChild(createTag(text.trim()));
                }
            });
            tagsContainer.insertBefore(fragment, tagInput);
            // 滚动到底部
            tagsContainer.scrollTop = tagsContainer.scrollHeight;
        }

        // 容器点击聚焦输入框
        tagsContainer.onclick = (e) => {
            if (e.target === tagsContainer) {
                tagInput.focus();
            }
        };

        // 输入框事件
        tagInput.addEventListener('keydown', (e) => {
            if (e.key === 'Enter') {
                e.preventDefault();
                const text = tagInput.value.trim();
                if (text) {
                    addTags([text]);
                    tagInput.value = '';
                }
            } else if (e.key === 'Backspace' && tagInput.value === '') {
                // 删除最后一个标签
                const tags = tagsContainer.querySelectorAll('.tag');
                if (tags.length > 0) {
                    tags[tags.length - 1].remove();
                }
            }
        });

        // 粘贴事件处理
        tagInput.addEventListener('paste', (e) => {
            e.preventDefault();
            const text = (e.clipboardData || window.clipboardData).getData('text');
            if (text && text!= '') {
                // 按换行、空格、逗号分割
                let items = text.split(/[\r\n\s]+/).filter(i => i);
                // 原输入中的组合项临时容器，用于删除分割后的原组合项
                const temparr = [];
                // 临时拆分后存储的组合项，用于加入到最终列表
                let temparrly = [];
                items.forEach(item => {
                    if (item.includes(',') || item.includes('，')) {
                        temparr.push(item);
                        const subItems = item.split(/[,，]/).map(i => i.trim()).filter(i => i);
                        temparrly.push(...subItems);
                        ArrItem[item] = subItems;
                    }
                });

                temparr.forEach(item => {
                    items = items.filter(i => i !== item);
                });

                items.push(...temparrly);

                addTags(items);
                tagInput.value = ''; // 确保清空
            }
        });

        // 失去焦点时如果还有内容，自动转换
        tagInput.addEventListener('blur', () => {
            const text = tagInput.value.trim();
            if (text) {
                addTags([text]);
                tagInput.value = '';
            }
        });

        // 切换面板状态
        fab.onclick = () => {
            panel.classList.toggle('open');
            const isOpen = panel.classList.contains('open');
            fab.innerText = isOpen ? '>' : '<';
            if (isOpen) {
                setTimeout(() => tagInput.focus(), 100); // 自动聚焦
            }
        };

        // 清空按钮事件
        clearBtn.onclick = () => {
            // 移除所有 .tag 元素
            const tags = tagsContainer.querySelectorAll('.tag');
            tags.forEach(tag => tag.remove());
            ArrItem = [];
            tagInput.value = '';
            tagInput.focus();
        };

        // 执行查询
        searchBtn.onclick = async () => {
            // 防重复点击检查
            if (searchBtn.classList.contains('loading') || searchBtn.disabled) {
                return;
            }

            // 获取所有标签文本
            const tags = Array.from(tagsContainer.querySelectorAll('.tag span:first-child')).map(span => span.innerText);

            if (tags.length === 0) return;
            // 填入 Iid 变量
            Iid = tags;
            const iidString = tags.join(',');

            // 执行查询
            searchBtn.classList.add('loading');
            searchBtn.disabled = true;

            try {
                const data = await querySalesData(iidString);
                await send(data);
            } catch (e) {
                console.error("批量查询过程出错:", e);
                alert("查询过程中发生错误，请查看控制台。");
            } finally {
                searchBtn.classList.remove('loading');
                searchBtn.disabled = false;
                searchBtn.innerText = '抓取数据';
                clearBtn.onclick();
                alert('任务完成');
            }
        };
    }

    window.addEventListener('load', function() {
        const currentHref = window.location.href;
        // 检查是否是目标页面
        if (!currentHref.includes('subject/itemskuanalysis/sku_iid.aspx')) return;
        console.log('页面逻辑执行');
        console.log('当前页面Frame ID:', window.frameElement.id);

        // 初始化UI
        initUI();
    });

})();
