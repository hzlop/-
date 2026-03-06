// ==UserScript==
// @name         备货表辅助插件
// @namespace    http://tampermonkey.net/
// @version      0.3
// @match        https://bi.erp321.com/*
// @grant        GM_xmlhttpRequest
// @connect      ivxubxfk0fm.feishu.cn
// ==/UserScript==

(function() {
    'use strict';

    // 店铺配置
    const config_shop = {
        'XLT':{ 
            id: '10476202',
            name: '夏丽塔',
            url: 'https://ivxubxfk0fm.feishu.cn/base/automation/webhook/event/Lrxlaz90ywTnGEhmFYvcIUwonjh',
            header: {
                "Authorization": "Bearer Y8SQZgyjpl_F3MhrDi1_LHTH",
                "Content-Type": "application/json"
            }
        },

        'FAN':{ 
            id: '10270308',
            name: '法澳娜',
            url: 'https://ivxubxfk0fm.feishu.cn/base/automation/webhook/event/UdPTaCDkuwgUdhhaF24c0eKEn9c',
            header: {
                "Authorization": "Bearer IrJ0BoKDhRqNtHwWaEGlRKF_",
                "Content-Type": "application/json"
            }
        },

        'HYH':{ 
            id: '18023884',
            name: '韩一禾',
            url: 'https://ivxubxfk0fm.feishu.cn/base/automation/webhook/event/Qbvrahue0ws2uehTbykcQQU2nne',
            header: {
                "Authorization": "Bearer 7VodUCmhveKiRf5_gkzuNGAu",
                "Content-Type": "application/json"
            }
        }
    };

    let Iid = [];
    let send_data = [];
    let current_shop_key = 'FAN';
    let current_shop = config_shop[current_shop_key]; // 默认店铺
    let ArrItem = {};
    const version = '0.4';
    Object.freeze(version);

    // 核心查询函数
    async function querySalesData(iids) {
        console.log(`[备货表]开始准备查询...`);

        const viewState = window.document.querySelector('#__VIEWSTATE').value;
        const viewStateGen = window.document.querySelector('#__VIEWSTATEGENERATOR').value;
        if (!viewState || !viewStateGen) {
            alert('[备货表]错误：关键元素缺失，停止执行');
            return null;
        }

        // 3. 构建查询条件 (Search)
        const searchParams = [
            {"k":"sales_qty_30","v":"0","c":">=","t":""},
            {"k":"A.seller_flag","v":"-5","c":"@=","t":""},
            {"k":"A.status","v":"WAITCONFIRM,WAITDELIVER,DELIVERING,SENT,QUESTION,WAITOUTERSENT,CANCELLED","c":"@=","t":""},
            {"k":"C.sent_flag","v":"1","c":"@=","t":""},{"k":"A.shop_id","v":current_shop.id,"c":"@=","t":""},
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
                    // 处理数据：排除最后一行（总计行），并映射为中文键名
                    const processedData = data.slice(0, data.length - 1).map(item => ({
                        "图片": item.pic,
                        "款式编码": item.i_id,
                        "款式名称": item.name,
                        "今日销量": item.sale_qty_today,
                        "昨日销量": item.sales_qty_yesterday,
                        "七日销量": item.sales_qty_7,
                        "主仓实际库存": item.qty,
                        "订单占有数": item.order_lock,
                        "采购在途数": item.purchase_qty,
                        "实际库存": item.qty - item.order_lock + item.purchase_qty
                    }));
                    return processedData;
                } catch (parseErr) {
                    console.error("[备货表]解析 JSON 失败。");
                    console.error(parseErr);
                    return null;
                }
        } catch (err) {
            console.error("[备货表]请求发送失败", err);
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

            /* 结果表格面板样式 */
            #sales-result-panel {
                position: fixed;
                top: 50%;
                left: 50%;
                transform: translate(-50%, -50%);
                width: 800px;
                max-width: 90vw;
                height: 600px;
                max-height: 90vh;
                background-color: #fff;
                box-shadow: 0 4px 24px rgba(0,0,0,0.15);
                z-index: 100001;
                border-radius: 8px;
                display: none;
                flex-direction: column;
                font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, "Helvetica Neue", Arial, sans-serif;
            }
            #sales-result-panel.show {
                display: flex;
            }
            .result-header {
                padding: 16px 24px;
                border-bottom: 1px solid #f0f0f0;
                display: flex;
                justify-content: space-between;
                align-items: center;
                background: #fafafa;
                border-radius: 8px 8px 0 0;
            }
            .result-title {
                font-size: 18px;
                font-weight: 600;
                color: #262626;
            }
            .result-close {
                cursor: pointer;
                font-size: 20px;
                color: #999;
                transition: color 0.3s;
            }
            .result-close:hover {
                color: #ff4d4f;
            }
            .result-body {
                flex: 1;
                overflow: auto;
                padding: 16px;
            }
            .sales-table {
                width: 100%;
                border-collapse: collapse;
                border: 1px solid #f0f0f0;
            }
            .sales-table th, .sales-table td {
                padding: 12px 16px;
                border: 1px solid #f0f0f0;
                text-align: center;
            }
            .sales-table th {
                background: #fafafa;
                font-weight: 600;
                color: #262626;
                position: sticky;
                top: 0;
                z-index: 10;
                box-shadow: 0 1px 0 #f0f0f0;
            }
            /* 实际库存提示图标 */
            .stock-hint{
                display:inline-block;
                width:12px;
                height:12px;
                line-height:16px;
                border-radius:50%;
                background:#8f8f8f;
                color:#fff;
                font-size:10px;
                text-align:center;
                margin-left:6px;
                cursor:default;
                vertical-align:middle;
            }
            .stock-hint:hover{
                filter:brightness(0.95);
            }
            .sales-table tbody tr:hover {
                background-color: #f6faff;
            }
            .result-footer {
                padding: 16px 24px;
                border-top: 1px solid #f0f0f0;
                display: flex;
                justify-content: flex-end;
                gap: 12px;
                background: #fff;
                border-radius: 0 0 8px 8px;
            }
            #shop-select{
                padding: 2px 10px;
                border: 1px solid #d9d9d9;
                border-radius: 4px;
                background: #fff;
                color: #333;
                font-size: 13px;
            }
            .loading{
                opacity: 0.7;
                wait-cursor: wait;
            }
            #varsion-info {
                position: fixed;
                bottom: 10px;
                right: 10px;
                font-size: 12px;
                color: #999;
            }
        `;
        document.head.appendChild(style);

        // 创建侧边面板 (包含FAB)
        const panel = document.createElement('div');
        panel.id = 'sales-query-panel';
        panel.innerHTML = `
            <button id="sales-query-fab"><</button>
            <div class="panel-header">
                <span class="panel-title">${current_shop.name || '店铺'}数据抓取面板</span>
            </div>
            <div id="tags-container">
                <input type="text" id="tag-input" placeholder="输入款式编码，回车确认" />
            </div>
            <div class="btn-group">
                <button class="action-btn secondary-btn" id="clear-tags-btn">清空</button>
                <button class="action-btn" id="do-search-btn">抓取数据</button>
            </div>
            <div id="varsion-info">版本: ${version}</div>
        `;
        document.body.appendChild(panel);

        // 在面板头部右侧插入店铺下拉框
        (function addShopSelect(){
            try {
                const header = panel.querySelector('.panel-header');
                if (!header) return;
                const select = document.createElement('select');
                select.id = 'shop-select';
                // 填充选项
                Object.keys(config_shop).forEach(key => {
                    const opt = document.createElement('option');
                    opt.value = key;
                    opt.text = config_shop[key].name || key;
                    select.appendChild(opt);
                });
                // 默认值
                select.value = current_shop_key;
                // 绑定切换事件
                select.addEventListener('change', (e) => {
                    const k = e.target.value;
                    if (config_shop[k]) {
                        current_shop_key = k;
                        current_shop = config_shop[k];
                    }
                    select.classList.add('loading');
                    panel.querySelector('.panel-title').innerText = 'loading...';
                    clearBtn.click(); 
                    clearBtn.disabled = true;
                    searchBtn.disabled = true;
                    clearBtn.classList.add('disabled');
                    searchBtn.classList.add('disabled');
                    setTimeout(() => { 
                        select.classList.remove('loading'); 
                        panel.querySelector('.panel-title').innerText = `${current_shop.name || '店铺'}数据抓取面板`;
                        clearBtn.disabled = false;
                        searchBtn.disabled = false;
                        clearBtn.classList.remove('disabled');
                        searchBtn.classList.remove('disabled');
                    }, 800);
                });
                header.appendChild(select);
            } catch (e) {
                console.error('添加店铺下拉失败', e);
            }
        })();

        // 创建结果表格面板
        const resultPanel = document.createElement('div');
        resultPanel.id = 'sales-result-panel';
        resultPanel.innerHTML = `
            <div class="result-header">
                <span class="result-title">抓取结果</span>
                <span class="result-close">×</span>
            </div>
            <div class="result-body">
                <table class="sales-table">
                    <thead>
                        <tr>
                            <th>款式编码</th>
                            <th>七天销量</th>
                            <th>七天日均</th>
                            <th>昨日销量</th>
                            <th>当天销量</th>
                            <th>实际库存 <span class="stock-hint" title="实际库存=主仓库实际库存+采购在途库存-订单占有数">?</span></th>
                        </tr>
                    </thead>
                    <tbody id="sales-table-body">
                        <!-- 数据填充区 -->
                    </tbody>
                </table>
            </div>
            <div class="result-footer">
                <button class="action-btn secondary-btn" id="result-close-btn">关闭</button>
                <button class="action-btn" id="result-send-btn">发送</button>
            </div>
        `;
        document.body.appendChild(resultPanel);

        // 获取元素
        const fab = panel.querySelector('#sales-query-fab');
        const tagsContainer = panel.querySelector('#tags-container');
        const tagInput = panel.querySelector('#tag-input');
        const searchBtn = panel.querySelector('#do-search-btn');
        const clearBtn = panel.querySelector('#clear-tags-btn');

        // 结果面板元素
        const resultCloseIcon = resultPanel.querySelector('.result-close');
        const resultCloseBtn = resultPanel.querySelector('#result-close-btn');
        const resultSendBtn = resultPanel.querySelector('#result-send-btn');
        const salesTableBody = resultPanel.querySelector('#sales-table-body');

        //发送webhook（优先使用 GM_xmlhttpRequest，回退到 fetch）
        function sendWebhook(data) {
            // 如果 Tampermonkey 的 GM_xmlhttpRequest 可用，则使用它绕过页面级拦截
            if (typeof GM_xmlhttpRequest === 'function') {
                return new Promise((resolve, reject) => {
                    try {
                        GM_xmlhttpRequest({
                            method: 'POST',
                            url: current_shop.url,
                            headers: current_shop.header,
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
                        console.error('调用 GM_xmlhttpRequest 失败，回退至 fetch:', e);
                        // 若抛出异常则回退到 fetch
                        fetch(url, {
                            method: 'POST',
                            headers: { 'Content-Type': 'application/json' },
                            body: JSON.stringify(data)
                        }).then(r => {
                            if (r.ok) { console.log('Webhook发送成功 (fetch 回退)'); resolve(r); }
                            else { console.error('Webhook发送失败 (fetch 回退)'); reject(r); }
                        }).catch(err => { console.error('fetch 回退发送出错:', err); reject(err); });
                    }
                });
            }

            // 页面环境下回退到普通 fetch（可能会被广告/隐私扩展拦截）
            return fetch(url, {
                method: 'POST',
                headers: { 'Content-Type': 'application/json' },
                body: JSON.stringify(data)
            }).then(response => {
                if (response.ok) {
                    console.log('Webhook发送成功 (fetch)');
                } else {
                    console.error('Webhook发送失败 (fetch)');
                }
                return response;
            }).catch(err => {
                console.error('发送Webhook时出错 (fetch):', err);
                throw err;
            });
        }

        //发送按钮事件：每条间隔0.3s，每4条额外间隔0.5s，发送完为止
        resultSendBtn.onclick = async () => {

            // 防重复点击
            if (resultSendBtn.disabled) return;
            resultSendBtn.disabled = true;

            if(!send_data || send_data.length === 0) {
                alert('没有可发送的数据');
                return;
            }

            let rows = salesTableBody.querySelectorAll('tr');
            try {
                for (let i = 0; i < send_data.length; i++) {
                    const item = send_data[i];
                    try {
                        await sendWebhook(item);
                        if (rows[i]) rows[i].style.opacity = '0.4';
                        resultSendBtn.textContent = `${i + 1}/${send_data.length}`;
                    } catch (e) {
                        console.error('单条发送出错，继续下一条:', e);
                        if (rows[i]) rows[i].style.color = 'red';
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
                resultSendBtn.disabled = false;
            } finally {
                resultSendBtn.disabled = false;
            }
        };

        // 关闭结果面板函数
        const closeResultPanel = () => {
            resultPanel.classList.remove('show');
        };

        resultCloseIcon.onclick = closeResultPanel;
        resultCloseBtn.onclick = closeResultPanel;

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
            if (text) {
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

        // 标记是否已经查询完成
        let isQueryCompleted = false;

        // 执行查询
        searchBtn.onclick = async () => {
            // 如果已经查询完成，点击按钮直接显示面板
            if (isQueryCompleted) {
                resultPanel.classList.add('show');
                return;
            }

            // 防重复点击检查
            if (searchBtn.classList.contains('loading') || searchBtn.disabled) {
                console.warn('[备货表] 正在查询中，请勿重复点击');
                return;
            }

            // 获取所有标签文本
            const tags = Array.from(tagsContainer.querySelectorAll('.tag span:first-child')).map(span => span.innerText);

            if (tags.length === 0) {
                if(!confirm('没有输入任何款式编码\n是否要查询所有款式？')) {
                    return;
                }
            }
            // 填入 Iid 变量
            Iid = tags;
            const iidString = tags.join(',');

            // 执行查询
            searchBtn.classList.add('loading');
            searchBtn.disabled = true;
            //searchBtn.innerText = '查询中...';

            try {
                const data = await querySalesData(iidString);

                // 填充表格
                salesTableBody.innerHTML = ''; // 清空旧数据
                send_data = []; // 重置发送数据

                // 建立编码到返回数据的映射
                const dataMap = {};
                data.forEach(d => { dataMap[d["款式编码"]] = d; });

                // 提取所有原始的组合项
                const combinedGroups = Object.values(ArrItem || {}).filter(it => Array.isArray(it) && it.length > 1);
                const combinedCodeSet = new Set();
                combinedGroups.forEach(g => g.forEach(code => combinedCodeSet.add(code)));

                // 先处理组合项
                combinedGroups.forEach(group => {
                    const key = Object.keys(ArrItem).find(k => ArrItem[k] === group);
                    let sevenDaysSum = 0, yesterdaySum = 0, oneDaySum = 0, qtySum = 0;
                    group.forEach(code => {
                        const item = dataMap[code];
                        if (item) {
                            sevenDaysSum += Number(item["七日销量"]) || 0;
                            yesterdaySum += Number(item["昨日销量"]) || 0;
                            oneDaySum += Number(item["今日销量"]) || 0;
                            qtySum += Number(item["实际库存"]) || 0;
                        }
                    });
                    const sevenDaysAvg = Math.round(sevenDaysSum / 7);

                    send_data.push({
                        "款式编码": key,
                        "店铺名": current_shop.name,
                        "七日日均": sevenDaysAvg,
                        "昨日销量": yesterdaySum,
                        "今日销量": oneDaySum,
                        "实际库存": qtySum
                    });

                    const tr = document.createElement('tr');
                    tr.innerHTML = `
                        <td>${key + ' [组*]'}</td>
                        <td>${sevenDaysSum}</td>
                        <td>${sevenDaysAvg}</td>
                        <td>${yesterdaySum}</td>
                        <td>${oneDaySum}</td>
                        <td>${qtySum}</td>
                    `;
                    tr.style.fontWeight = '800'; 
                    salesTableBody.appendChild(tr);
                });

                // 再处理非组合的单项
                data.forEach(iid => {
                    const id = iid["款式编码"];
                    if (combinedCodeSet.has(id)) return; // 跳过已处理的组合内编码
                    const sevenDaysQty = iid["七日销量"] || 0;
                    const yesterdayQty = iid["昨日销量"] || 0;
                    const oneDayQty = iid["今日销量"] || 0;
                    const qty = iid["实际库存"] || 0;
                    const sevenDaysAvg = Math.round(sevenDaysQty / 7); // 四舍五入

                    send_data.push({
                        "款式编码": id,
                        "店铺名": current_shop.name,
                        "七日日均": sevenDaysAvg,
                        "昨日销量": yesterdayQty,
                        "今日销量": oneDayQty,
                        "实际库存": qty
                    });

                    const tr = document.createElement('tr');
                    tr.innerHTML = `
                        <td>${id}</td>
                        <td>${sevenDaysQty}</td>
                        <td>${sevenDaysAvg}</td>
                        <td>${yesterdayQty}</td>
                        <td>${oneDayQty}</td>
                        <td>${qty}</td>
                    `;
                    salesTableBody.appendChild(tr);
                });

                // 显示结果面板
                resultPanel.classList.add('show');
                isQueryCompleted = true; // 标记查询完成

                alert(`查询完成！`);

            } catch (e) {
                console.error("批量查询过程出错:", e);
                alert("查询过程中发生错误，请查看控制台。");
            } finally {
                searchBtn.classList.remove('loading');
                searchBtn.disabled = false;
                if (isQueryCompleted) {
                    searchBtn.innerText = '查看结果';
                } else {
                    searchBtn.innerText = '抓取数据';
                }
            }
        };
    }

    window.addEventListener('load', function() {
        const currentHref = window.location.href;
        // 检查是否是目标页面
        if (!currentHref.includes('subject/itemskuanalysis/sku_iid.aspx')) return;
        console.log('[备货表] 页面逻辑执行');
        console.log('[备货表] 当前页面Frame ID:', window.frameElement.id);

        // 初始化UI
        initUI();
    });

})();
