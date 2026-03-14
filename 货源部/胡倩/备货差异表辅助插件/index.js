// ==UserScript==
// @name         备货差异表辅助插件
// @namespace    http://tampermonkey.net/
// @version      1.0
// @match        https://bi.erp321.com/*
// @grant        GM_xmlhttpRequest
// @connect      ivxubxfk0fm.feishu.cn
// @connect      apiweb.erp321.com
// @require      https://cdn.jsdelivr.net/npm/xlsx@0.18.5/dist/xlsx.full.min.js
// @updateURL    https://raw.githubusercontent.com/hzlop/-/main/货源部/胡倩/备货差异表辅助插件/index.js
// @downloadURL  https://raw.githubusercontent.com/hzlop/-/main/货源部/胡倩/备货差异表辅助插件/index.js
// ==/UserScript==

(function() {
    let Iid = null;
    let ArrItem = {};
    let cool_data = {};
    const version = '1.0';
    Object.freeze(version);

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
            {"k":"sales_qty_30",    "v":"0","c":">=","t":""},
            {"k":"A.seller_flag",   "v":"-5","c":"@=","t":""},
            {"k":"A.status",        "v":"WAITCONFIRM,WAITDELIVER,DELIVERING,SENT,QUESTION,WAITOUTERSENT,CANCELLED","c":"@=","t":""},
            {"k":"C.sent_flag",     "v":"1","c":"@=","t":""},
            {"k":"is_filter_empty", "v":"true","c":"@=","t":""}
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
                    if(!data || data.length <= 1){
                        alert("聚水潭没有返回数据");
                        return [];
                    };
                    // 处理数据：排除最后一行（总计行），并映射为中文键名
                    const processedData = {};

                    data.removeAt(data.length-1);
                    data.forEach(item =>{
                        processedData[item.i_id] = {"qty":item.qty - item.order_lock + item.purchase_qty,"vc_name":"","remark":""}
                    });

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

        /**
     * 输入款式编码，输出查询结果。
     * @param {Array<String>} Iid 要查询的款式编码
     * @returns {Promise<Object>} 查询结果
     */
    async function searchSrc(Iid) {
        let url = "https://apiweb.erp321.com/webapi/ItemApi/ItemSkuIm/GetPageListV2?__from=web_component&owner_co_id=10192205&authorize_co_id=10192205";
        let headers = {
            "content-type": "application/json; charset=utf-8",
            "gwfp": "8a9d59f44e15cc9d2be4acd656e93905",
            "priority": "u=1, i",
            "u_sso_token": "",
            "origin": "https://src.erp321.com",
            "referer": "https://src.erp321.com/",
            "webbox-request-id": "be0818c2-51e9-4191-83c7-ca4c4c723b05",
            "webbox-route-path": "/erp-components/goods-selector/"
        };
        let data = JSON.stringify({
            "ip":"",
            "uid":"21937473",
            "coid":"10192205",
            "page":{"currentPage":1,"pageSize":500},
            "data":{"queryFlds":
                ["pic","i_id","name","s_price","market_price","c_price","brand","c_name","vc_name","supplier_i_id","weight","l","w","h","volume","unit","onsale","remark","created","modified","c_id"],"orderBy":"","onsale":"1","i_id":Iid.join(','),"sku_id":"","c_id":""}
        });

        return new Promise((resolve, reject) => {
            GM_xmlhttpRequest({
                method: "POST",
                url: url,
                headers: headers,
                data: data,
                onload: function(response) {
                    console.log("请求成功，状态码：", response.status);
                    try {
                        let responseData = JSON.parse(response.responseText);
                        let dataList = responseData["data"]
                        let Datas = {};
                        dataList.forEach(item => {
                            Datas[item["i_id"]] = {
                                "i_id": item["i_id"],
                                "vc_name": item["vc_name"],
                                "remark": item["remark"],
                            };
                        });
                        resolve(Datas);
                    } catch (parseError) {
                        console.error("JSON解析失败：", parseError);
                        reject(parseError);
                    }
                },
                onerror: function(error) {
                    console.error("请求失败（网络/连接）:", error);
                    reject(error);
                },
                onabort: function() {
                    console.error("请求被中止");
                    reject(new Error("请求被中止"));
                },
                ontimeout: function() {
                    console.error("请求超时");
                    reject(new Error("请求超时"));
                }
            });
        });
    };

    function randomDelay(min = 2500, max = 4000) {
        const delay = Math.floor(Math.random() * (max - min + 1)) + min;
        return new Promise(resolve => setTimeout(resolve, delay));
    }

    function toXlsx(map){
        let xlsx = [['商家编码','法澳娜/（付锐）','韩一禾（丁成龙）','夏丽塔（小海）','抖音（黄芝情）','夏丽塔抖音','库存预警上限','实际库存','预警库存','虚拟分类','备注']]
        let i = 1;
        for (let item in map){
            xlsx.push([item,...cool_data[item],,map[item]['qty'],,map[item]['vc_name'],map[item]['remark']]);
            xlsx[i][6] = {t:'n',f:`SUM(B${i+1}:F${i+1})`};
            xlsx[i][8] = {t:'n',f:`H${i+1}-G${i+1}`};
            i+=1;
        };
        return(xlsx);
    };

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
                <span class="panel-title">备货差异表辅助插件</span>
            </div>
            <div id="tags-container">
                <input type="text" id="tag-input" placeholder="输入款式编码，回车确认" />
            </div>
            <div class="btn-group">
                <button class="action-btn secondary-btn" id="clear-tags-btn">清空</button>
                <button class="action-btn" id="do-search-btn">生成表格</button>
            </div>
            <div id="varsion-info">版本: ${version}</div>
        `;
        document.body.appendChild(panel);

        // 获取元素
        const fab = panel.querySelector('#sales-query-fab');
        const tagsContainer = panel.querySelector('#tags-container');
        const tagInput = panel.querySelector('#tag-input');
        const searchBtn = panel.querySelector('#do-search-btn');
        const clearBtn = panel.querySelector('#clear-tags-btn');

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
            const clipboardText = (e.clipboardData || window.clipboardData).getData('text');
            if (clipboardText){
                let items = clipboardText.split('\n').filter(i => i);
                items.forEach(item =>{
                    let row = item.split('\t').filter(i => i);
                    cool_data[row[0]] = [Number(row[1].trim()),Number(row[2].trim()),Number(row[3].trim()),Number(row[4].trim()),Number(row[5].trim())];
                })

                let text = Object.keys(cool_data);
                // 原输入中的组合项临时容器，用于删除分割后的原组合项
                const temparr = [];
                // 临时拆分后存储的组合项，用于加入到最终列表
                let temparrly = [];
                text.forEach(item => {
                    if (item.includes(',') || item.includes('，')) {
                        temparr.push(item);
                        const subItems = item.split(/[,，]/).map(i => i.trim()).filter(i => i);
                        temparrly.push(...subItems);
                        ArrItem[item] = subItems;
                    }
                });

                temparr.forEach(item => {
                    text = text.filter(i => i !== item);
                });

                text.push(...temparrly);

                addTags(text);
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
            searchBtn.classList.add('loading');

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

            try {
                const data = await querySalesData(iidString);

                for(let Group in ArrItem){
                    let qtySum = 0;
                    ArrItem[Group].forEach(g => {
                        if(data[g]){
                            qtySum += data[g]['qty'] || 0;
                            delete data[g];
                        }

                    })
                    data[Group] = {'qty':qtySum,'vc_name':"","remark":""};
                }

                await randomDelay();
                let arrtags = [];
                for(let item in data){
                    if(item.includes(',')){
                        let s = item.split(',');
                        arrtags.push(s[0]);
                        continue;
                    }
                    arrtags.push(item);
                };
                let vc = await searchSrc(arrtags);

                for(let item in data){
                    if(item.includes(',')){
                        let s = item.split(',');
                        data[item]['vc_name'] = vc[s[0]]['vc_name'];
                        data[item]['remark'] = vc[s[0]]['remark'];
                        continue;
                    }
                    data[item]['vc_name'] = vc[item]['vc_name'];
                    data[item]['remark'] = vc[item]['remark'];
                };
                
                const wb = XLSX.utils.book_new();
                const ws = XLSX.utils.aoa_to_sheet(toXlsx(data));
                XLSX.utils.book_append_sheet(wb, ws, "sheet1");
                XLSX.writeFile(wb, "备货差异表.xlsx");
                
                searchBtn.classList.remove('loading');
            } catch (e) {
                console.error("批量查询过程出错:", e);
                alert("查询过程中发生错误，请查看控制台。");
            } finally {
                searchBtn.classList.remove('loading');
                searchBtn.disabled = false;
                searchBtn.innerText = '抓取数据';
            }
        };


    }

    window.addEventListener('load', function() {
        const currentHref = window.location.href;
        // 检查是否是目标页面
        if (!currentHref.includes('subject/itemskuanalysis/sku_iid.aspx')) return;
        console.log('页面逻辑执行');
        // 初始化UI
        initUI();
    });

})();
