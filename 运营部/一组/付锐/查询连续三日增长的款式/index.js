// ==UserScript==
// @name         推送插件
// @namespace    http://tampermonkey.net/
// @version      3.1
// @match        https://*.erp321.com/*
// @grant        GM_xmlhttpRequest
// @connect      ivxubxfk0fm.feishu.cn
// @run-at       document-end
// @updateURL    https://raw.githubusercontent.com/hzlop/-/main/运营部/一组/付锐/查询连续三日增长的款式/index.js
// @downloadURL  https://raw.githubusercontent.com/hzlop/-/main/运营部/一组/付锐/查询连续三日增长的款式/index.js
// ==/UserScript==

(function () {
    // ─── 常量 & 配置 ────────────────────────────────────────────────────────────

    const SHOPS = {
        XLT: { ID: '10476202', Name: '夏丽塔旗舰店' },
        FAN: { ID: '10270308', Name: '法澳娜旗舰店' },
        HYH: { ID: '18023884', Name: 'hamyiho韩一禾' },
    };

    const TEST_MODE = false;

    const FEISHU_MSG_URL  = 'https://ivxubxfk0fm.feishu.cn/base/workflow/webhook/event/NJEca7If4wUFEWhUqNYcRWxUnlb';
    const FEISHU_XLSX_URL = 'https://ivxubxfk0fm.feishu.cn/base/workflow/webhook/event/Rnlqa8oYrwFau6hIDEecD66Knrc';
    const FEISHU_XLSX_HEADERS = {
        Authorization:  'Bearer MtzVjcKs4WivXsQ5O4UeAjg-',
        'Content-Type': 'application/json',
    };

    // ─── 运行时状态 ─────────────────────────────────────────────────────────────

    let currentShopID = SHOPS.FAN.ID;
    const btnEl   = createBtn();
    const wordsEl = createWordsSpan();
    const testBtnEl = createTestBtn();

    // ─── 工具函数 ────────────────────────────────────────────────────────────────

    /** 通过店铺 ID 反查店铺名称 */
    function shopNameById(id) {
        if (typeof id !== 'string') return '-';
        const match = Object.values(SHOPS).find(shop => shop.ID === id);
        return match ? match.Name : '-';
    }

    function shopIdByName(name) {
        if (typeof name !== "string") return '-';
        const match = Object.values(SHOPS).find(shop => shop.Name === name);
        return match ? match.ID : '-';
    }

    /** 返回 [n天前日期, 今日日期]，格式 YYYY-MM-DD */
    function getDateRange(n = 3) {
        const today   = new Date();
        const daysAgo = new Date(today);
        daysAgo.setDate(today.getDate() - n);

        const fmt = date => {
            const y = date.getFullYear();
            const m = String(date.getMonth() + 1).padStart(2, '0');
            const d = String(date.getDate()).padStart(2, '0');
            return `${y}-${m}-${d}`;
        };

        return [fmt(daysAgo), fmt(today)];
    }

    /** 时间戳 → MM-DD */
    function timestampToMonthDay(ts) {
        const date = new Date(ts);
        const m = String(date.getMonth() + 1).padStart(2, '0');
        const d = String(date.getDate()).padStart(2, '0');
        return `${m}-${d}`;
    }

    /** 随机延迟（ms） */
    function randomDelay(min = 2500, max = 4000) {
        const delay = Math.floor(Math.random() * (max - min + 1)) + min;
        return new Promise(resolve => setTimeout(resolve, delay));
    }

    // ─── 数据处理 ────────────────────────────────────────────────────────────────

    /**
     * 校验数组：按第一个元素升序排列后，第二个元素是否严格递增（连续三天增长）。
     * @returns {{ isValid: boolean, sortedArr: Array }}
     */
    function checkGrowthArray(arr) {
        const sorted = [...arr].sort((a, b) => a[0] - b[0]);
        if (sorted.length < 3) return { isValid: false, sortedArr: sorted };

        for (let i = 1; i < sorted.length; i++) {
            if (sorted[i][1] <= sorted[i - 1][1]) {
                return { isValid: false, sortedArr: sorted };
            }
        }
        return { isValid: true, sortedArr: sorted };
    }

    /**
     * 过滤出连续增长的款式，并将时间戳转换为 MM-DD 格式。
     */
    function cleanData(rawDict) {
        const result = {};
        for (const key of Object.keys(rawDict)) {
            const { isValid, sortedArr } = checkGrowthArray(rawDict[key]);
            if (!isValid) continue;
            // 避免直接修改 sortedArr 内部对象（浅拷贝每一项）
            result[key] = sortedArr.map(item => [timestampToMonthDay(item[0]), item[1]]);
        }
        console.log('map:', result);
        return result;
    }

    // ─── 格式化输出 ──────────────────────────────────────────────────────────────

    /**
     * 生成 Markdown 表格文本（用于飞书消息）。
     * 要求 map 非空且每条数据至少3项，调用前应已由 cleanData 保证。
     */
    function formatMarkdownTable(map) {
        const firstKey = Object.keys(map)[0];
        const dates    = map[firstKey];

        if (!dates || dates.length < 3) {
            console.warn('[formatMarkdownTable] 数据不足3天，跳过');
            return '';
        }

        let table = `| 款式 | ${dates[0][0]} | ${dates[1][0]} | ${dates[2][0]} |\n| ------ | ---- | ---- | ---- |\n`;

        for (const key of Object.keys(map)) {
            const row = map[key].map(item => ` ${item[1]} |`).join('');
            table += `| ${key} |${row}\n`;
        }
        return table;
    }

    /**
     * 生成多维表格行数组（用于飞书表格）。
     * 要求 map 非空且每条数据至少3项，调用前应已由 cleanData 保证。
     */
    async function formatXlsxRows(map, shopName) {
        const firstKey = Object.keys(map)[0];
        const dates    = map[firstKey];
        const styleId = await fetchStyleId(Array.from(Object.keys(map)),shopIdByName(shopName));

        if (!dates || dates.length < 3) {
            console.warn('[formatXlsxRows] 数据不足3天，跳过');
            return [];
        }

        const rows = [{
            店铺: shopName,
            款式: '日期',
            大前天: dates[0][0],
            前天:   dates[1][0],
            昨天:   dates[2][0],
            ID: '-',
        }];

        for (const key of Object.keys(map)) {
            rows.push({
                店铺: shopName,
                款式: key,
                大前天: map[key][0][1],
                前天:   map[key][1][1],
                昨天:   map[key][2][1],
                ID: styleId[key],
            });
        }

        console.log('Xlsx:', rows);
        return rows;
    }

    // ─── 飞书推送 ────────────────────────────────────────────────────────────────

    /**
     * 发送文本消息到飞书，返回 Promise。
     */
    function sendFeishuMessage(text) {
        return new Promise((resolve, reject) => {
            GM_xmlhttpRequest({
                method:  'POST',
                url:     FEISHU_MSG_URL,
                headers: { 'Content-Type': 'application/json; charset=utf-8' },
                data:    JSON.stringify({ text }),
                onload: res => {
                    if (res.status < 200 || res.status >= 300) {
                        console.error('[sendFeishuMessage] 发送失败:', res.status, res.responseText);
                        reject(new Error(`飞书消息发送失败: HTTP ${res.status}`));
                    } else {
                        console.log('Message sent:', res.status, res.responseText);
                        resolve();
                    }
                },
                onerror: err => {
                    console.error('[sendFeishuMessage] 网络错误:', err);
                    reject(new Error('飞书消息发送网络错误'));
                },
            });
        });
    }

    /**
     * 逐条（串行）发送表格数据到飞书多维表格。
     * 修复原版中并行触发所有请求、Promise 提前 resolve 的问题。
     */
    async function sendXlsxToFeishu(rows) {
        for (let i = 0; i < rows.length; i++) {
            await new Promise((resolve, reject) => {
                GM_xmlhttpRequest({
                    method:  'POST',
                    url:     FEISHU_XLSX_URL,
                    headers: FEISHU_XLSX_HEADERS,
                    data:    JSON.stringify(rows[i]),
                    onload: res => {
                        if (res.status < 200 || res.status >= 300) {
                            console.error('[sendXlsxToFeishu] Webhook发送失败:', res.status, rows[i]);
                            // 单条失败记录日志，不中断整体流程
                            resolve();
                        } else {
                            wordsEl.innerHTML = `已经发送给表格 ${i + 1}/${rows.length} 条`;
                            resolve();
                        }
                    },
                    onerror: err => {
                        console.error('[sendXlsxToFeishu] GM_xmlhttpRequest 错误:', err, rows[i]);
                        // 单条网络错误同样不中断整体
                        resolve();
                    },
                });
            });
            // 每条之间加短暂延迟，避免触发限流
            if (i < rows.length - 1) {
                await randomDelay(300, 500);
            }
        }
    }

    // ─── 数据拉取 ────────────────────────────────────────────────────────────────

    /** 分页拉取搜索数据，返回合并后的原始记录数组 */
    async function fetchSearchData() {
        const viewStateEl    = document.querySelector('#__VIEWSTATE');
        const viewStateGenEl = document.querySelector('#__VIEWSTATEGENERATOR');

        if (!viewStateEl || !viewStateGenEl) {
            console.error('[fetchSearchData] 找不到 VIEWSTATE 表单元素');
            return Promise.reject('参数错误：找不到 VIEWSTATE');
        }

        const viewState    = viewStateEl.value;
        const viewStateGen = viewStateGenEl.value;

        if (!viewState || !viewStateGen) {
            console.error('[fetchSearchData] VIEWSTATE 值为空');
            return Promise.reject('参数错误：VIEWSTATE 为空');
        }

        const [daysAgo, today] = getDateRange();
        const search = [
            { k: 'nolabels',         v: '特殊单,统计排除标', c: '@like', t: '' },
            { k: 'A.seller_flag',    v: '-5',               c: '@=',    t: '' },
            { k: 'cost_type',        v: '1',                c: '@=',    t: '' },
            { k: 'A.status',         v: 'MERGED,SPLIT',     c: '@!=',   t: '' },
            { k: 'C.afterstatus',    v: 'CONFIRMED',        c: '@=',    t: '' },
            { k: 'A.shop_id',        v: currentShopID,      c: '@=',    t: '' },
            { k: 'combinesku_type',  v: 2,                  c: '@=',    t: '' },
            { k: 'combinesku',       v: 1,                  c: '@=',    t: '' },
            { k: 'C.sent_flag',      v: '1',                c: '@=',    t: '' },
            { k: 'A.send_date',      v: daysAgo,            c: '>=',    t: 'date' },
            { k: 'export_date_begin',v: daysAgo,            c: '>=' },
            { k: 'A.send_date',      v: today,              c: '<',     t: 'date' },
            { k: 'export_date_end',  v: today,              c: '<' },
        ];

        let data      = [];
        let page      = 1;
        let pageCount = 5;
        let dataCount = '';

        while (page <= pageCount) {
            const ts  = new Date().getTime();
            const url = `${window.location.href}&ts___=${ts}&am___=LoadDataToJSON`;
            const callbackParam = {
                Method:      'LoadDataToJSON',
                Args:        [String(page), '', '{"fld":"销售数量","type":"desc"}'],
                CallControl: '{page}',
            };

            const formData = new URLSearchParams({
                __VIEWSTATE:          viewState,
                __VIEWSTATEGENERATOR: viewStateGen,
                search:               JSON.stringify(search),
                dataPageCount:        dataCount,
                __CALLBACKID:         'ACall1',
                __CALLBACKPARAM:      JSON.stringify(callbackParam),
            });
            formData.append('column_name', '日期');
            formData.append('column_name', '款式编码');

            try {
                const response = await fetch(url, {
                    method:  'POST',
                    headers: {
                        'content-type':     'application/x-www-form-urlencoded; charset=UTF-8',
                        'x-requested-with': 'XMLHttpRequest',
                    },
                    body: formData,
                });

                const text      = await response.text();
                const jsonStart = text.indexOf('{');

                // 防御：找不到 JSON 起始位置时提前失败，避免误解析
                if (jsonStart === -1) {
                    wordsEl.innerHTML = `[${shopNameById(currentShopID)}] 第${page}页响应中未找到有效 JSON`;
                    console.error('[fetchSearchData] 响应内容:', text);
                    return Promise.reject('响应中未找到有效 JSON');
                }

                const jsonStr = text.substring(jsonStart);

                try {
                    const outer    = JSON.parse(jsonStr);
                    let parsed     = JSON.parse(outer.ReturnValue);
                    const pageData = parsed.datas;
                    pageCount  = parsed.dp.PageCount;
                    dataCount  = String(parsed.dp.DataCount);
                    parsed     = null;

                    if (!pageData || pageData.length === 0) {
                        wordsEl.innerHTML = `[${shopNameById(currentShopID)}] 第${page}页没有返回数据`;
                        page++;
                        await randomDelay();
                        continue;
                    }

                    // 合并数据，排除每页最后一条（汇总行）
                    data = data.concat(pageData.slice(0, -1));
                } catch (parseErr) {
                    wordsEl.innerHTML = `[${shopNameById(currentShopID)}] 解析 JSON 失败，请查看控制台`;
                    console.error('[fetchSearchData] 解析错误:', parseErr);
                    return Promise.reject('解析 JSON 失败');
                }

            } catch (fetchErr) {
                wordsEl.innerHTML = `${shopNameById(currentShopID)} 请求发送失败`;
                console.error('[fetchSearchData] fetch 错误:', fetchErr);
                return Promise.reject('请求发送失败');
            }

            if (page >= pageCount) {
                wordsEl.innerHTML = `[${shopNameById(currentShopID)}] 总计获取 ${data.length} 条数据`;
                return Promise.resolve(data);
            }

            wordsEl.innerHTML = `[${shopNameById(currentShopID)}] 当前页: ${page}, 总页数: ${pageCount} 数据：${data.length}/${dataCount}。`;
            page++;
            await randomDelay();
        }

        // while 正常退出（page > pageCount 但未走 resolve 分支）时也返回数据
        return data;
    }

    async function fetchStyleId(style,shopId) {
        if (typeof(style) != "object" || typeof(shopId) != "string") {return -1};
        if (style.length <= 0 || shopId == "") {return -1};
        
        const url = 'https://goods.scm121.com/api/goods/shopManage/queryShopItemSpuList';
        const header = {
            "Authorization" : `eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJ1aWQiOiIyMTkzNzQ3MyIsInJvbGVJZHMiOltdLCJ1c2VyX25hbWUiOiIxNTY1ODc2MTU5NSIsImNvSWQiOiIxMDE5MjIwNSIsImV4cGlyYXRpb24iOjE3NzM3NjU1ODQzOTksInVzZXIiOnsiY29JZCI6IjEwMTkyMjA1IiwiY29OYW1lIjoi5rip5bee5b6u5oOcIiwibG9naW5OYW1lIjoiMTU2NTg3NjE1OTUiLCJsb2dpbldheSI6IlVOSU9OIiwibmlja05hbWUiOiLpu4TpnZ7lh6ExIiwicm9sZUlkcyI6IjEyLDEzLDE0LDE4LDI3LDI4LDI5LDMwLDMyLDMzLDM1LDM2LDQwLDQxIiwidWlkIjoiMjE5Mzc0NzMifSwiYXV0aG9yaXRpZXMiOlsiSlNULWNoYW5uZWwiLCJKU1Qtc3VwcGxpZXIiLCJtdWx0aUxvZ2luIl0sImNsaWVudF9pZCI6InBjIiwianRpIjoiMTczZDc4OGItYjMxNC00ODY2LWE3Y2EtNDA5ZWFmYTE4NzFlIiwiZXhwIjoxNzczNzY1NTg0fQ.b7ysR-dOHOlaRRRAqVXRyM4E7yiFVa-TqJ1hrlM6JCo`,
            "content-Type" : `application/json;charset=UTF-8`,
            "Gwfp" : "2f584d6593ab8409cb000809bbf5f2e5",
            "Origin" : "https://goods.scm121.com",
            "referer" : "https://goods.scm121.com/manage/goods/onlineStoreGoods/index"
        };
        let styleIdList = {};

        let i = 1;
        for(let item of style){
            wordsEl.innerHTML = `[${i}/${style.length}]查询${item}的ID中`;
            const body = {
                "enabledStatus" : "ON_SALE",
                "pageNum" : 1,
                "pageSize" : 500,
                "salesType" : null,
                "searchType" : 1,
                "sevenSaleSort" : "DESC",
                "shopIds": [shopId],
                "styleCodeVague": "1",
                "styleCodes": item
            }

            try{
                const response = await fetch(url,{
                    method:  'POST',
                    headers: header,
                    body: JSON.stringify(body),
                });

                const responseData = await response.json();
                const data = responseData["data"][0];
                styleIdList[item] = data['shopStyleCode'];
            } catch (error){
                console.warn("[获取款式ID]发送fetch出现错误",error);
                styleIdList[item] = '-';
            }
            i += 1;
            await randomDelay();
        }
        return styleIdList;
    }

    // ─── UI 构建 ─────────────────────────────────────────────────────────────────

    function createWordsSpan() {
        const span = document.createElement('span');
        span.style.cssText = 'position:absolute; left:250px; top:16px;';
        return span;
    }

    function createBtn() {
        const span = document.createElement('span');
        span.classList.add('btn_line', 'h30');
        span.style.cssText = 'position:absolute; left:120px; top:10px;';
        span.title = '将过去三天销量连续增长的款式推送到飞书多维表格。';
        span.innerHTML = '<span class="excel"><span>推送到飞书</span></span>';
        return span;
    }

    function createTestBtn() {
        const span = document.createElement('span');
        span.classList.add('btn_line','h30');
        span.style.cssText = 'position:absolute; left:120px; top:10px;';
        span.title = '测试按钮';
        span.innerHTML = '<span class="excel"><span>测试按钮</span></span>';
        return span;
    }

    /** 创建查询面板（遮罩 + 卡片） */
    function createQueryPanel() {
        const el = (tag, css) => {
            const e = document.createElement(tag);
            if (css) e.style.cssText = css;
            return e;
        };

        const overlay = el('div', `
            position:fixed; top:0; left:0; width:100%; height:100%;
            background-color:rgba(0,0,0,0.5); z-index:9999;
            display:flex; align-items:center; justify-content:center;
        `);

        const panel = el('div', `
            background-color:#fff; border-radius:4px;
            box-shadow:0 2px 12px rgba(0,0,0,0.15);
            width:400px; max-height:80vh; overflow-y:auto; z-index:10000;
        `);

        const header = el('div', `
            padding:16px; border-bottom:1px solid #e8e8e8;
            font-size:16px; font-weight:bold; color:#333;
        `);
        header.textContent = '查询过去三天销量连续增长的款式';

        const content = el('div', 'padding:16px;');

        const blockTitle = el('div', `
            font-size:14px; font-weight:bold; color:#333; margin-bottom:12px;
        `);
        blockTitle.textContent = '查询店铺';

        const checkboxContainer = el('div', `
            display:flex; flex-direction:column; gap:8px; margin-bottom:16px;
        `);

        const checkboxes = {};
        for (const shopKey of Object.keys(SHOPS)) {
            const label = el('label', `
                display:flex; align-items:center; cursor:pointer;
                font-size:14px; color:#333; gap:8px;
            `);
            const checkbox = el('input', 'cursor:pointer; width:16px; height:16px;');
            checkbox.type  = 'checkbox';
            checkbox.value = shopKey;
            checkboxes[shopKey] = checkbox;

            const span = document.createElement('span');
            span.textContent = SHOPS[shopKey].Name;

            label.append(checkbox, span);
            checkboxContainer.appendChild(label);
        }

        content.append(blockTitle, checkboxContainer);

        const footer = el('div', `
            padding:12px 16px; border-top:1px solid #e8e8e8;
            display:flex; gap:8px; justify-content:flex-end;
        `);

        const queryBtn = el('button', `
            padding:8px 16px; background-color:#409EFF; color:white;
            border:none; border-radius:4px; cursor:pointer;
            font-size:14px; transition:background-color 0.3s;
        `);
        queryBtn.textContent = '查询';
        queryBtn.onmouseover = () => queryBtn.style.backgroundColor = '#66B1FF';
        queryBtn.onmouseout  = () => queryBtn.style.backgroundColor = '#409EFF';

        const closeBtn = el('button', `
            padding:8px 16px; background-color:#F5F7FA; color:#333;
            border:1px solid #DCDFE6; border-radius:4px; cursor:pointer;
            font-size:14px; transition:all 0.3s;
        `);
        closeBtn.textContent = '关闭';
        closeBtn.onmouseover = () => closeBtn.style.backgroundColor = '#F0F2F5';
        closeBtn.onmouseout  = () => closeBtn.style.backgroundColor = '#F5F7FA';

        footer.append(queryBtn, closeBtn);
        panel.append(header, content, footer);
        overlay.appendChild(panel);

        const closePanel = () => {
            Object.values(checkboxes).forEach(cb => { cb.checked = false; });
            if (document.body.contains(overlay)) {
                document.body.removeChild(overlay);
            }
        };

        closeBtn.addEventListener('click', closePanel);

        /** 统一恢复按钮状态 */
        const resetBtns = () => {
            queryBtn.disabled    = false;
            queryBtn.textContent = '查询';
            closeBtn.disabled    = false;
        };

        queryBtn.addEventListener('click', async () => {
            const selectedShops = Object.keys(checkboxes).filter(k => checkboxes[k].checked);

            if (selectedShops.length === 0) {
                alert('请至少选择一个店铺');
                return;
            }

            queryBtn.disabled    = true;
            queryBtn.textContent = '查询中...';
            closeBtn.disabled    = true;
            closePanel();

            try {
                for (const shopKey of selectedShops) {
                    currentShopID = SHOPS[shopKey].ID;
                    const rawData = await fetchSearchData();

                    const grouped = {};
                    for (const item of rawData) {
                        const code = String(item.款式编码);
                        const ts   = new Date(item.日期).getTime();
                        const qty  = Number(item.销售数量);
                        (grouped[code] = grouped[code] || []).push([ts, qty]);
                    }

                    const checkedData = cleanData(grouped);

                    if (Object.keys(checkedData).length === 0) {
                        wordsEl.innerHTML = `${SHOPS[shopKey].Name} 没有数据`;
                        continue;
                    }

                    const msgText  = `【${SHOPS[shopKey].Name}】\n${formatMarkdownTable(checkedData)}`;
                    const xlsxRows = await formatXlsxRows(checkedData, SHOPS[shopKey].Name);

                    if (!msgText.trim() || xlsxRows.length === 0) {
                        wordsEl.innerHTML = `${SHOPS[shopKey].Name} 数据格式异常，跳过推送`;
                        continue;
                    }

                    console.log(msgText);

                    try {
                        await sendFeishuMessage(msgText);
                        wordsEl.innerHTML = `${SHOPS[shopKey].Name} 快揽推送成功`;
                        await sendXlsxToFeishu(xlsxRows);
                        wordsEl.innerHTML = `${SHOPS[shopKey].Name} 表格推送成功`;
                    } catch (pushErr) {
                        wordsEl.innerHTML = `${SHOPS[shopKey].Name} 推送失败: ${pushErr.message || pushErr}`;
                        console.error('[推送失败]', pushErr);
                    }

                    if (shopKey !== selectedShops[selectedShops.length - 1]) {
                        await randomDelay();
                    }
                }
                wordsEl.innerHTML = '所有选中的店铺推送完毕。';
            } catch (err) {
                console.error('查询出错:', err);
                wordsEl.innerHTML = `查询出错: ${err.message || err}`;
                alert('查询出错，请查看控制台');
            } finally {
                // 无论成功还是失败，都恢复按钮状态
                resetBtns();
            }
        });

        return overlay;
    }

    // ─── 初始化 ──────────────────────────────────────────────────────────────────

    function init() {
        const btnBar = document.querySelector('#form1 > div.rpt.btn_bar');

        if (!btnBar) {
            console.error('[推送插件] 找不到按钮栏，初始化终止');
            return;
        }

        
        if (!TEST_MODE){
            btnBar.append(btnEl, wordsEl);
            btnEl.addEventListener('click', () => {
                document.body.appendChild(createQueryPanel());
            });

        } else {
            console.log('[推送插件]测试模式中...');
            btnBar.append(testBtnEl, wordsEl)
            testBtnEl.addEventListener('click', () => {
                fetchStyleId(["FA6015","FA23688"],"10270308");
            })
        }
    }

    window.addEventListener('load', () => {
        if (!window.location.href.includes('multidimension.aspx')) return;
        console.log('[推送插件] 查询连续三天增长款式功能已载入页面。');
        init();
    });
})();
