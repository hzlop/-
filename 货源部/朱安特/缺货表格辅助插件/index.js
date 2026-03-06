// ==UserScript==
// @name         采购表格辅助插件
// @namespace    http://tampermonkey.net/
// @version      2.0
// @match        https://*.erp321.com/*
// @grant        GM_xmlhttpRequest
// @connect      ivxubxfk0fm.feishu.cn
// @connect      www.erp321.com
// @connect      src.erp321.com
// @connect      apiweb.erp321.com
// @run-at       document-end
// @updateURL    https://raw.githubusercontent.com/hzlop/-/main/货源部/朱安特/缺货表格辅助插件/index.js
// @downloadURL  https://raw.githubusercontent.com/hzlop/-/main/货源部/朱安特/缺货表格辅助插件/index.js
// ==/UserScript==

(function () {
    'use strict';

    // ─────────────────────────────────────────────
    // 常量配置
    // ─────────────────────────────────────────────

    const WEBHOOK_URL = "https://ivxubxfk0fm.feishu.cn/base/automation/webhook/event/GZBSaqxZyw2ZBmhzm0mcBdW3nUf";
    const WEBHOOK_TOKEN = "Bearer pbwDi1_udjUJgTUEmmyNpcHe";

    const PURCHASE_URL = (ts) =>
        `https://www.erp321.com/app/scm/PurchaseSalesMerge/SalesPurchaseSuggest.aspx?_c=jst-epaas&ts___=${ts}&am___=LoadDataToJSON`;

    const ORDER_URL = (ts) =>
        `https://www.erp321.com/app/order/order/list.aspx?_c=jst-epaas&ts___=${ts}&am___=LoadDataToJSON`;

    const SRC_API_URL =
        "https://apiweb.erp321.com/webapi/ItemApi/ItemSkuIm/GetPageListV2?__from=web_component&owner_co_id=10192205&authorize_co_id=10192205";

    /** 建议采购数门槛：>= 此值的款式单独列行 */
    const HIGH_QTY_THRESHOLD = 20;

    /**
     * 各店铺分组配置
     * key:       内部统计标识
     * shopName:  用于 ERP 表单的 shop_name 参数（逗号分隔）
     * shopIds:   用于 CALLBACKPARAM 中的 shop_id（逗号分隔）
     * label:     日志/统计展示名
     */
    const SHOP_CONFIGS = [
        {
            key: 'FAN',
            label: '法澳娜',
            shopName: "法澳娜得物,法澳娜得物（现货）,法澳娜女装旗舰店,法澳娜拼多多,法澳娜旗舰店,法澳娜旗舰店（小红书）",
            shopIds: "16331577,16346064,14120250,15425438,10270308,14579746"
        },
        {
            key: 'XLT',
            label: '夏丽塔',
            shopName: "夏丽塔官方旗舰店,夏丽塔女装（抖店）,夏丽塔女装(视频号）,夏丽塔女装旗舰店,夏丽塔旗舰店,夏丽塔旗舰店（快手）",
            shopIds: "16193344,14164163,18327577,14349056,10476202,18403642"
        },
        {
            key: 'HYH',
            label: '韩一禾',
            shopName: "hamyiho韩一禾,HAMYIHO韩一禾的店,HAMYIHO韩一禾品质女装",
            shopIds: "18023884,18257506,18257256"
        },
        {
            key: 'YYS',
            label: '朝九晚五衣研社',
            shopName: "朝九晚五衣研社",
            shopIds: "20110469"
        }
    ];

    // ─────────────────────────────────────────────
    // 工具函数
    // ─────────────────────────────────────────────

    /** 随机延迟，默认 2500~4000ms */
    function randomDelay(min = 2500, max = 4000) {
        const ms = Math.floor(Math.random() * (max - min + 1)) + min;
        return new Promise(resolve => setTimeout(resolve, ms));
    }

    /** 固定延迟 */
    function delay(ms) {
        return new Promise(resolve => setTimeout(resolve, ms));
    }

    /**
     * 解析 ERP 响应文本，提取 JSON 主体
     * @param {string} text 原始响应文本
     * @returns {{ datas: Array, sumData: Object }}
     */
    function parseErpResponse(text) {
        const start = text.indexOf('{');
        if (start === -1) throw new Error('响应数据格式错误：未找到 JSON 起始位置');
        const outer = JSON.parse(text.substring(start));
        const inner = JSON.parse(outer['ReturnValue']);
        return {
            datas: inner['datas'] || [],
            sumData: inner['sumData'] || {}
        };
    }

    /**
     * 将 GM_xmlhttpRequest 包装为 Promise
     * @param {Object} options GM_xmlhttpRequest 参数
     * @returns {Promise<Object>} response 对象
     */
    function gmRequest(options) {
        return new Promise((resolve, reject) => {
            GM_xmlhttpRequest({
                ...options,
                onload: (res) => {
                    if (res.status >= 200 && res.status < 300) {
                        resolve(res);
                    } else {
                        reject(new Error(`HTTP ${res.status}: ${res.statusText}`));
                    }
                },
                onerror: (err) => reject(new Error(`网络错误: ${JSON.stringify(err)}`)),
                onabort: () => reject(new Error('请求被中止')),
                ontimeout: () => reject(new Error('请求超时'))
            });
        });
    }

    /**
     * 构建采购建议接口的 CALLBACKPARAM
     * @param {string} shopIds 逗号分隔的 shop_id，空字符串表示全部
     * @param {string} page 页码占位符，默认 "{page}"
     */
    function buildPurchaseCallbackParam(shopIds = '', page = '{page}') {
        const filter = JSON.stringify({
            sku_id: "", i_id: "", name: "", wms_co_ids: "",
            qty_type: "suggest_purchase_qty", qty_begin: "1", qty_end: "",
            supplier_group: "", supplier_id: "", supplier_handle_type: "",
            supplier_sku_id: "", supplier_i_id: "",
            shop_id: shopIds,
            not_shop_id: "", vc_name_type: "包含分类", vc_name: "",
            properties_value: "", question_type: "", no_question_type: "",
            drp_co_id: "", goods_type: "", brand: "", input_category: "",
            remark: "", enabled: "", mulriple: "", ishasBomMap: "",
            labels: "", nolabels: "特殊单", skulabels: "", noskulabels: "",
            node: "", sku_code: "", begin_order_date: "", end_order_date: "",
            no_exist_sku: false, fast_type: ""
        });
        return {
            Method: "LoadDataToJSON",
            Args: ["1", "1000", filter, '{"fld":"","type":"asc"}', "sku"],
            CallControl: page
        };
    }

    /**
     * 获取页面 ViewState 等表单隐藏域，获取失败时使用备用硬编码值
     */
    function getViewState() {
        const viewState = document.querySelector('#__VIEWSTATE')?.value;
        const viewStateGenerator = document.querySelector('#__VIEWSTATEGENERATOR')?.value;
        const eventValidation = document.querySelector('#__EVENTVALIDATION')?.value;

        if (viewState && eventValidation) {
            return { viewState, viewStateGenerator: viewStateGenerator || '', eventValidation };
        }
        // 备用硬编码值（页面 DOM 尚未就绪时的兜底）
        return {
            viewState: '/wEPDwULLTIxMDgyNDUzMzhkZID/fTXtwe65bgIGeK8EbTQKAWiI',
            viewStateGenerator: '549B40FC',
            eventValidation: ''
        };
    }

    /**
     * 构建采购建议接口的完整 FormData
     * @param {string} shopName  shop_name 参数
     * @param {string} shopIds   shop_id（用于 CALLBACKPARAM）
     */
    function buildPurchaseFormData(shopName, shopIds) {
        const { viewState, viewStateGenerator, eventValidation } = getViewState();
        const callbackParam = buildPurchaseCallbackParam(shopIds);

        const fields = {
            __VIEWSTATE: viewState,
            __VIEWSTATEGENERATOR: viewStateGenerator,
            search: "", begin_order_date: "", end_order_date: "",
            skuid: "", i_id: "", name: "", properties_value: "",
            shop_name: shopName, not_shop_name: "", labels: "",
            nolabels: "特殊单", order_include: "", order_oper: "", node: "",
            input_category: "", category_names: "", vc_name: "", brand: "",
            brand_name: "", ishasBomMap: "", sku_code: "", supplier_group: "",
            supplier_name: "", supplier_handle_type: "", supplier_sku_id: "",
            supplier_i_id: "", labels_search: "", labels_exclude_search: "",
            enabled: "", mulriple: "", remark: "", wms_co_ids_v: "",
            wms_co_ids: "", drp_co_id_v: "", drp_co_id: "", goods_type: "",
            qty_type: "suggest_purchase_qty", qty_begin: "1", qty_end: "",
            _cbb_vc_name: "", _cbb_brand_name: "", _cbb_wms_co_ids: "",
            _cbb_drp_co_id: "", no_exist_sku: "",
            __CALLBACKID: "ACall1",
            __CALLBACKPARAM: JSON.stringify(callbackParam),
            __EVENTVALIDATION: eventValidation
        };

        const form = new URLSearchParams();
        Object.entries(fields).forEach(([k, v]) => form.append(k, v));
        return form;
    }

    /**
     * 构建订单列表接口的 FormData
     * @param {Object} callbackParam
     */
    function buildOrderFormData(callbackParam) {
        const fields = {
            __VIEWSTATE: "/wEPDwUKLTk3OTc4NTg5MWRkAAjb5TI3+T+7YJtzHDn1bxVwpEg=",
            __VIEWSTATEGENERATOR: "C8154B07",
            insurePrice: "", _jt_page_count_enabled: "",
            _jt_page_increament_enabled: "true", _jt_page_increament_page_mode: "",
            _jt_page_increament_key_value: "", _jt_page_increament_business_values: "",
            _jt_page_increament_key_name: "o_id", _jt_page_size: "2000",
            _jt_page_action: "1", fe_node_desc: "", receiver_state: "",
            receiver_city: "", receiver_district: "", receiver_address: "",
            receiver_name: "", receiver_phone: "", receiver_mobile: "",
            check_name: "", check_address: "", fe_remark_type: "single",
            fe_flag: "", fe_is_append_remark: "",
            __CALLBACKID: "JTable1",
            __CALLBACKPARAM: JSON.stringify(callbackParam)
        };

        const form = new URLSearchParams();
        Object.entries(fields).forEach(([k, v]) => form.append(k, v));
        return form;
    }

    // ─────────────────────────────────────────────
    // 数据抓取层
    // ─────────────────────────────────────────────

    const PURCHASE_HEADERS = {
        'content-type': 'application/x-www-form-urlencoded; charset=UTF-8',
        'x-requested-with': 'XMLHttpRequest'
    };

    /**
     * 抓取一个店铺分组的建议采购汇总数量
     * @param {string} shopName
     * @param {string} shopIds
     * @returns {Promise<number>} suggest_purchase_qty 合计
     */
    async function fetchPurchaseSummary(shopName, shopIds) {
        const res = await fetch(PURCHASE_URL(Date.now()), {
            method: 'POST',
            headers: PURCHASE_HEADERS,
            body: buildPurchaseFormData(shopName, shopIds)
        });
        if (!res.ok) throw new Error(`HTTP ${res.status}: ${res.statusText}`);
        const { sumData } = parseErpResponse(await res.text());
        return Number(sumData['suggest_purchase_qty']) || 0;
    }

    /**
     * 抓取全部数据的详细列表（无 shop 过滤）
     * @returns {Promise<{ datas: Array, sumData: Object }>}
     */
    async function fetchAllPurchaseData() {
        const res = await fetch(PURCHASE_URL(Date.now()), {
            method: 'POST',
            headers: PURCHASE_HEADERS,
            body: buildPurchaseFormData('', '')
        });
        if (!res.ok) throw new Error(`HTTP ${res.status}: ${res.statusText}`);
        return parseErpResponse(await res.text());
    }

    /**
     * 顺序抓取全量数据 + 各分组数量
     * @returns {Promise<{ datas: Array, counts: Object }>}
     */
    async function fetchAllData() {
        console.log('[采购插件] 开始抓取全量数据...');
        const { datas, sumData } = await fetchAllPurchaseData();
        const counts = { all: Number(sumData['suggest_purchase_qty']) || 0 };

        for (const cfg of SHOP_CONFIGS) {
            console.log(`[采购插件] 正在抓取 ${cfg.label}...`);
            await randomDelay();
            counts[cfg.key] = await fetchPurchaseSummary(cfg.shopName, cfg.shopIds);
        }

        console.log('[采购插件] 全部数据抓取完毕', counts);
        return { datas, counts };
    }

    /**
     * 查询无货留言订单数量
     * @returns {Promise<number>}
     */
    async function fetchOutOfStockCount() {
        const callbackParam = {
            Method: "LoadDataToJSON",
            Args: ["1", '[{"k":"remark_note","v":"无货留言","c":"like"}]', "{}"]
        };
        try {
            const res = await fetch(ORDER_URL(Date.now()), {
                method: 'POST',
                headers: {
                    referer: "https://www.erp321.com/app/order/order/list.aspx?_c=jst-epass",
                    'Content-Type': 'application/x-www-form-urlencoded; charset=UTF-8',
                    'X-Requested-With': 'XMLHttpRequest'
                },
                body: buildOrderFormData(callbackParam)
            });
            if (!res.ok) return 0;
            const { datas } = parseErpResponse(await res.text());
            return datas.length;
        } catch (e) {
            console.error('[采购插件] 获取无货留言失败:', e);
            return 0;
        }
    }

    /**
     * 查询指定款式编码的缺货订单，返回 { 商品编码: 主要缺货店铺 }
     * @param {string[]} itemIds
     * @returns {Promise<Object>}
     */
    async function fetchOutOfStockShops(itemIds) {
        if (!itemIds || itemIds.length === 0) return {};

        const callbackParam = {
            Method: "LoadDataToJSON",
            Args: [
                "1",
                `[{"k":"i_id","v":"${itemIds.join(',')}","c":"="},{"k":"status","v":"question","c":"@="},{"k":"question_type","v":"缺货","c":"@="}]`,
                "{}"
            ]
        };

        try {
            const res = await fetch(ORDER_URL(Date.now()), {
                method: 'POST',
                headers: {
                    referer: "https://www.erp321.com/app/order/order/list.aspx?_c=jst-epass",
                    'Content-Type': 'application/x-www-form-urlencoded; charset=UTF-8',
                    'X-Requested-With': 'XMLHttpRequest'
                },
                body: buildOrderFormData(callbackParam)
            });
            if (!res.ok) return {};
            const { datas } = parseErpResponse(await res.text());
            return computeTopShopPerItem(datas, itemIds);
        } catch (e) {
            console.error('[采购插件] 获取缺货店铺失败:', e);
            return {};
        }
    }

    /**
     * 统计每个款式编码对应缺货最多的店铺
     * @param {Array} orders 订单列表
     * @param {string[]} itemIds 目标款式编码列表
     * @returns {Object} { 商品编码: 主要缺货店铺名 }
     */
    function computeTopShopPerItem(orders, itemIds) {
        const shopQtyMap = {}; // { i_id: { shop_name: qty } }

        for (const order of orders) {
            const shopName = order['shop_name'] || '-';
            for (const sku of (order['items'] || [])) {
                const iid = sku['i_id'];
                if (!itemIds.includes(iid)) continue;
                if (!shopQtyMap[iid]) shopQtyMap[iid] = {};
                shopQtyMap[iid][shopName] = (shopQtyMap[iid][shopName] || 0) + (Number(sku['qty']) || 1);
            }
        }

        const result = {};
        for (const [iid, shopMap] of Object.entries(shopQtyMap)) {
            result[iid] = Object.entries(shopMap).reduce(
                (best, [shop, qty]) => qty > best.qty ? { shop, qty } : best,
                { shop: '', qty: 0 }
            ).shop;
        }
        return result;
    }

    /**
     * 通过 src API 查询款式的供应商备注和虚拟分类
     * @param {string[]} itemIds
     * @returns {Promise<Object>} { i_id: { vc_name, remark } }
     */
    async function fetchItemSrcData(itemIds) {
        const body = JSON.stringify({
            ip: "", uid: "21937473", coid: "10192205",
            page: { currentPage: 1, pageSize: 500 },
            data: {
                queryFlds: [
                    "pic", "i_id", "name", "s_price", "market_price", "c_price",
                    "brand", "c_name", "vc_name", "supplier_i_id", "weight",
                    "l", "w", "h", "volume", "unit", "onsale", "remark",
                    "created", "modified", "c_id"
                ],
                orderBy: "", onsale: "1",
                i_id: itemIds.join(','),
                sku_id: "", c_id: ""
            }
        });

        const res = await gmRequest({
            method: 'POST',
            url: SRC_API_URL,
            headers: {
                'content-type': 'application/json; charset=utf-8',
                'gwfp': '8a9d59f44e15cc9d2be4acd656e93905',
                'priority': 'u=1, i',
                'u_sso_token': '',
                'origin': 'https://src.erp321.com',
                'referer': 'https://src.erp321.com/',
                'webbox-request-id': 'be0818c2-51e9-4191-83c7-ca4c4c723b05',
                'webbox-route-path': '/erp-components/goods-selector/'
            },
            data: body
        });

        const parsed = JSON.parse(res.responseText);
        const result = {};
        for (const item of (parsed['data'] || [])) {
            result[item['i_id']] = { vc_name: item['vc_name'] || '', remark: item['remark'] || '' };
        }
        return result;
    }

    // ─────────────────────────────────────────────
    // 数据处理层
    // ─────────────────────────────────────────────

    /** 将虚拟分类"成品采购"规范化为"唐杰" */
    function normalizeCategory(cat) {
        return (cat === '成品采购') ? '唐杰' : (cat || '');
    }

    /**
     * 将原始 dataList 聚合为飞书多维表格所需的行列表
     * @param {Array} dataList  ERP 建议采购列表
     * @param {Object} counts   各分组 suggest_purchase_qty 汇总
     * @returns {Promise<Array>}
     */
    async function buildFeishuRows(dataList, counts) {
        if (!dataList || dataList.length === 0) {
            console.warn('[采购插件] 数据列表为空');
            return [];
        }

        // 1. 按款式编码合并采购数量
        const productMap = {}; // { i_id: { productCode, purchaseQty, category } }
        for (const item of dataList) {
            const iid = item['i_id'];
            const qty = Number(item['suggest_purchase_qty']) || 0;
            const cat = normalizeCategory(item['td_vc_name']);

            if (productMap[iid]) {
                productMap[iid].purchaseQty += qty;
            } else {
                productMap[iid] = { productCode: iid, purchaseQty: qty, category: cat };
            }
        }

        // 2. 分流：高数量款式 vs. 按负责人汇总
        const highQtyItems = {};  // { i_id: { 商品编码, 建议采购数, 虚拟分类 } }
        const summaryQty = { 方孙俊: 0, 王苗苗: 0, 唐杰: 0 };

        for (const { productCode, purchaseQty, category } of Object.values(productMap)) {
            if (purchaseQty >= HIGH_QTY_THRESHOLD) {
                highQtyItems[productCode] = {
                    '商品编码': productCode,
                    '建议采购数': purchaseQty,
                    '虚拟分类': category
                };
            } else {
                if (summaryQty[category] !== undefined) {
                    summaryQty[category] += purchaseQty;
                }
            }
        }

        // 3. 并行拉取辅助数据
        const highQtyIds = Object.keys(highQtyItems);
        const outOfStockCount = await fetchOutOfStockCount();

        let srcData = {};
        let outOfStockShops = {};
        if (highQtyIds.length > 0) {
            await randomDelay();
            [outOfStockShops, srcData] = await Promise.all([
                fetchOutOfStockShops(highQtyIds),
                fetchItemSrcData(highQtyIds)
            ]);
        }

        // 4. 构建行列表
        const makeRow = (款号, 缺货数量, 负责人 = '', 主要加工厂 = '', 主要店铺 = '') => ({
            款号, 真实库存数: 0, 缺货数量, 到货时间: '', 主要加工厂, 主要店铺, 负责人
        });

        const rows = [];

        // 汇总行（小量款式）
        rows.push(makeRow('其他', summaryQty['方孙俊'], '方孙俊'));
        rows.push(makeRow('其他', summaryQty['王苗苗'], '王苗苗'));
        rows.push(makeRow('成品采购', summaryQty['唐杰'], '唐杰'));
        rows.push(makeRow('无货留言', outOfStockCount, '朱安特'));

        // 高数量款式明细行
        let highQtyTotal = 0;
        for (const item of Object.values(highQtyItems)) {
            const iid = item['商品编码'];
            const src = srcData[iid] || { vc_name: '', remark: '' };
            const superintendent = normalizeCategory(src.vc_name) || item['虚拟分类'];
            const isExternal = (item['虚拟分类'] === '唐杰' || item['虚拟分类'] === '成品采购');
            const remark = isExternal ? '外采' : src.remark;

            highQtyTotal += item['建议采购数'];
            rows.push(makeRow(iid, item['建议采购数'], superintendent, remark, outOfStockShops[iid] || ''));
        }

        // 欠单统计汇总行
        const shopTotal = counts.FAN + counts.XLT + counts.HYH + counts.YYS;
        const totalOwed =
            (summaryQty['方孙俊'] + summaryQty['王苗苗'] + summaryQty['唐杰'] +
                outOfStockCount + highQtyTotal) - shopTotal;

        rows.push(makeRow('欠单数量', totalOwed));
        rows.push(makeRow('韩一禾欠单总数', counts.HYH));
        rows.push(makeRow('夏丽塔欠单总数', counts.XLT));
        rows.push(makeRow('法澳娜欠单总数', counts.FAN));
        rows.push(makeRow('朝九晚五衣研社欠单总数', counts.YYS));

        return rows;
    }

    // ─────────────────────────────────────────────
    // 飞书发送层
    // ─────────────────────────────────────────────

    /**
     * 发送单条 Webhook 请求
     * @param {Object} item
     * @returns {Promise}
     */
    function sendWebhook(item) {
        return gmRequest({
            method: 'POST',
            url: WEBHOOK_URL,
            headers: {
                Authorization: WEBHOOK_TOKEN,
                'Content-Type': 'application/json'
            },
            data: JSON.stringify(item)
        });
    }

    /**
     * 批量发送到飞书，带节流（每条 300ms，每 4 条额外 500ms）
     * @param {Array} rows
     * @param {function} onProgress (sent, total) => void
     */
    async function sendToFeishu(rows, onProgress) {
        if (!rows || rows.length === 0) {
            alert('没有可发送的数据');
            return;
        }
        for (let i = 0; i < rows.length; i++) {
            try {
                await sendWebhook(rows[i]);
            } catch (e) {
                console.error(`[采购插件] 第 ${i + 1} 条发送失败，继续:`, e);
            }
            if (onProgress) onProgress(i + 1, rows.length);
            await delay(300);
            if ((i + 1) % 4 === 0) await delay(500);
        }
        console.log(`[采购插件] 发送完毕，共 ${rows.length} 条`);
    }

    // ─────────────────────────────────────────────
    // UI 层
    // ─────────────────────────────────────────────

    /**
     * 初始化插件 UI（按钮）
     */
    function initUI() {
        const td = document.querySelector(
            '#queryPoolsRegion > div.panel > table > tbody > tr > td:nth-child(1)'
        );
        if (!td) {
            console.warn('[采购插件] 未找到目标 TD，UI 初始化跳过');
            return;
        }

        // 插入分隔符
        td.insertAdjacentHTML('beforeend', `
            <li class="_jt_tool_spt">
                <span class="_jt_tool_spt_bg"></span>
            </li>`
        );

        // 创建按钮
        const wrap = document.createElement('span');
        wrap.className = 'ding_db';
        const btn = document.createElement('span');
        btn.id = 'purchase-helper-btn';
        btn.className = 'ding_db_txt ding_db_txt_hover';
        btn.textContent = '一键发送到多维表格';
        wrap.appendChild(btn);
        td.appendChild(wrap);

        btn.addEventListener('click', async () => {
            if (btn.dataset.busy === '1') return;

            btn.dataset.busy = '1';
            btn.textContent = '抓取数据中...';

            try {
                const { datas, counts } = await fetchAllData();
                btn.textContent = '处理数据中...';

                const rows = await buildFeishuRows(datas, counts);
                btn.textContent = `发送中 0/${rows.length}`;

                await sendToFeishu(rows, (sent, total) => {
                    btn.textContent = `发送中 ${sent}/${total}`;
                });

                btn.textContent = '发送完成';
            } catch (err) {
                console.error('[采购插件] 执行失败:', err);
                alert('执行失败: ' + err.message);
                btn.textContent = '发生错误，可重试';
            } finally {
                delete btn.dataset.busy;
                setTimeout(() => {
                    btn.textContent = '一键发送到多维表格';
                }, 4000);
            }
        });
    }

    // ─────────────────────────────────────────────
    // 入口
    // ─────────────────────────────────────────────

    window.addEventListener('load', () => {
        if (!window.location.href.includes('SalesPurchaseSuggest.aspx')) return;
        initUI();
    });

})();
