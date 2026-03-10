// ==UserScript==
// @name         售后自用插件
// @namespace    http://tampermonkey.net/
// @version      7.1
// @description  重构版本
// @author       达摩
// @match        https://www.erp321.com/*
// @match        https://w.erp321.com/*
// @icon         data:image/gif;base64,R0lGODlhAQABAAAAACH5BAEKAAEALAAAAAABAAEAAAICTAEAOw==
// @grant        GM_xmlhttpRequest
// @updateURL    https://raw.githubusercontent.com/hzlop/-/main/客服部/售后组/售后自用插件/index.js
// @downloadURL  https://raw.githubusercontent.com/hzlop/-/main/客服部/售后组/售后自用插件/index.js
// @run-at       document-end
// ==/UserScript==

(function() {
    'use strict';

    // ==================== 配置 ====================
    const CONFIG = {
        WEBHOOKS: {
            AFTERSALES: "https://ivxubxfk0fm.feishu.cn/base/automation/webhook/event/JAPNaHvxwwr9jNhca2CcQXzlnsd",
            WANDAN: "https://ivxubxfk0fm.feishu.cn/base/automation/webhook/event/YoGba8hJhwYC5IhVlHOced3bnFJ",
            CUOFA: "https://ivxubxfk0fm.feishu.cn/base/automation/webhook/event/TWyjaL215w45aYhdapecmWaIn9b",
            SIYU: "https://ivxubxfk0fm.feishu.cn/base/automation/webhook/event/WTr0aETTqwQ7p7h8c7YcewNYnFb",
            DIUJIAN: "https://ivxubxfk0fm.feishu.cn/base/automation/webhook/event/WrGKazuj3wPHhdhdPunc01vDnUb"
        },
        MESSAGES: {
            SUCCESS: "发送成功!",
            ERROR: "发送失败:",
            SENDING: "发送中...",
            NO_ORDER_LIST: "未找到订单列表节点",
            NO_SELECTED_ROW: "请先选中订单行",
            SELECT_PRODUCT_STYLE: "请选择款式",
            SELECT_OR_INPUT_CONTENT: "请选择或输入内容",
            SELECT_REFUND_TYPE: "请选择退款类型",
            INPUT_RETURN_REASON: "请输入退货原因",
            INPUT_RESULT: "请输入结果",
            NO_ORDER_ID: "未找到订单ID，不要选择多订单号的行！！！",
            NO_DATAS: "未获取到订单数据"
        },

        TARGET_DOMAINS: { // 目标执行域名
            W: 'https://w.erp321.com/app/order/order/list.aspx',
            WWW: 'https://www.erp321.com/app/order/order/list.aspx'
        },
        SHOP_NAME_MAP: {
            "法澳娜旗舰店": "法澳娜",
            "夏丽塔旗舰店": "夏丽塔"
        },
        CONTENT_OPTIONS: ["污渍", "起球", "开线", "次品", "掉毛", "粘毛", "掉色", "破洞", "做工问题",
            "发白", "勾丝", "批次", "异味", "缩水", "色差", "领子变形", "袖子变形", "0"],
        REMARK_OPTIONS: ["", "退款", "补发", "换货", "补偿", "运费"],
        RESPONSIBILITY_OPTIONS: ["公司", "质检", "厂家"],
        INSPECTORS: [
            "合格01", "合格03", "合格02", "合格19", "合格15", "合格09", "合格07", "合格06",
            "合格11", "合格01-06", "合格01-09", "合格17", "合格01-01", "检验33", "检验28",
            "检验27", "检验10", "检验12", "检验32", "检验51", "检验05", "检验35", "检验20",
            "检验06(302款)", "检验39", "检验22", "检验23", "检验18", "检验66", "检验76",
            "检验75", "检验85", "检验98", "检验88", "检验78", "检验68", "检验58", "检验66",
            "检验77", "检验72", "检验71", "检验87", "检验02", "检验08", "检验26", "检验16",
            "检验06", "检验53", "检验55", "检验52", "检验37", "合格05", "合格01-07", "合格01-03",
            "合格01-05", "检验13", "合格01-02", "合格40"
        ],
        AMOUNT_SHORTCUTS: ["5.00", "7.00", "8.00", "10.00", "12.00", "15.00", "20.00"],
        INSPECTOR_SHORTCUTS: ["检验18", "检验08", "检验52", "检验06", "检验39", "合格06"],
        WANDAN: {
            TYPE_OPTIONS: ["退货退款", "换货", "仅退款"],
            RETURN_REASON_SHORTCUTS: ["贵了", "不想要了", "领子紧", "大", "小", "袖子长", "颜色不喜欢"],
            RESULT_SHORTCUTS: ["留下", "收到试试", "再拍一件"],
            NAME: "罗清晨"
        },

        HOTKEYS: {
            TOGGLE_AFTERSALES: 's',
            TOGGLE_WANDAN: 'a',
            TOGGLE_CUOFA: 'f',
            TOGGLE_SIYU: '4',
            TOGGLE_DIUJIAN: 'd',
            CLOSE_ALL: 'Escape'
        },

        SIYU: {
            VEKZ_OPTIONS: ["半价", "清仓", "半价+清仓", "95折"],
            TYPE_OPTIONS: ["空", "晒图", "线下支付"],
            Q_SHORTCUTS: ["19.90", "29.90", "39.90", "49.90", "69.90", "34.90", "65.00"]
        },
        CUOFA: {
            QKKL_OPTIONS: ["无走件", "故意错发", "条形码与实物颜色不符", "发错货", "发错颜色", "编码错误",
                          "条形码与吊牌实物不符", "客服换错货", "条形码和吊牌与实物不符", "条形码吊牌与实物不符",
                          "条形码与实物不符", "条形码与吊牌尺码不符", "发错尺码", "重量正确收到少一件",
                          "收到少一件", "换错货", "吊牌与水洗标尺码不符"],
            QKKL_SHORTCUTS: ["无走件", "发错货", "编码错误", "客服操作失误","故意错发","条形码和吊牌与实物不符", "条形码和吊牌实物不符"],
            ZERF_OPTIONS: ["公司", "厂家", "质检", "出库"],
            DJHC_SHORTCUTS: ["检验77", "检验98", "检验52", "合格06"]
        },
        DIUJIAN: {
            QKKL_OPTIONS: ["快递丢件", "快递破损"]
        }
    };

    // ==================== 工具函数 ====================
    class Utils {
        // 显示提示信息的静态方法
        static showToast(message, duration = 3000) {
            if (!Utils.toastElement) {
                Utils.toastElement = Utils.createToastElement();
            }
            Utils.toastElement.textContent = message;
            Utils.toastElement.classList.add('show');
            setTimeout(() => {
                Utils.toastElement.classList.remove('show');
            }, duration);
        }

        // 创建提示信息元素的静态方法
        static createToastElement() {
            const style = `
                .custom-toast {
                    position: fixed;
                    top: 20px;
                    left: 50%;
                    transform: translateX(-50%);
                    padding: 10px 20px;
                    background: #333;
                    color: #fff;
                    border-radius: 4px;
                    z-index: 9999999;
                    opacity: 0;
                    transition: opacity 0.3s ease;
                    pointer-events: none;
                    font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, sans-serif;
                    font-size: 14px;
                }
                .custom-toast.show {
                    opacity: 1;
                }
            `;
            const styleElement = document.createElement('style');
            styleElement.textContent = style;
            document.head.appendChild(styleElement);

            const toast = document.createElement('div');
            toast.className = 'custom-toast';
            document.body.appendChild(toast);
            return toast;
        }

        // 安全获取DOM节点的静态方法
        static getSafeNode(className, parent = document) {
            const node = parent.querySelector(`.${className}`);
            if (!node) {
                console.warn(`未找到节点: .${className}`);
            }
            return node;
        }

        // 发送网络请求的静态方法
        static sendRequest(url, data, onSuccess, onError) {
            const handleResponse = (status, responseText) => {
                if (status >= 200 && status < 300) {
                    onSuccess(responseText);
                } else {
                    onError(`状态码:${status}\n响应:${responseText}`);
                }
            };

            if (typeof GM_xmlhttpRequest !== 'undefined') {
                GM_xmlhttpRequest({
                    method: "POST",
                    url: url,
                    headers: { "Content-Type": "application/json" },
                    data: JSON.stringify(data),
                    onload: (response) => handleResponse(response.status, response.responseText),
                    onerror: (error) => onError(error.message)
                });
            } else {
                const xhr = new XMLHttpRequest();
                xhr.open("POST", url, true);
                xhr.setRequestHeader("Content-Type", "application/json");
                xhr.onload = () => handleResponse(xhr.status, xhr.responseText);
                xhr.onerror = () => onError("网络请求失败");
                xhr.send(JSON.stringify(data));
            }
        }
        static creakeSearchParams(callbackParam, page_action = false) {
            let formData = new URLSearchParams();
            let VIEWSTATE = window.document.getElementById('__VIEWSTATE')?.value || '/wEPDwULLTE5NDE3OTc0MzhkZFljjZ+j39FvCUEtOmvLM52MuNDh';
            let viewStateGenerator = window.document.getElementById('__VIEWSTATEGENERATOR')?.value || 'C8154B07';
            formData.append('__VIEWSTATE', VIEWSTATE);
            formData.append('__VIEWSTATEGENERATOR', viewStateGenerator);
            formData.append('insurePrice', '');
            formData.append('_jt_page_count_enabled', '');
            formData.append('_jt_page_increament_enabled', true);
            formData.append('_jt_page_increament_page_mode', '');
            formData.append('_jt_page_increament_key_value', '');
            formData.append('_jt_page_increament_business_values', '');
            formData.append('_jt_page_increament_key_name', 'o_id');
            formData.append('_jt_page_size', '2000');
            if (page_action) formData.append('_jt_page_action', '1');
            formData.append('fe_node_desc', '');
            formData.append('receiver_state', '');
            formData.append('receiver_city', '');
            formData.append('receiver_district', '');
            formData.append('receiver_address', '');
            formData.append('receiver_name', '');
            formData.append('receiver_phone', '');
            formData.append('receiver_mobile', '');
            formData.append('check_name', '');
            formData.append('check_address', '');
            formData.append('fe_remark_type', 'single');
            formData.append('fe_flag', '');
            formData.append('fe_is_append_remark', '');
            formData.append('__CALLBACKID', 'JTable1');
            formData.append('__CALLBACKPARAM', JSON.stringify(callbackParam));
            return formData;
        }

        static async GetOrderDataByOrder(Oid) {
            if (!Oid || Oid.length !== 8 || typeof Oid != 'string' || !(/^-?\d+$/.test(Oid))) {
                console.error("请检查内部单号：",Oid);
                return Promise.reject('订单号不正确');
            }
            const timestamp = new Date().getTime();
            const url = `https://www.erp321.com/app/order/order/list.aspx?_c=jst-epaas&ts___=${timestamp}&am___=LoadDataToJSON`;
            const callbackParam = {
                "Method": "LoadDataToJSON",
                "Args": ["1", `[{\"k\":\"o_id\",\"v\":\"${Oid}\",\"c\":\"@=\"}]`, "{}"]
            };
            const formData = Utils.creakeSearchParams(callbackParam, true);
            const headers = {
                "Content-Type": "application/x-www-form-urlencoded; charset=UTF-8",
                "X-Requested-With": "XMLHttpRequest"
            };
            return new Promise((resolve, reject) => {
                // 使用 fetch 发送 POST 请求
                fetch(url, {
                    method: "POST",
                    headers: headers,
                    body: formData,
                    credentials: 'include'
                })
                .then(response => {
                    if (!response.ok) {
                        throw new Error(`HTTP 请求失败，状态码: ${response.status}`);
                    }
                    return response.text();
                })
                .then(text => {
                    let realJsonStr = text.substring(text.indexOf('{'));
                    const Datas = JSON.parse(JSON.parse(realJsonStr)['ReturnValue'])['datas'];

                    if (!Datas || Datas.length === 0) {
                        resolve(null);
                    } else {
                        resolve(Datas[0]);
                    }
                })
                .catch(error => {
                    console.error("订单数据获取失败 data:", error);
                    resolve(null);
                });
            });
        }

        static async SaveAppendRemarks(flag=null,remark,o_id,append = true) {
            const timestamp = new Date().getTime();
            let url = `https://www.erp321.com/app/order/order/list.aspx?_c=jst-epaas&ts___=${timestamp}&am___=SaveAppendRemarks`;
            if (!o_id){
                return Promise.resolve(false);
            };
            let callbackParam = {
                "Method":"SaveAppendRemarks",
                "Args":[flag,remark,o_id,append],
                "callControl":"{page}"
            };
            let formData = Utils.creakeSearchParams(callbackParam, false);
            const headers = {
                "Content-Type": "application/x-www-form-urlencoded; charset=UTF-8",
                "X-Requested-With": "XMLHttpRequest"
            };
            return new Promise((resolve, reject) => {
                GM_xmlhttpRequest({
                    method: "POST",
                    url: url,
                    headers: headers,
                    data: formData,
                    onload: function(response) {
                        console.log("备注保存成功:", response.status);
                        resolve(true);
                    },
                    onerror: function(error) {
                        console.error("备注保存失败:", error);
                        resolve(false);
                    }
                })});
    };
    }


    // ==================== ERP数据提取 ====================
    class ERPDataExtractor {
        // 获取当前选中的订单行
        static getCurrentRow() {
            const rowList = Utils.getSafeNode('rowList');
            if (!rowList) {
                Utils.showToast(CONFIG.MESSAGES.NO_ORDER_LIST);
                return null;
            }

            const currentRow = Utils.getSafeNode('_jt_row_current', rowList);
            if (!currentRow) {
                Utils.showToast(CONFIG.MESSAGES.NO_SELECTED_ROW);
                return null;
            }

            return currentRow;
        }

        // 从订单行获取短单号
        static getOid(row) {
            const osIdNode = row.querySelector('._jt_cell_o_id');
            const Oid = osIdNode.querySelector('.o_id').innerText.trim() || '';
            if (!Oid || Oid === '') return "";
            return Oid || "";
        }

        // 提取订单数据的主方法
        static async extractOrderData() {
            const row = this.getCurrentRow();

            if (!row) {
                Utils.showToast(CONFIG.MESSAGES.NO_SELECTED_ROW);
                return null;
            }

            const datas = await Utils.GetOrderDataByOrder(this.getOid(row));
            if (!datas) {
                Utils.showToast(CONFIG.MESSAGES.NO_DATAS);
                return null;
            }

            let productList = [];

            const items = datas.items;
            
            items.forEach(item => {
                productList.push({
                    productStyle: item.i_id,
                    productCode: item.sku_id,
                    price: item.price,
                    quantity: item.qty
                })
            })

            // 为了保持向后兼容，创建productStyleList和productStylePriceMap
            // 对款式进行去重处理
            const uniqueStyles = new Set();
            const productStylePriceMap = {};

            productList.forEach(product => {
                uniqueStyles.add(product.productStyle);
                if (!productStylePriceMap.hasOwnProperty(product.productStyle)) {
                    productStylePriceMap[product.productStyle] = product.price;
                }
            });

            const productStyleList = Array.from(uniqueStyles);

            return {
                orderId: datas.so_id,
                Oid:datas.o_id,
                shopName: datas.shop_name,
                productList: productList,
                productStyleList: productStyleList,
                productStylePriceMap: productStylePriceMap,
                logisticsCompany: datas.logistics_company,
                trackingNumber: datas.l_id
        };
    }
    }

    // ==================== 面板基类 ====================
    class BasePanel {
        // 构造方法，初始化面板
        constructor(panelId, config) {
            this.panelId = panelId;
            this.config = config;
            this.panel = document.getElementById(panelId);
            this.focusableElements = [];
            this.Oid = '';
        }

        // 显示面板并填充数据
        async show() {
            this.Oid = await this.populate();
            this.panel.classList.add('show');
            if (this.focusableElements.length > 0) {
                this.focusableElements[0].focus();
            }
        }

        // 隐藏面板并重置数据
        hide() {
            this.panel.classList.remove('show');
            this.Oid = '';
            this.reset();
        }

        // 切换面板的显示/隐藏状态
        async toggle() {
            if (this.panel.classList.contains('show')) {
                this.hide();
            } else {
                await this.show();
            }
        }

        // 填充面板数据
        async populate() {
            // 由子类实现
        }

        // 重置面板数据
        reset() {
            // 由子类实现
        }

        // 收集表单数据
        collectData() {
            // 由子类实现
            return null;
        }

        // 发送表单数据到服务器
        send() {
            const data = this.collectData();
            if (!data) return;

            const sendBtn = this.panel.querySelector('.btn-send');
            sendBtn.disabled = true;
            sendBtn.textContent = CONFIG.MESSAGES.SENDING;
            sendBtn.style.backgroundColor = '#9e9e9e';

            // 获取追加备注复选框状态
            const appendRemarksCheckbox = this.getAppendRemarksCheckbox();
            const shouldAppendRemarks = appendRemarksCheckbox && appendRemarksCheckbox.checked;

            Utils.sendRequest(
                this.config.webhookUrl,
                data,
                (response) => {
                    Utils.showToast(`${CONFIG.MESSAGES.SUCCESS}\n响应:${response}`);
                    
                    if (shouldAppendRemarks && this.Oid) {
                        const remarkText = this.generateAppendRemarkText(data);
                        if (remarkText) {
                            Utils.SaveAppendRemarks(2, remarkText, this.Oid, true)
                                .then(() => {
                                    this.hide();
                                });
                        } else {
                            this.hide();
                        }
                    } else {
                        this.hide();
                    }
                },
                (error) => {
                    Utils.showToast(`${CONFIG.MESSAGES.ERROR}${error}`);
                    sendBtn.disabled = false;
                    sendBtn.textContent = this.config.sendButtonText;
                    sendBtn.style.backgroundColor = '#3b82f6';
                }
            );

            setTimeout(() => {
                sendBtn.disabled = false;
                sendBtn.textContent = this.config.sendButtonText;
                sendBtn.style.backgroundColor = '#3b82f6';
            }, 500);
        }

        // 获取追加备注复选框 - 由子类实现
        getAppendRemarksCheckbox() {
            return null;
        }

        // 生成追加备注文本 - 由子类实现
        generateAppendRemarkText(data) {
            return null;
        }

        // 处理回车键事件，实现表单元素间的焦点切换
        handleEnterKey(e, focusableElements) {
            if (e.key !== 'Enter') return;

            e.preventDefault();
            const currentIndex = focusableElements.findIndex(el => el === document.activeElement);

            if (currentIndex === focusableElements.length - 1) {
                this.send();
            } else if (currentIndex !== -1 && focusableElements[currentIndex + 1]) {
                focusableElements[currentIndex + 1].focus();

                const nextElement = focusableElements[currentIndex + 1];
                if (nextElement.type === 'radio') {
                    nextElement.checked = true;
                    this.onRadioChange?.(nextElement);
                }
            }
        }

        // 生成商品款式单选按钮组
        generateProductStyleRadios(productStyleList, container, productStylePriceMap = {}) {
            container.innerHTML = '';

            if (productStyleList.length === 0) {
                container.innerHTML = '<div>未获取到款式信息</div>';
                return;
            }

            productStyleList.forEach((productStyle, index) => {
                const radioId = `product-style-${this.panelId}-${Date.now()}-${index}`;
                const price = productStylePriceMap[productStyle] || 0;
                const radioHtml = `
                    <div class="radio-item">
                        <input type="radio"
                               name="product-style-${this.panelId}"
                               id="${radioId}"
                               value="${productStyle}"
                               data-price="${price}"
                               class="panel-radio"
                               ${index === 0 ? 'checked' : ''}>
                        <label for="${radioId}">${productStyle}</label>
                    </div>
                `;
                container.insertAdjacentHTML('beforeend', radioHtml);
            });

            return container.querySelectorAll('input[type="radio"]');
        }

        setupShortcuts(shortcuts, inputElement, callback, idPrefix = null) {
            shortcuts.forEach((value, index) => {
                const shortcutId = idPrefix ? `${idPrefix}-${index}` : `${this.panelId}-shortcut-${index}`;
                const shortcut = document.getElementById(shortcutId);
                if (shortcut) {
                    shortcut.addEventListener('click', () => {
                        inputElement.value = value;
                        inputElement.focus();
                        callback?.();
                    });
                }
            });
        }

        // 通用的质检员自动完成功能
        setupInspectorAutocomplete(inspectorInput, inspectorDropdown, callback = null, inspectorList = CONFIG.INSPECTORS) {
            let selectedIndex = -1;

            inspectorInput.addEventListener('input', function() {
                const value = this.value.trim().toLowerCase();
                if (value.length === 0) {
                    inspectorDropdown.style.display = 'none';
                    selectedIndex = -1;
                    return;
                }

                const matches = inspectorList.filter(inspector =>
                    inspector.toLowerCase().includes(value)
                );

                if (matches.length === 0) {
                    inspectorDropdown.style.display = 'none';
                    selectedIndex = -1;
                    return;
                }

                inspectorDropdown.innerHTML = '';
                matches.forEach(inspector => {
                    const div = document.createElement('div');
                    div.className = 'inspector-item';
                    div.textContent = inspector;
                    div.addEventListener('click', () => {
                        inspectorInput.value = inspector;
                        inspectorDropdown.style.display = 'none';
                        inspectorInput.dispatchEvent(new Event('input'));
                        callback?.();
                    });
                    inspectorDropdown.appendChild(div);
                });

                selectedIndex = 0;
                const items = inspectorDropdown.querySelectorAll('.inspector-item');
                if (items.length > 0) {
                    items[0].classList.add('selected');
                }

                inspectorDropdown.style.display = 'block';
            });

            // 下拉菜单键盘导航
            inspectorInput.addEventListener('keydown', (e) => {
                const items = inspectorDropdown.querySelectorAll('.inspector-item');
                const isVisible = inspectorDropdown.style.display !== 'none';

                if (!isVisible || items.length === 0) return;

                switch(e.key) {
                    case 'ArrowDown':
                        e.preventDefault();
                        if (selectedIndex >= 0) items[selectedIndex].classList.remove('selected');
                        selectedIndex = (selectedIndex + 1) % items.length;
                        items[selectedIndex].classList.add('selected');
                        items[selectedIndex].scrollIntoView({block: 'nearest'});
                        break;
                    case 'ArrowUp':
                        e.preventDefault();
                        if (selectedIndex >= 0) items[selectedIndex].classList.remove('selected');
                        selectedIndex = (selectedIndex - 1 + items.length) % items.length;
                        items[selectedIndex].classList.add('selected');
                        items[selectedIndex].scrollIntoView({block: 'nearest'});
                        break;
                    case 'Enter':
                        e.preventDefault();
                        if (selectedIndex >= 0) {
                            inspectorInput.value = items[selectedIndex].textContent;
                            inspectorDropdown.style.display = 'none';
                            callback?.();
                        }
                        break;
                    case 'Escape':
                        inspectorDropdown.style.display = 'none';
                        selectedIndex = -1;
                        break;
                }
            });

            // 点击外部关闭下拉菜单
            document.addEventListener('click', (e) => {
                if (!inspectorInput.contains(e.target) && !inspectorDropdown.contains(e.target)) {
                    inspectorDropdown.style.display = 'none';
                    selectedIndex = -1;
                }
            });
        }
    }

    // ==================== 售后面板 ====================
    class AfterSalesPanel extends BasePanel {
        // 构造方法，初始化售后面板
        constructor() {
            super('custom-panel', {
                webhookUrl: CONFIG.WEBHOOKS.AFTERSALES,
                sendButtonText: '发送'
            });

            this.elements = {
                title: document.getElementById('panel-title'),
                shopLabel: document.getElementById('shop-bottom'),
                productStyleGroup: document.getElementById('product-style-group'),
                inspector: document.getElementById('inspector'),
                compensationAmount: document.getElementById('compensation-amount'),
                otherContentInput: document.getElementById('other-content-input'),
                otherContentRadio: document.getElementById('content-other')
            };

            this.setupEventListeners();
        }

        // 设置事件监听器
        setupEventListeners() {
            // Send button
            this.panel.querySelector('.btn-send').addEventListener('click', () => this.send());

            // Close button
            this.panel.querySelector('.panel-close').addEventListener('click', () => this.hide());

            // Compensation amount change
            this.elements.compensationAmount.addEventListener('input', () => this.handleCompensationChange());

            // Inspector change
            this.elements.inspector.addEventListener('input', () => this.calculateResponsibility());

            // Other content radio
            this.elements.otherContentRadio.addEventListener('change', () => this.updateOtherContentState());

            // All content radios
            document.querySelectorAll('input[name="content"]').forEach(radio => {
                radio.addEventListener('change', () => this.updateOtherContentState());
            });

            // Setup autocomplete for inspector
            this.setupInspectorAutocomplete();

            // Setup shortcuts
            this.setupShortcuts(
                CONFIG.AMOUNT_SHORTCUTS,
                this.elements.compensationAmount,
                () => this.handleCompensationChange(),
                'amount-shortcut'
            );

            // Setup inspector shortcuts
            this.setupShortcuts(
                CONFIG.INSPECTOR_SHORTCUTS,
                this.elements.inspector,
                () => this.calculateResponsibility(),
                'inspector-shortcut'
            );

            // Setup percentage shortcuts
            this.setupPercentageShortcuts();
        }

        setupPercentageShortcuts() {
            // 30% shortcut
            const thirtyPercentShortcut = document.getElementById('custom-panel-30-percent-shortcut');
            if (thirtyPercentShortcut) {
                thirtyPercentShortcut.addEventListener('click', () => {
                    this.applyPercentage(0.3);
                });
            }

            // 50% shortcut
            const fiftyPercentShortcut = document.getElementById('custom-panel-50-percent-shortcut');
            if (fiftyPercentShortcut) {
                fiftyPercentShortcut.addEventListener('click', () => {
                    this.applyPercentage(0.5);
                });
            }
        }

        applyPercentage(percentage) {
            // 获取选中的款式价格
            const selectedRadio = this.elements.productStyleGroup.querySelector('input[name^="product-style-"]:checked');
            if (!selectedRadio) {
                Utils.showToast(CONFIG.MESSAGES.SELECT_PRODUCT_STYLE);
                return;
            }

            const selectedPrice = parseFloat(selectedRadio.getAttribute('data-price')) || 0;
            const result = (selectedPrice * percentage).toFixed(2);
            this.elements.compensationAmount.value = result;
            this.elements.compensationAmount.focus();
            this.handleCompensationChange();
        }

        async populate() {
            this.reset();
            const data = await ERPDataExtractor.extractOrderData();
            if (!data) {
                this.hide();
                return null;
            }

            this.elements.title.textContent = data.orderId;
            this.elements.shopLabel.textContent = data.shopName;

            this.generateProductStyleRadios(
                data.productStyleList,
                this.elements.productStyleGroup,
                data.productStylePriceMap
            );

            this.calculateResponsibility();
            return data.Oid;
        }

        reset() {
            this.elements.title.textContent = "";
            this.elements.shopLabel.textContent = "";
            this.elements.inspector.value = '';
            this.elements.compensationAmount.value = '';
            this.elements.productStyleGroup.innerHTML = '';
            this.elements.otherContentInput.disabled = true;

            document.querySelectorAll('.panel-radio').forEach(radio => {
                radio.checked = false;
            });

            // 设置默认值
            const firstContent = document.querySelector('input[name="content"]');
            const firstRemark = document.querySelector('input[name="remark"]');
            if (firstContent) firstContent.checked = true;
            if (firstRemark) firstRemark.checked = true;
        }

        collectData() {
            // 获取选定的产品样式
            const productStyleRadios = document.querySelectorAll('#product-style-group input[name^="product-style-"]');
            let selectedProductStyle = '';
            productStyleRadios.forEach(radio => {
                if (radio.checked) selectedProductStyle = radio.value;
            });

            if (!selectedProductStyle) {
                Utils.showToast(CONFIG.MESSAGES.SELECT_PRODUCT_STYLE);
                return null;
            }

            // 获取选定的内容
            let selectedContent = '';
            document.querySelectorAll('input[name="content"]').forEach(radio => {
                if (radio.checked) {
                    selectedContent = radio.id === 'content-other'
                        ? this.elements.otherContentInput.value.trim()
                        : radio.value;
                }
            });

            if (!selectedContent) {
                Utils.showToast(CONFIG.MESSAGES.SELECT_OR_INPUT_CONTENT);
                return null;
            }

            // 获取选定的责任
            let selectedResponsibility = '';
            document.querySelectorAll('input[name="responsibility"]').forEach(radio => {
                if (radio.checked) selectedResponsibility = radio.value;
            });

            // 获取选定的备注
            let selectedRemark = '';
            document.querySelectorAll('input[name="remark"]').forEach(radio => {
                if (radio.checked) selectedRemark = radio.value;
            });

            return {
                os_id: this.elements.title.textContent.trim(),
                zerf: selectedResponsibility,
                buih: selectedRemark,
                krui: selectedProductStyle,
                shop: this.elements.shopLabel.textContent.trim(),
                q: parseFloat(this.elements.compensationAmount.value) || 0,
                text: selectedContent,
                check: this.elements.inspector.value.trim()
            };
        }

        // 计算责任方
        calculateResponsibility() {
            const checkValue = this.elements.inspector.value.trim();
            let responsibility = '';

            if (!checkValue) {
                responsibility = '公司';
            } else if (checkValue.startsWith('检验06')) {
                responsibility = '质检';
            } else if (checkValue.startsWith('合格')) {
                responsibility = '质检';
            } else if (checkValue.startsWith('检验')) {
                responsibility = '厂家';
            }

            document.querySelectorAll('input[name="responsibility"]').forEach(radio => {
                if (radio.value === responsibility) {
                    radio.checked = true;
                }
            });
        }

        handleCompensationChange() {
            const amount = this.elements.compensationAmount.value;
            const shouldSelectCompensation = amount && parseFloat(amount) > 0;

            document.querySelectorAll('input[name="remark"]').forEach(radio => {
                const label = document.querySelector(`label[for="${radio.id}"]`);
                const labelText = label?.textContent.trim();

                if (shouldSelectCompensation) {
                    radio.checked = labelText === "补偿";
                } else {
                    radio.checked = labelText === "空";
                }
            });
        }

        // 更新其他内容输入框状态
        updateOtherContentState() {
            const isOtherSelected = this.elements.otherContentRadio.checked;
            this.elements.otherContentInput.disabled = !isOtherSelected;
            if (isOtherSelected) {
                this.elements.otherContentInput.focus();
            }
        }

        setupInspectorAutocomplete() {
            const dropdown = document.getElementById('inspector-dropdown');
            super.setupInspectorAutocomplete(this.elements.inspector, dropdown, () => this.calculateResponsibility());
        }

        // 获取追加备注复选框
        getAppendRemarksCheckbox() {
            return document.getElementById('custom-panel-append-remarks');
        }

        // 生成追加备注文本 - 格式：[data.selectedProductStyle]text
        generateAppendRemarkText(data) {
            if (!data.krui || !data.text) {
                return null;
            }
            return `[${data.krui}]${data.text}`;
        }
    }

    // ==================== 挽单面板 ====================
    class WandanPanel extends BasePanel {
        // 构造方法，初始化挽单面板
        constructor() {
            super('wandan-panel', {
                webhookUrl: CONFIG.WEBHOOKS.WANDAN,
                sendButtonText: '发送到挽单表'
            });

            this.elements = {
                title: document.getElementById('wandan-panel-title'),
                shopLabel: document.getElementById('wandan-shop-bottom'),
                nameLabel: document.getElementById('wandan-name-bottom'),
                productStyleGroup: document.getElementById('wandan-product-style-group'),
                amount: document.getElementById('wandan-amount'),
                compensationAmount: document.getElementById('wandan-compensation-amount'),
                returnReason: document.getElementById('wandan-yryb'),
                result: document.getElementById('wandan-jpgo'),
                isKept: document.getElementById('wandan-done')
            };

            this.setupEventListeners();
        }

        setupEventListeners() {
            // 发送按钮点击事件
            this.panel.querySelector('.btn-send').addEventListener('click', () => this.send());

            // 关闭按钮点击事件
            this.panel.querySelector('.panel-close').addEventListener('click', () => this.hide());

            // 结果输入框变化事件
            this.elements.result.addEventListener('input', () => {
                this.elements.isKept.checked = (this.elements.result.value.trim() === '留下');
            });

            // 设置快捷方式
            this.setupShortcuts(
                CONFIG.AMOUNT_SHORTCUTS,
                this.elements.compensationAmount,
                null,
                'wandan-amount-shortcut'
            );

            this.setupShortcuts(
                CONFIG.WANDAN.RETURN_REASON_SHORTCUTS,
                this.elements.returnReason,
                null,
                'return-reason-shortcut'
            );

            this.setupShortcuts(
                CONFIG.WANDAN.RESULT_SHORTCUTS,
                this.elements.result,
                () => {
                    this.elements.isKept.checked = (this.elements.result.value === '留下');
                },
                'result-shortcut'
            );

            // 设置百分比快捷按钮
            this.setupPercentageShortcuts();
        }

        async populate() {
            this.reset();
            const data = await ERPDataExtractor.extractOrderData();
            if (!data) {
                this.hide();
                return null;
            }

            this.elements.title.textContent = data.orderId;
            this.elements.shopLabel.textContent = data.shopName;

            const newRadios = this.generateProductStyleRadios(
                data.productStyleList,
                this.elements.productStyleGroup,
                data.productStylePriceMap
            );

            // 设置默认金额基于第一个商品样式
            if (data.productStyleList.length > 0) {
                const firstPrice = data.productStylePriceMap[data.productStyleList[0]] || 0;
                this.elements.amount.value = firstPrice.toFixed(2);
            }

            // 添加事件监听器以更新商品样式变化时的金额
            if (newRadios) {
                newRadios.forEach(radio => {
                    radio.addEventListener('change', () => {
                        const selectedPrice = parseFloat(radio.getAttribute('data-price')) || 0;
                        this.elements.amount.value = selectedPrice.toFixed(2);
                    });
                });
            }

            // 设置默认类型
            const firstType = document.querySelector('input[name="wandan-type"]');
            if (firstType) firstType.checked = true;

            return data.Oid;
        }

        // 设置百分比快捷按钮
        setupPercentageShortcuts() {
            const thirtyPercentShortcut = document.getElementById('wandan-panel-30-percent-shortcut');
            if (thirtyPercentShortcut) {
                thirtyPercentShortcut.addEventListener('click', () => this.applyPercentage(0.3));
            }
            const fiftyPercentShortcut = document.getElementById('wandan-panel-50-percent-shortcut');
            if (fiftyPercentShortcut) {
                fiftyPercentShortcut.addEventListener('click', () => this.applyPercentage(0.5));
            }
        }

        applyPercentage(percentage) {
            const selectedRadio = this.elements.productStyleGroup.querySelector('input[name^="product-style-"]:checked');
            if (!selectedRadio) {
                Utils.showToast(CONFIG.MESSAGES.SELECT_PRODUCT_STYLE);
                return;
            }
            const selectedPrice = parseFloat(selectedRadio.getAttribute('data-price')) || 0;
            const result = (selectedPrice * percentage).toFixed(2);
            this.elements.compensationAmount.value = result;
            this.elements.compensationAmount.focus();
        }

        reset() {
            this.elements.title.textContent = "";
            this.elements.shopLabel.textContent = "";
            this.elements.compensationAmount.value = '';
            this.elements.returnReason.value = '不想要了';
            this.elements.result.value = '留下';
            this.elements.isKept.checked = true;
            this.elements.amount.value = '';
            this.elements.productStyleGroup.innerHTML = '';

            const firstType = document.querySelector('input[name="wandan-type"]');
            if (firstType) firstType.checked = true;
        }

        collectData() {
            // 获取选中的商品样式
            const productStyleRadios = document.querySelectorAll('#wandan-product-style-group input[name^="product-style-"]');
            let selectedProductStyle = '';
            productStyleRadios.forEach(radio => {
                if (radio.checked) selectedProductStyle = radio.value;
            });

            if (!selectedProductStyle) {
                Utils.showToast(CONFIG.MESSAGES.SELECT_PRODUCT_STYLE);
                return null;
            }

            // 获取选中的退款类型
            let selectedType = '';
            document.querySelectorAll('input[name="wandan-type"]').forEach(radio => {
                if (radio.checked) selectedType = radio.value;
            });

            if (!selectedType) {
                Utils.showToast(CONFIG.MESSAGES.SELECT_REFUND_TYPE);
                return null;
            }

            // 验证输入
            const returnReason = this.elements.returnReason.value.trim();
            if (!returnReason) {
                Utils.showToast(CONFIG.MESSAGES.INPUT_RETURN_REASON);
                return null;
            }

            const result = this.elements.result.value.trim();
            if (!result) {
                Utils.showToast(CONFIG.MESSAGES.INPUT_RESULT);
                return null;
            }

            return {
                os_id: this.elements.title.textContent.trim(),
                shop: this.elements.shopLabel.textContent.trim(),
                name: this.elements.nameLabel.textContent.trim(),
                type: selectedType,
                yryb: returnReason,
                jpgo: result,
                jxge: parseFloat(this.elements.amount.value) || 0,
                krui: selectedProductStyle,
                done: this.elements.isKept.checked,
                q: parseFloat(this.elements.compensationAmount.value) || 0
            };
        }

        // 获取追加备注复选框
        getAppendRemarksCheckbox() {
            return document.getElementById('wandan-append-remarks');
        }

        // 生成追加备注文本 - 格式：[data.name]挽单q元
        generateAppendRemarkText(data) {
            if (!data.name || data.q === undefined) {
                return null;
            }
            const amount = parseFloat(data.q) || 0;
            return `[${data.name}]挽单${amount}元`;
        }
    }

    // ==================== 错发面板 ====================
    class CuofaPanel extends BasePanel {
        // 构造方法，初始化错发面板
        constructor() {
            super('cuofa-panel', {
                webhookUrl: CONFIG.WEBHOOKS.CUOFA,
                sendButtonText: '发送到错发表'
            });

            this.elements = {
                title: document.getElementById('cuofa-panel-title'),
                shopLabel: document.getElementById('cuofa-shop-bottom'),
                productStyleGroup: document.getElementById('cuofa-product-style-group'),
                qkkl: document.getElementById('cuofa-qkkl'),
                qkklDropdown: document.getElementById('cuofa-qkkl-dropdown'),
                zerfOtherRadio: document.getElementById('cuofa-zerf-other'),
                zerfOtherInput: document.getElementById('cuofa-zerf-other-input'),
                djhc: document.getElementById('cuofa-djhc'),
                djhcDropdown: document.getElementById('cuofa-djhc-dropdown'),
                djhcTrackingNumber: document.getElementById('cuofa-djhc-tracking-number'),
                compensationAmount: document.getElementById('cuofa-compensation-amount')
            };

            this.setupEventListeners();
        }

        setupEventListeners() {
            // 发送按钮点击事件
            this.panel.querySelector('.btn-send').addEventListener('click', () => this.send());

            // 关闭按钮点击事件
            this.panel.querySelector('.panel-close').addEventListener('click', () => this.hide());

            // 责任人其他选项切换
            this.elements.zerfOtherRadio.addEventListener('change', () => {
                this.elements.zerfOtherInput.disabled = !this.elements.zerfOtherRadio.checked;
                if (this.elements.zerfOtherRadio.checked) {
                    this.elements.zerfOtherInput.focus();
                }
            });

            // 物流单号快捷按钮
            this.elements.djhcTrackingNumber.addEventListener('click', () => {
                if (this.currentTrackingNumber) {
                    this.elements.djhc.value = this.currentTrackingNumber;
                    this.elements.djhc.focus();
                }
            });

            // 监听错发情况的input事件
            this.elements.qkkl.addEventListener('input', (e) => {
                this.handleQkklChange(e.target.value);
            });

            // 错发情况自动完成
            this.setupQkklAutocomplete();

            // 质检员自动完成
            super.setupInspectorAutocomplete(this.elements.djhc, this.elements.djhcDropdown);

            // 设置快捷方式
            this.setupShortcuts(
                CONFIG.CUOFA.QKKL_SHORTCUTS,
                this.elements.qkkl,
                null,
                'cuofa-qkkl-shortcut'
            );

            this.setupShortcuts(
                CONFIG.CUOFA.DJHC_SHORTCUTS,
                this.elements.djhc,
                null,
                'cuofa-djhc-shortcut'
            );

            this.setupShortcuts(
                CONFIG.AMOUNT_SHORTCUTS,
                this.elements.compensationAmount,
                null,
                'cuofa-amount-shortcut'
            );

            // 为动态价格快捷按钮添加点击事件
            const setupShortcutEvent = (shortcutId) => {
                const shortcut = document.getElementById(shortcutId);
                if (shortcut) {
                    shortcut.addEventListener('click', () => {
                        const value = shortcut.getAttribute('data-value');
                        if (value) {
                            this.elements.compensationAmount.value = value;
                            this.elements.compensationAmount.focus();
                        }
                    });
                }
            };

            setupShortcutEvent('cuofa-panel-total-amount-shortcut');
            setupShortcutEvent('cuofa-panel-style-amount-shortcut');

            // 添加回车键支持
            const focusableElements = [
                this.elements.qkkl,
                this.elements.djhc,
                this.elements.compensationAmount
            ];

            focusableElements.forEach(element => {
                element.addEventListener('keydown', (e) => {
                    this.handleEnterKey(e, focusableElements);
                });
            });
        }

        async populate() {
            this.reset();
            const data = await ERPDataExtractor.extractOrderData();
            if (!data) {
                this.hide();
                return null;
            }

            this.elements.title.textContent = data.orderId;
            this.elements.shopLabel.textContent = data.shopName;
            this.currentTrackingNumber = data.trackingNumber;

            // 更新物流单号快捷按钮文本
            this.elements.djhcTrackingNumber.textContent = data.trackingNumber || '无物流单号';

            // 生成款式单选框
            const newRadios = this.generateProductStyleRadios(
                data.productStyleList,
                this.elements.productStyleGroup,
                data.productStylePriceMap
            );

            // 监听款式变化，更新补偿金额第一个快捷按钮
            if (newRadios) {
                newRadios.forEach(radio => {
                    radio.addEventListener('change', () => {
                        this.updateCompensationAmountShortcut();
                    });
                });

                // 初始调用，设置默认款式的价格
                this.updateCompensationAmountShortcut();
            }

            // 设置默认值
            const firstResponsibility = document.querySelector('input[name="cuofa-zerf"]');
            if (firstResponsibility) firstResponsibility.checked = true;
            return data.Oid;
        }

        reset() {
            this.elements.title.textContent = "";
            this.elements.shopLabel.textContent = "";
            this.elements.qkkl.value = '';
            this.elements.djhc.value = '';
            this.elements.compensationAmount.value = '';
            this.elements.productStyleGroup.innerHTML = '';
            this.elements.zerfOtherInput.disabled = true;
            this.elements.zerfOtherRadio.checked = false;
            this.currentTrackingNumber = '';
            this.elements.djhcTrackingNumber.textContent = '物流单号';

            // 设置默认值
            const firstResponsibility = document.querySelector('input[name="cuofa-zerf"]');
            if (firstResponsibility) firstResponsibility.checked = true;
        }

        collectData() {
            // 获取选中的商品样式
            const productStyleRadios = document.querySelectorAll('#cuofa-product-style-group input[name^="product-style-"]');
            let selectedProductStyle = '';
            productStyleRadios.forEach(radio => {
                if (radio.checked) selectedProductStyle = radio.value;
            });

            if (!selectedProductStyle) {
                Utils.showToast(CONFIG.MESSAGES.SELECT_PRODUCT_STYLE);
                return null;
            }

            // 获取错发情况
            const qkkl = this.elements.qkkl.value.trim();
            if (!qkkl) {
                Utils.showToast('请输入错发情况');
                return null;
            }

            // 获取责任人
            let selectedResponsibility = '';
            document.querySelectorAll('input[name="cuofa-zerf"]').forEach(radio => {
                if (radio.checked) {
                    selectedResponsibility = radio.id === 'cuofa-zerf-other'
                        ? this.elements.zerfOtherInput.value.trim()
                        : radio.value;
                }
            });

            if (!selectedResponsibility) {
                Utils.showToast('请选择责任人');
                return null;
            }

            // 获取单号
            const djhc = this.elements.djhc.value.trim();

            return {
                shop: this.elements.shopLabel.textContent.trim(),
                os_id: this.elements.title.textContent.trim(),
                krui: selectedProductStyle,
                qkkl: qkkl,
                zerf: selectedResponsibility,
                djhc: djhc,
                q: parseFloat(this.elements.compensationAmount.value) || 0
            };
        }

        // 处理错发情况变化的规则
        handleQkklChange(value) {
            value = value.trim();

            // 规则1：错发情况为“无走件”时自动填写包裹运单号
            if (value === '无走件') {
                if (this.currentTrackingNumber) {
                    this.elements.djhc.value = this.currentTrackingNumber;
                    this.elements.djhc.focus();
                }
            }

            // 规则3：错发情况为特定值时自动选择编辑框并聚焦
            if (value === '编码错误' || value === '客服换错货') {
                // 选择“其他”责任并聚焦输入框
                this.elements.zerfOtherRadio.checked = true;
                this.elements.zerfOtherInput.disabled = false;
                this.elements.zerfOtherInput.focus();
            }
        }

        setupQkklAutocomplete() {
            let selectedIndex = -1;

            this.elements.qkkl.addEventListener('input', (e) => {
                const value = e.target.value.trim().toLowerCase();
                // 应用错发情况变化规则
                this.handleQkklChange(e.target.value);

                if (value.length === 0) {
                    this.elements.qkklDropdown.style.display = 'none';
                    selectedIndex = -1;
                    return;
                }

                const matches = CONFIG.CUOFA.QKKL_OPTIONS.filter(option =>
                    option.toLowerCase().includes(value)
                );

                if (matches.length === 0) {
                    this.elements.qkklDropdown.style.display = 'none';
                    selectedIndex = -1;
                    return;
                }

                this.elements.qkklDropdown.innerHTML = '';
                matches.forEach(option => {
                    const div = document.createElement('div');
                    div.className = 'inspector-item';
                    div.textContent = option;
                    div.addEventListener('click', () => {
                        this.elements.qkkl.value = option;
                        this.elements.qkklDropdown.style.display = 'none';
                        this.elements.qkkl.focus();
                        // 应用错发情况变化规则
                        this.handleQkklChange(option);
                    });
                    this.elements.qkklDropdown.appendChild(div);
                });

                selectedIndex = 0;
                const items = this.elements.qkklDropdown.querySelectorAll('.inspector-item');
                if (items.length > 0) {
                    items[0].classList.add('selected');
                }

                this.elements.qkklDropdown.style.display = 'block';
            });

            // 下拉菜单键盘导航
            this.elements.qkkl.addEventListener('keydown', (e) => {
                const items = this.elements.qkklDropdown.querySelectorAll('.inspector-item');
                const isVisible = this.elements.qkklDropdown.style.display !== 'none';

                if (!isVisible || items.length === 0) return;

                switch(e.key) {
                    case 'ArrowDown':
                        e.preventDefault();
                        if (selectedIndex >= 0) items[selectedIndex].classList.remove('selected');
                        selectedIndex = (selectedIndex + 1) % items.length;
                        items[selectedIndex].classList.add('selected');
                        items[selectedIndex].scrollIntoView({block: 'nearest'});
                        break;
                    case 'ArrowUp':
                        e.preventDefault();
                        if (selectedIndex >= 0) items[selectedIndex].classList.remove('selected');
                        selectedIndex = (selectedIndex - 1 + items.length) % items.length;
                        items[selectedIndex].classList.add('selected');
                        items[selectedIndex].scrollIntoView({block: 'nearest'});
                        break;
                    case 'Enter':
                        e.preventDefault();
                        if (selectedIndex >= 0) {
                            const selectedValue = items[selectedIndex].textContent;
                            this.elements.qkkl.value = selectedValue;
                            this.elements.qkklDropdown.style.display = 'none';
                            // 应用错发情况变化规则
                            this.handleQkklChange(selectedValue);
                        }
                        break;
                    case 'Escape':
                        this.elements.qkklDropdown.style.display = 'none';
                        selectedIndex = -1;
                        break;
                }
            });

            // 点击外部关闭下拉菜单
            document.addEventListener('click', (e) => {
                if (!this.elements.qkkl.contains(e.target) && !this.elements.qkklDropdown.contains(e.target)) {
                    this.elements.qkklDropdown.style.display = 'none';
                    selectedIndex = -1;
                }
            });
        }

        // 更新补偿金额快捷按钮
        async updateCompensationAmountShortcut() {
            // 获取当前订单数据
            const data = await ERPDataExtractor.extractOrderData();
            if (!data) return;

            // 计算总金额价值（所有商品价格乘以数量之和）
            const totalAmount = data.productList.reduce((sum, product) => {
                return sum + (product.price * product.quantity);
            }, 0);

            // 获取唯一商品编码数量
            const uniqueProductCodes = new Set(data.productList.map(product => product.productCode)).size;

            // 获取选中的款式价格（无需乘以数量）
            const selectedRadio = this.elements.productStyleGroup.querySelector('input[name^="product-style-"]:checked');
            const selectedPrice = selectedRadio ? parseFloat(selectedRadio.getAttribute('data-price')) || 0 : 0;

            const totalShortcut = document.getElementById('cuofa-panel-total-amount-shortcut');
            const styleShortcut = document.getElementById('cuofa-panel-style-amount-shortcut');

            // 设置总金额价值按钮
            if (totalShortcut) {
                totalShortcut.style.display = 'inline';
                totalShortcut.textContent = totalAmount.toFixed(2);
                totalShortcut.setAttribute('data-value', totalAmount.toFixed(2));
            }

            // 根据唯一商品编码数量显示/隐藏选中款式的价值按钮
            if (styleShortcut) {
                if (uniqueProductCodes > 1) {
                    styleShortcut.style.display = 'inline';
                    styleShortcut.textContent = selectedPrice.toFixed(2);
                    styleShortcut.setAttribute('data-value', selectedPrice.toFixed(2));
                } else {
                    styleShortcut.style.display = 'none';
                }
            }
        }

        // 获取追加备注复选框
        getAppendRemarksCheckbox() {
            return document.getElementById('cuofa-panel-remarks');
        }

        // 生成追加备注文本 - 格式：此单错发了[data.krui]，原因：[qkkl]，责任人：[zerf]
        generateAppendRemarkText(data) {
            if (!data.krui || !data.qkkl || !data.zerf) {
                return null;
            }
            return `此单错发了[${data.krui}]，原因：[${data.qkkl}]，责任人：[${data.zerf}]`;
        }
    }

    // ==================== 私域面板 ====================
    class SiyuPanel extends BasePanel {
        constructor() {
            super('siyu-panel', {
                webhookUrl: CONFIG.WEBHOOKS.SIYU,
                sendButtonText: '发送到私域表格'
            });

            this.elements = {
                title: document.getElementById('siyu-panel-title'),
                shopLabel: document.getElementById('siyu-shop-bottom'),
                productStyleGroup: document.getElementById('siyu-product-style-group'),
                q: document.getElementById('siyu-q'),
                uull: document.getElementById('siyu-uull')
            };

            this.setupEventListeners();
        }

        setupEventListeners() {

            // 发送按钮点击事件
            this.panel.querySelector('.btn-send').addEventListener('click', () => this.send());

            // 关闭按钮点击事件
            this.panel.querySelector('.panel-close').addEventListener('click', () => this.hide());

            // title点击事件

            this.panel.querySelector('#siyu-panel-title').addEventListener('click',() => {
                if(this.isOid){
                    this.isOid = false;
                    this.elements.title.textContent = `${this.currentOrderId || '未找到'}`;
                } else {
                    this.isOid = true;
                    this.elements.title.textContent = `${this.OrderOid || '未找到'}`;
                };
            });

            // 设置快捷方式
            this.setupShortcuts(
                CONFIG.SIYU.Q_SHORTCUTS,
                this.elements.q,
                null,
                'siyu-q-shortcut'
            );

            // 添加回车键支持
            const focusableElements = [
                ...Array.from(this.elements.productStyleGroup.querySelectorAll('input[type="radio"]')),
                this.elements.q,
                ...Array.from(this.panel.querySelectorAll('input[name="siyu-vekz"]')),
                ...Array.from(this.panel.querySelectorAll('input[name="siyu-type"]'))
            ];

            focusableElements.forEach(element => {
                element.addEventListener('keydown', (e) => this.handleEnterKey(e, focusableElements));
            });
        }

        async populate() {
            this.reset();
            const data = await ERPDataExtractor.extractOrderData();
            if (!data) {
                this.hide();
                return null;
            }

            this.currentOrderId = data.orderId;
            this.OrderOid = data.Oid;
            this.currentShopName = data.shopName;

            this.elements.title.textContent = `${this.OrderOid || '未找到'}`;

            this.generateProductStyleRadios(
                data.productStyleList,
                this.elements.productStyleGroup,
                data.productStylePriceMap
            );

            const totalQuantity = data.productList.reduce((sum, product) => sum + (product.quantity || 0), 0);
            this.elements.uull.value = totalQuantity;
            return data.Oid;
        }

        reset() {
            // 重置输入字段
            this.elements.q.value = '';
            this.elements.uull.value = '';

            // 重置单选按钮
            this.panel.querySelectorAll('input[type="radio"]').forEach(radio => {
                radio.checked = false;
            });

            // 默认选中第一个选项
            const firstVekzRadio = this.panel.querySelector('input[name="siyu-vekz"]');
            const firstTypeRadio = this.panel.querySelector('input[name="siyu-type"]');
            if (firstVekzRadio) firstVekzRadio.checked = true;
            if (firstTypeRadio) firstTypeRadio.checked = true;

            // 清空商品款式选项
            this.elements.productStyleGroup.innerHTML = '';

            this.isOid = false;
        }

        collectData() {
            const selectedRadio = this.elements.productStyleGroup.querySelector('input[name^="product-style-"]:checked');
            if (!selectedRadio) {
                Utils.showToast(CONFIG.MESSAGES.SELECT_PRODUCT_STYLE);
                return null;
            }

            const productStyle = selectedRadio.value;
            const vekz = this.panel.querySelector('input[name="siyu-vekz"]:checked')?.value || '';
            const q = parseFloat(this.elements.q.value) || 0;
            const uull = parseInt(this.elements.uull.value) || 0;
            const type = this.panel.querySelector('input[name="siyu-type"]:checked')?.value || '';
            let os_id = '';
            if (this.isOid){
                os_id = this.OrderOid;
            }else{
                os_id = this.currentOrderId;
            };
            if (type === '空') type == '';

            if (!this.currentOrderId) {
                Utils.showToast(CONFIG.MESSAGES.NO_ORDER_ID);
                return null;
            }

            return {

                os_id: os_id,
                krui: productStyle,
                vekz: vekz,
                q: q,
                uull: uull,
                type: type
            };
        }
    }

    // ==================== 快递丢件面板 ====================
    class DiuJianPanel extends BasePanel {
        constructor() {
            super('diujian-panel', {
                webhookUrl: CONFIG.WEBHOOKS.DIUJIAN,
                sendButtonText: '发送到快递丢件表'
            });

            this.elements = {
                title: document.getElementById('diujian-panel-title'),
                shopLabel: document.getElementById('diujian-shop-bottom'),
                trackingNumber: document.getElementById('diujian-tracking-number'),
                logisticsCompany: document.getElementById('diujian-logistics-company'),
                jxvi: document.getElementById('diujian-jxvi')
            };

            this.setupEventListeners();
        }

        setupEventListeners() {
            this.panel.querySelector('.btn-send').addEventListener('click', () => this.send());
            this.panel.querySelector('.panel-close').addEventListener('click', () => this.hide());
	//监听鼠标点击标签事件
            this.elements.logisticsCompany.addEventListener('click', () => {
                const currentText = this.elements.logisticsCompany.textContent.trim();
                if (currentText === '申通E物流') {
                    this.elements.logisticsCompany.textContent = '广州申通';
                    this.currentLogisticsCompany = '广州申通';
                } else if (currentText === '广州申通') {
                    this.elements.logisticsCompany.textContent = '申通E物流';
                    this.currentLogisticsCompany = '申通E物流';
                }
            });

            const focusableElements = Array.from(this.panel.querySelectorAll('input[name="diujian-qkkl"]'));

            focusableElements.forEach(element => {
                element.addEventListener('keydown', (e) => this.handleEnterKey(e, focusableElements));
            });
        }

        async populate() {
            this.reset();
            const data = await ERPDataExtractor.extractOrderData();
            if (!data) {
                this.hide();
                return null;
            }

            this.currentOrderId = data.orderId;
            this.currentShopName = data.shopName;
            this.currentTrackingNumber = data.trackingNumber;
            this.currentLogisticsCompany = data.logisticsCompany;

            if (!this.currentTrackingNumber || this.currentTrackingNumber.startsWith('@')) {
                Utils.showToast('没有查询到包裹运单号');
                this.hide();
                return;
            }

            this.elements.title.textContent = `${this.currentOrderId || '未找到'}`;
            this.elements.trackingNumber.textContent = this.currentTrackingNumber || '未获取';
            this.elements.logisticsCompany.textContent = this.currentLogisticsCompany || '未获取';

            const totalAmount = data.productList.reduce((sum, product) => {
                return sum + (product.price * product.quantity);
            }, 0);
            this.elements.jxvi.value = totalAmount.toFixed(2);
            return data.Oid;
        }

        reset() {
            this.elements.jxvi.value = '';
            this.elements.trackingNumber.textContent = '';
            this.elements.logisticsCompany.textContent = '';

            this.panel.querySelectorAll('input[type="radio"]').forEach(radio => {
                radio.checked = false;
            });

            const firstQkklRadio = this.panel.querySelector('input[name="diujian-qkkl"]');
            if (firstQkklRadio) firstQkklRadio.checked = true;
        }

        collectData() {
            const qkkl = this.panel.querySelector('input[name="diujian-qkkl"]:checked')?.value || '';
            const jxvi = parseFloat(this.elements.jxvi.value) || 0;

            if (!this.currentOrderId) {
                Utils.showToast(CONFIG.MESSAGES.NO_ORDER_ID);
                return null;
            }

            return {
                os_id: this.currentOrderId,
                dh_id: this.currentTrackingNumber || '',
                gssi: this.currentLogisticsCompany || '',
                qkkl: qkkl,
                jxvi: jxvi
            };
        }

        // 获取追加备注复选框
        getAppendRemarksCheckbox() {
            return document.getElementById('diujian-append-remarks');
        }

        // 生成追加备注文本 - 格式：快递丢件
        generateAppendRemarkText(data) {
            return '快递丢件';
        }
    }

    // ==================== 面板管理器 ====================
    class PanelManager {
        // 构造方法，初始化面板管理器
        constructor() {
            this.panels = {
                afterSales: new AfterSalesPanel(),
                wandan: new WandanPanel(),
                cuofa: new CuofaPanel(),
                siyu: new SiyuPanel(),
                diujian: new DiuJianPanel()
            };

            this.setupGlobalEventListeners();
            this.injectStyles();
            this.makePanelsDraggable();
        }

        // 设置全局事件监听器
        setupGlobalEventListeners() {
            let lastKey = null;
            let lastKeyTime = 0;
            const DOUBLE_CLICK_DELAY = 500; // 双击间隔时间(毫秒)

            document.addEventListener('keydown', (e) => {
                // 处理双击快捷键
                const currentTime = Date.now();
                const isDoubleClick = (
                    e.key === lastKey &&
                    currentTime - lastKeyTime < DOUBLE_CLICK_DELAY
                );

                // 记录当前按键信息
                if (e.key === lastKey && !isDoubleClick) {
                    lastKeyTime = currentTime;
                } else {
                    lastKey = e.key;
                    lastKeyTime = currentTime;
                }

                // 切换售后面板
                if (isDoubleClick && e.key === CONFIG.HOTKEYS.TOGGLE_AFTERSALES) {
                    e.preventDefault();
                    this.togglePanel('afterSales');
                    // 重置状态避免连续触发
                    lastKey = null;
                }

                // 切换挽单面板
                if (isDoubleClick && e.key === CONFIG.HOTKEYS.TOGGLE_WANDAN) {
                    e.preventDefault();
                    this.togglePanel('wandan');
                    lastKey = null;
                }

                // 切换错发面板
                if (isDoubleClick && e.key === CONFIG.HOTKEYS.TOGGLE_CUOFA) {
                    e.preventDefault();
                    this.togglePanel('cuofa');
                    lastKey = null;
                }

                // 切换私域面板
                if (isDoubleClick && e.key === CONFIG.HOTKEYS.TOGGLE_SIYU) {
                    e.preventDefault();
                    this.togglePanel('siyu');
                    lastKey = null;
                }

                // 切换丢件面板
                if (isDoubleClick && e.key === CONFIG.HOTKEYS.TOGGLE_DIUJIAN) {
                    e.preventDefault();
                    this.togglePanel('diujian');
                    lastKey = null;
                }

                // Escape: 关闭所有面板
                if (e.key === CONFIG.HOTKEYS.CLOSE_ALL) {
                    e.preventDefault();
                    this.closeAllPanels();
                }
            });
        }

        // 切换面板显示状态
        togglePanel(panelName) {
            const targetPanel = this.panels[panelName];
            if (!targetPanel) return;

            // 关闭所有其他面板
            Object.entries(this.panels).forEach(([name, panel]) => {
                if (name !== panelName && panel.panel.classList.contains('show')) {
                    panel.hide();
                }
            });

            // 切换目标面板
            targetPanel.toggle();
        }

        // 关闭所有面板
        closeAllPanels() {
            Object.values(this.panels).forEach(panel => panel.hide());
        }

        injectStyles() {
            const styles = `
                #custom-panel, #wandan-panel, #cuofa-panel, #siyu-panel, #diujian-panel {
                    position: fixed;
                    top: 50%;
                    left: 50%;
                    transform: translate(-50%, -50%);
                    width: 400px;
                    padding: 15px;
                    background: rgba(255, 255, 255, 0.95);
                    border: 1px solid #e5e7eb;
                    border-radius: 6px;
                    box-shadow: 0 2px 16px rgba(0,0,0,0.15);
                    z-index: 999999;
                    display: none;
                    font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, sans-serif;
                    font-size: 13px;
                    cursor: default;
                    user-select: none;
                }

                #custom-panel.show, #wandan-panel.show, #cuofa-panel.show, #siyu-panel.show, #diujian-panel.show {
                    display: block;
                }

                #custom-panel.dragging, #wandan-panel.dragging, #cuofa-panel.dragging, #siyu-panel.dragging, #diujian-panel.dragging {
                    opacity: 0.8;
                }

                #custom-panel .panel-header, #wandan-panel .panel-header, #cuofa-panel .panel-header, #siyu-panel .panel-header, #diujian-panel .panel-header {
                    display: flex;
                    justify-content: center;
                    align-items: center;
                    margin-bottom: 10px;
                    padding-bottom: 5px;
                    border-bottom: 1px solid #f3f4f6;
                    position: relative;
                    cursor: default;
                }

                #custom-panel .panel-title, #wandan-panel .panel-title, #cuofa-panel .panel-title, #siyu-panel .panel-title, #diujian-panel .panel-title {
                    font-size: 16px;
                    font-weight: 600;
                    color: #111827;
                }

                #custom-panel .panel-close, #wandan-panel .panel-close, #cuofa-panel .panel-close, #siyu-panel .panel-close, #diujian-panel .panel-close {
                    position: absolute;
                    right: 0;
                    cursor: pointer;
                    color: #6b7280;
                    font-size: 16px;
                    line-height: 1;
                }

                #custom-panel .panel-type, #wandan-panel .panel-type, #cuofa-panel .panel-type, #siyu-panel .panel-type, #diujian-panel .panel-type {
                    position: absolute;
                    right: 25px;
                    background-color: #f59e0b;
                    color: white;
                    padding: 2px 6px;
                    border-radius: 4px;
                    font-size: 12px;
                    font-weight: bold;
                }

                #custom-panel .panel-body, #wandan-panel .panel-body, #cuofa-panel .panel-body, #siyu-panel .panel-body, #diujian-panel .panel-body {
                    margin-bottom: 15px;
                }

                #custom-panel .form-item, #wandan-panel .form-item, #cuofa-panel .form-item, #siyu-panel .form-item, #diujian-panel .form-item {
                    margin-bottom: 8px;
                }

                #custom-panel .form-section, #wandan-panel .form-section, #cuofa-panel .form-section, #siyu-panel .form-section, #diujian-panel .form-section {
                    margin-bottom: 12px;
                    padding-bottom: 10px;
                    border-bottom: 1px dashed #f3f4f6;
                }

                #custom-panel .form-section:last-child, #wandan-panel .form-section:last-child, #cuofa-panel .form-section:last-child, #siyu-panel .form-section:last-child, #diujian-panel .form-section:last-child {
                    border-bottom: none;
                }

                #custom-panel label, #wandan-panel label, #cuofa-panel label, #siyu-panel label, #diujian-panel label {
                    display: block;
                    margin-bottom: 3px;
                    color: #374151;
                    font-size: 12px;
                }

                #custom-panel input[type="text"],
                #custom-panel input[type="number"],
                #wandan-panel input[type="text"],
                #wandan-panel input[type="number"],
                #cuofa-panel input[type="text"],
                #cuofa-panel input[type="number"],
                #siyu-panel input[type="text"],
                #siyu-panel input[type="number"],
                #diujian-panel input[type="text"],
                #diujian-panel input[type="number"] {
                    width: 100%;
                    padding: 6px 8px;
                    border: 1px solid #d1d5db;
                    border-radius: 3px;
                    font-size: 12px;
                    background: rgba(255, 255, 255, 0.8);
                    box-sizing: border-box;
                    outline: none;
                }

                #custom-panel input[type="text"]:focus,
                #custom-panel input[type="number"]:focus,
                #wandan-panel input[type="text"]:focus,
                #wandan-panel input[type="number"]:focus,
                #cuofa-panel input[type="text"]:focus,
                #cuofa-panel input[type="number"]:focus,
                #siyu-panel input[type="text"]:focus,
                #siyu-panel input[type="number"]:focus,
                #diujian-panel input[type="text"]:focus,
                #diujian-panel input[type="number"]:focus {
                    border-color: #93c5fd;
                    box-shadow: 0 0 0 1px #93c5fd;
                }

                #custom-panel input[type="checkbox"], #wandan-panel input[type="checkbox"], #cuofa-panel input[type="checkbox"], #siyu-panel input[type="checkbox"], #diujian-panel input[type="checkbox"] {
                    width: auto;
                }

                #custom-panel .radio-group, #wandan-panel .radio-group, #cuofa-panel .radio-group, #siyu-panel .radio-group, #diujian-panel .radio-group {
                    display: flex;
                    flex-wrap: wrap;
                    gap: 10px;
                    margin-top: 3px;
                }

                #custom-panel .radio-item, #wandan-panel .radio-item, #cuofa-panel .radio-item, #siyu-panel .radio-item, #diujian-panel .radio-item {
                    display: flex;
                    align-items: center;
                    gap: 4px;
                }

                #custom-panel .radio-item.btn-style, #wandan-panel .radio-item.btn-style, #cuofa-panel .radio-item.btn-style, #siyu-panel .radio-item.btn-style, #diujian-panel .radio-item.btn-style {
                    position: relative;
                }

                #custom-panel .radio-item.btn-style input[type="radio"], #wandan-panel .radio-item.btn-style input[type="radio"], #cuofa-panel .radio-item.btn-style input[type="radio"], #siyu-panel .radio-item.btn-style input[type="radio"], #diujian-panel .radio-item.btn-style input[type="radio"] {
                    display: none;
                }

                #custom-panel .radio-item.btn-style label, #wandan-panel .radio-item.btn-style label, #cuofa-panel .radio-item.btn-style label, #siyu-panel .radio-item.btn-style label, #diujian-panel .radio-item.btn-style label {
                    display: inline-block;
                    padding: 4px 10px;
                    border-radius: 4px;
                    background-color: #e7e7e7;
                    color: #929292;
                    cursor: pointer;
                    transition: all 0.2s ease;
                    margin-bottom: 0;
                }

                #custom-panel .radio-item.btn-style input[type="radio"]:checked + label, #wandan-panel .radio-item.btn-style input[type="radio"]:checked + label, #cuofa-panel .radio-item.btn-style input[type="radio"]:checked + label, #siyu-panel .radio-item.btn-style input[type="radio"]:checked + label, #diujian-panel .radio-item.btn-style input[type="radio"]:checked + label {
                    background-color: #bfdbfe;
                    color: #6786eb;
                }

                #custom-panel input[type="radio"], #wandan-panel input[type="radio"], #cuofa-panel input[type="radio"], #siyu-panel input[type="radio"], #diujian-panel input[type="radio"] {
                    width: 12px;
                    height: 12px;
                }

                #custom-panel .btn-group, #wandan-panel .btn-group, #cuofa-panel .btn-group, #siyu-panel .btn-group, #diujian-panel .btn-group {
                    display: flex;
                    justify-content: flex-end;
                    margin-bottom: 10px;
                }

                #custom-panel button, #wandan-panel button, #cuofa-panel button, #siyu-panel button, #diujian-panel button {
                    padding: 6px 15px;
                    border: none;
                    border-radius: 3px;
                    cursor: pointer;
                    font-size: 12px;
                    transition: background 0.2s ease;
                }

                #custom-panel .btn-send, #wandan-panel .btn-send, #cuofa-panel .btn-send, #siyu-panel .btn-send, #diujian-panel .btn-send {
                    background: #3b82f6;
                    color: #fff;
                }

                #custom-panel .btn-send:hover, #wandan-panel .btn-send:hover, #cuofa-panel .btn-send:hover, #siyu-panel .btn-send:hover, #diujian-panel .btn-send:hover {
                    background: #2563eb;
                }

                #custom-panel .btn-send:focus, #wandan-panel .btn-send:focus, #cuofa-panel .btn-send:focus, #siyu-panel .btn-send:focus, #diujian-panel .btn-send:focus {
                    outline: 2px solid #93c5fd;
                    outline-offset: 1px;
                }

                #custom-panel .shortcut-group, #wandan-panel .shortcut-group, #cuofa-panel .shortcut-group, #siyu-panel .shortcut-group, #diujian-panel .shortcut-group {
                    display: flex;
                    flex-wrap: wrap;
                    gap: 8px;
                    margin-top: 3px;
                }

                #custom-panel .shortcut, #wandan-panel .shortcut, #cuofa-panel .shortcut, #siyu-panel .shortcut, #diujian-panel .shortcut {
                    color: #3b82f6;
                    cursor: pointer;
                    text-decoration: underline;
                    font-size: 12px;
                }

                #custom-panel .shortcut:hover, #wandan-panel .shortcut:hover, #cuofa-panel .shortcut:hover, #siyu-panel .shortcut:hover, #diujian-panel .shortcut:hover {
                    color: #2563eb;
                }

                #custom-panel .other-content-group, #wandan-panel .other-content-group, #cuofa-panel .other-content-group, #siyu-panel .other-content-group, #diujian-panel .other-content-group {
                    display: flex;
                    align-items: center;
                    gap: 8px;
                    margin-top: 5px;
                }

                #other-content-input {
                    flex: 1;
                }

                #custom-panel .bottom-labels, #wandan-panel .bottom-labels, #cuofa-panel .bottom-labels, #siyu-panel .bottom-labels, #diujian-panel .bottom-labels {
                    display: flex;
                    justify-content: space-between;
                    margin-top: 5px;
                    padding-top: 5px;
                    border-top: 1px solid #f3f4f6;
                }

                #custom-panel .bottom-label, #wandan-panel .bottom-label, #cuofa-panel .bottom-label, #siyu-panel .bottom-label, #diujian-panel .bottom-label {
                    display: inline-block;
                    padding: 4px 8px;
                    background-color: #f3f4f6;
                    border-radius: 3px;
                    font-size: 12px;
                    color: #6b7280;
                }

                #custom-panel .inspector-dropdown, #wandan-panel .inspector-dropdown, #cuofa-panel .inspector-dropdown, #siyu-panel .inspector-dropdown, #diujian-panel .inspector-dropdown {
                    position: absolute;
                    width: calc(100% - 16px);
                    max-height: 150px;
                    overflow-y: auto;
                    background: white;
                    border: 1px solid #d1d5db;
                    border-radius: 3px;
                    box-shadow: 0 2px 8px rgba(0,0,0,0.1);
                    z-index: 1000;
                    margin-top: 2px;
                }

                #custom-panel .inspector-item, #wandan-panel .inspector-item, #cuofa-panel .inspector-item, #siyu-panel .inspector-item, #diujian-panel .inspector-item {
                    padding: 5px 8px;
                    cursor: pointer;
                    font-size: 12px;
                }

                #custom-panel .inspector-item:hover,
                #custom-panel .inspector-item.selected,
                #wandan-panel .inspector-item:hover,
                #wandan-panel .inspector-item.selected,
                #cuofa-panel .inspector-item:hover,
                #cuofa-panel .inspector-item.selected,
                #siyu-panel .inspector-item:hover,
                #siyu-panel .inspector-item.selected,
                #diujian-panel .inspector-item:hover,
                #diujian-panel .inspector-item.selected {
                    background-color: #f3f4f6;
                }

                #custom-panel .inspector-container, #wandan-panel .inspector-container, #cuofa-panel .inspector-container, #siyu-panel .inspector-container, #diujian-panel .inspector-container {
                    position: relative;
                }

                #custom-panel .checkbox-item, #wandan-panel .checkbox-item, #cuofa-panel .checkbox-item, #siyu-panel .checkbox-item, #diujian-panel .checkbox-item {
                    display: flex;
                    align-items: center;
                    gap: 8px;
                    margin-top: 5px;
                }

                #custom-panel .flex-row, #wandan-panel .flex-row, #cuofa-panel .flex-row, #siyu-panel .flex-row, #diujian-panel .flex-row {
                    display: flex;
                    gap: 15px;
                    align-items: flex-start;
                }

                #custom-panel .flex-item, #wandan-panel .flex-item, #cuofa-panel .flex-item, #siyu-panel .flex-item, #diujian-panel .flex-item {
                    flex: 1;
                    min-width: 0;
                }
            `;

            const styleElement = document.createElement('style');
            styleElement.textContent = styles;
            document.head.appendChild(styleElement);
        }

        makePanelsDraggable() {
            Object.values(this.panels).forEach(panel => {
                this.makeDraggable(panel.panel);
            });
        }

        // 使单个元素可拖拽
        makeDraggable(element) {
            let pos1 = 0, pos2 = 0, pos3 = 0, pos4 = 0;
            const header = element.querySelector('.panel-header');

            if (header) {
                header.onmousedown = dragMouseDown;
            }

            function dragMouseDown(e) {
                e = e || window.event;
                e.preventDefault();
                pos3 = e.clientX;
                pos4 = e.clientY;
                document.onmouseup = closeDragElement;
                document.onmousemove = elementDrag;
                element.classList.add('dragging');

                // Convert centered transform to top/left positioning on first drag
                if (element.style.transform === '' || element.style.transform === 'translate(-50%, -50%)') {
                    const rect = element.getBoundingClientRect();
                    pos1 = pos3 - rect.left;
                    pos2 = pos4 - rect.top;
                    const top = rect.top + window.pageYOffset;
                    const left = rect.left + window.pageXOffset;

                    requestAnimationFrame(() => {
                        element.style.transform = 'none';
                        element.style.top = top + 'px';
                        element.style.left = left + 'px';
                    });
                } else {
                    const rect = element.getBoundingClientRect();
                    pos1 = pos3 - rect.left;
                    pos2 = pos4 - rect.top;
                }
            }

            function elementDrag(e) {
                e = e || window.event;
                e.preventDefault();
                pos1 = pos3 - e.clientX;
                pos2 = pos4 - e.clientY;
                pos3 = e.clientX;
                pos4 = e.clientY;
                element.style.top = (element.offsetTop - pos2) + "px";
                element.style.left = (element.offsetLeft - pos1) + "px";
            }

            function closeDragElement() {
                document.onmouseup = null;
                document.onmousemove = null;
                element.classList.remove('dragging');
            }
        }
    }

    // ==================== HTML模板生成器 ====================
    class HTMLGenerator {
        // 生成售后面板HTML
        static generateAfterSalesPanel() {
            const contentOptionsHTML = CONFIG.CONTENT_OPTIONS.map((item, index) => `
                <div class="radio-item btn-style">
                    <input type="radio" name="content" id="content-${index}" value="${item}"
                           class="panel-radio content-radio" ${index === 0 ? 'checked' : ''}>
                    <label for="content-${index}">${item}</label>
                </div>
            `).join('');

            const remarkOptionsHTML = CONFIG.REMARK_OPTIONS.map((item, index) => `
                <div class="radio-item btn-style">
                    <input type="radio" name="remark" id="remark-${index}" value="${item}"
                           class="panel-radio" ${index === 0 ? 'checked' : ''}>
                    <label for="remark-${index}">${item || '空'}</label>
                </div>
            `).join('');

            const responsibilityOptionsHTML = CONFIG.RESPONSIBILITY_OPTIONS.map((item, index) => `
                <div class="radio-item btn-style">
                    <input type="radio" name="responsibility" id="responsibility-${index}"
                           value="${item}" class="panel-radio">
                    <label for="responsibility-${index}">${item}</label>
                </div>
            `).join('');

            const staticShortcutsHTML = CONFIG.AMOUNT_SHORTCUTS.map((item, index) => `
                <span class="shortcut" data-value="${item}" id="amount-shortcut-${index}">${item}</span>
            `).join('');

            const amountShortcutsHTML = `
                ${staticShortcutsHTML}
                <span class="shortcut style-price-shortcut" data-value="0" id="custom-panel-30-percent-shortcut" style="color: #10b981;">30%</span>
                <span class="shortcut style-price-shortcut" data-value="0" id="custom-panel-50-percent-shortcut" style="color: #10b981;">50%</span>
            `;

            const inspectorShortcutsHTML = CONFIG.INSPECTOR_SHORTCUTS.map((item, index) => `
                <span class="shortcut" data-value="${item}" id="inspector-shortcut-${index}">${item}</span>
            `).join('');

            return `
                <div id="custom-panel">
                    <div class="panel-header">
                        <div class="panel-title" id="panel-title"></div>
                        <span class="panel-type">售后面板</span>
                        <span class="panel-close">&times;</span>
                    </div>
                    <div class="panel-body">
                        <div class="form-section">
                            <div class="form-item">
                                <label>款式:</label>
                                <div class="radio-group" id="product-style-group"></div>
                            </div>
                        </div>

                        <div class="form-section">
                            <div class="form-item">
                                <label for="compensation-amount">补偿金额:</label>
                                <input type="number" id="compensation-amount" class="panel-input"
                                       placeholder="可空,可点击下方快捷按钮" step="0.01" min="0" value="">
                                <div class="shortcut-group">${amountShortcutsHTML}</div>
                            </div>

                            <div class="form-item">
                                <label for="inspector">质检员:</label>
                                <div class="inspector-container">
                                    <input type="text" id="inspector" class="panel-input"
                                           placeholder="可空,输入数字下拉弹窗">
                                    <div class="inspector-dropdown" id="inspector-dropdown" style="display: none;"></div>
                                </div>
                                <div class="shortcut-group">${inspectorShortcutsHTML}</div>
                            </div>

                            <div class="form-item">
                                <label>内容:</label>
                                <div class="radio-group content-group">${contentOptionsHTML}</div>
                                <div class="other-content-group">
                                    <input type="radio" name="content" id="content-other" value="other"
                                           class="panel-radio content-radio">
                                    <input type="text" id="other-content-input" class="panel-input"
                                           placeholder="其他内容" disabled>
                                </div>
                            </div>
                        </div>

                        <div class="form-section">
                            <div class="form-item">
                                <label>责任:</label>
                                <div class="radio-group">${responsibilityOptionsHTML}</div>
                            </div>
                            <div class="form-item">
                                <label>备注:</label>
                                <div class="radio-group">${remarkOptionsHTML}</div>
                            </div>
                        </div>
                    </div>

                    <div class="form-section">
                            <div class="form-item">
                                <label>追加备注:</label>
                                <div class="checkbox-item">
                                    <input type="checkbox" id="custom-panel-append-remarks" checked>
                                    <label for="custom-panel-append-remarks">追加备注</label>
                                </div>
                            </div>
                    </div>
                    <div class="btn-group">
                        <button class="btn-send" id="panel-send-btn">发送到售后表</button>
                    </div>
                    <div class="bottom-labels">
                        <span id="shop-bottom" class="bottom-label"></span>
                    </div>
                </div>
            `;
        }

        // 生成挽单面板HTML
        static generateWandanPanel() {
            const typeOptionsHTML = CONFIG.WANDAN.TYPE_OPTIONS.map((item, index) => `
                <div class="radio-item btn-style">
                    <input type="radio" name="wandan-type" id="wandan-type-${index}"
                           value="${item}" class="panel-radio" ${index === 0 ? 'checked' : ''}>
                    <label for="wandan-type-${index}">${item}</label>
                </div>
            `).join('');

            const staticShortcutsHTML = CONFIG.AMOUNT_SHORTCUTS.map((item, index) => `
                <span class="shortcut" data-value="${item}" id="wandan-amount-shortcut-${index}">${item}</span>
            `).join('');

            const amountShortcutsHTML = `
                ${staticShortcutsHTML}
                <span class="shortcut style-price-shortcut" data-value="0" id="wandan-panel-30-percent-shortcut" style="color: #10b981;">30%</span>
                <span class="shortcut style-price-shortcut" data-value="0" id="wandan-panel-50-percent-shortcut" style="color: #10b981;">50%</span>
            `;

            const returnReasonShortcutsHTML = CONFIG.WANDAN.RETURN_REASON_SHORTCUTS.map((item, index) => `
                <span class="shortcut" data-value="${item}" id="return-reason-shortcut-${index}">${item}</span>
            `).join('');

            const resultShortcutsHTML = CONFIG.WANDAN.RESULT_SHORTCUTS.map((item, index) => `
                <span class="shortcut" data-value="${item}" id="result-shortcut-${index}">${item}</span>
            `).join('');

            return `
                <div id="wandan-panel">
                    <div class="panel-header">
                        <div class="panel-title" id="wandan-panel-title"></div>
                        <span class="panel-type">挽单面板</span>
                        <span class="panel-close">&times;</span>
                    </div>
                    <div class="panel-body">
                        <div class="form-section">
                            <div class="form-item">
                                <label>款式:</label>
                                <div class="radio-group" id="wandan-product-style-group"></div>
                            </div>
                            <div class="form-item">
                                <label for="wandan-amount">挽单金额:</label>
                                <input type="number" id="wandan-amount" class="panel-input"
                                       placeholder="自动获取" step="0.01" min="0">
                            </div>
                        </div>

                        <div class="form-section">
                            <div class="form-item">
                                <label>退款类型(type):</label>
                                <div class="radio-group">${typeOptionsHTML}</div>
                            </div>

                            <div class="form-item">
                                <label for="wandan-compensation-amount">补偿金额:</label>
                                <input type="number" id="wandan-compensation-amount" class="panel-input"
                                       placeholder="可空,可点击下方快捷按钮" step="0.01" min="0" value="">
                                <div class="shortcut-group">${amountShortcutsHTML}</div>
                            </div>
                        </div>

                        <div class="form-section">
                            <div class="form-item">
                                <label for="wandan-yryb">退货原因:</label>
                                <input type="text" id="wandan-yryb" class="panel-input"
                                       placeholder="请输入退货原因" value="不想要了">
                                <div class="shortcut-group">${returnReasonShortcutsHTML}</div>
                            </div>

                            <div class="form-item">
                                <label for="wandan-jpgo">结果:</label>
                                <input type="text" id="wandan-jpgo" class="panel-input"
                                       placeholder="请输入结果" value="留下">
                                <div class="shortcut-group">${resultShortcutsHTML}</div>
                            </div>

                            <div class="form-item">
                                <label>是否留下:</label>
                                <div class="checkbox-item">
                                    <input type="checkbox" id="wandan-done" checked>
                                    <label for="wandan-done">是</label>
                                </div>
                            </div>
                        </div>
                    </div>

                    <div class="form-section">
                            <div class="form-item">
                                <label>追加备注:</label>
                                <div class="checkbox-item">
                                    <input type="checkbox" id="wandan-append-remarks" checked>
                                    <label for="wandan-append-remarks">追加备注</label>
                                </div>
                            </div>
                    </div>

                    <div class="btn-group">
                        <button class="btn-send" id="wandan-panel-send-btn">发送到挽单表</button>
                    </div>
                    <div class="bottom-labels">
                        <span id="wandan-shop-bottom" class="bottom-label"></span>
                        <span id="wandan-name-bottom" class="bottom-label">${CONFIG.WANDAN.NAME}</span>
                    </div>
                </div>
            `;
        }

        // 生成错发面板HTML
        static generateCuofaPanel() {
            const responsibilityOptionsHTML = CONFIG.CUOFA.ZERF_OPTIONS.map((item, index) => `
                <div class="radio-item btn-style">
                    <input type="radio" name="cuofa-zerf" id="cuofa-zerf-${index}"
                           value="${item}" class="panel-radio" ${index === 0 ? 'checked' : ''}>
                    <label for="cuofa-zerf-${index}">${item}</label>
                </div>
            `).join('');

            const qkklShortcutsHTML = CONFIG.CUOFA.QKKL_SHORTCUTS.map((item, index) => `
                <span class="shortcut" data-value="${item}" id="cuofa-qkkl-shortcut-${index}">${item}</span>
            `).join('');

            const djhcShortcutsHTML = CONFIG.CUOFA.DJHC_SHORTCUTS.map((item, index) => `
                <span class="shortcut" data-value="${item}" id="cuofa-djhc-shortcut-${index}">${item}</span>
            `).join('');

            const staticShortcutsHTML = CONFIG.AMOUNT_SHORTCUTS.map((item, index) => `
                <span class="shortcut" data-value="${item}" id="cuofa-amount-shortcut-${index}">${item}</span>
            `).join('');

            const amountShortcutsHTML = `
                ${staticShortcutsHTML}
                <span class="shortcut style-price-shortcut" data-value="0" id="cuofa-panel-total-amount-shortcut" style="color: #10b981;">0.00</span>
                <span class="shortcut style-price-shortcut" data-value="0" id="cuofa-panel-style-amount-shortcut" style="color: #10b981;">0.00</span>
            `;

            return `
                <div id="cuofa-panel">
                    <div class="panel-header">
                        <div class="panel-title" id="cuofa-panel-title"></div>
                        <span class="panel-type">错发面板</span>
                        <span class="panel-close">&times;</span>
                    </div>
                    <div class="panel-body">
                        <div class="form-section">
                            <div class="form-item">
                                <label>款式:</label>
                                <div class="radio-group" id="cuofa-product-style-group"></div>
                            </div>
                        </div>

                        <div class="form-section">
                            <div class="form-item">
                                <label for="cuofa-compensation-amount">补偿金额:</label>
                                <input type="number" id="cuofa-compensation-amount" class="panel-input"
                                       placeholder="可空,可点击下方快捷按钮" step="0.01" min="0" value="">
                                <div class="shortcut-group">${amountShortcutsHTML}</div>
                            </div>

                            <div class="form-item">
                                <label for="cuofa-qkkl">错发情况:</label>
                                <div class="inspector-container">
                                    <input type="text" id="cuofa-qkkl" class="panel-input"
                                           placeholder="请输入错发情况">
                                    <div class="inspector-dropdown" id="cuofa-qkkl-dropdown" style="display: none;"></div>
                                </div>
                                <div class="shortcut-group">${qkklShortcutsHTML}</div>
                            </div>

                            <div class="form-item">
                                <label for="cuofa-djhc">单号/质检员:</label>
                                <div class="inspector-container">
                                    <input type="text" id="cuofa-djhc" class="panel-input"
                                           placeholder="物流单号或质检员">
                                    <div class="inspector-dropdown" id="cuofa-djhc-dropdown" style="display: none;"></div>
                                </div>
                                <div class="shortcut-group">
                                    <span class="shortcut" id="cuofa-djhc-tracking-number">物流单号</span>
                                    ${djhcShortcutsHTML}
                                </div>
                            </div>

                            <div class="form-item">
                                <label>责任:</label>
                                <div class="radio-group">${responsibilityOptionsHTML}</div>
                                <div class="other-content-group">
                                    <input type="radio" name="cuofa-zerf" id="cuofa-zerf-other" value="other"
                                           class="panel-radio">
                                    <input type="text" id="cuofa-zerf-other-input" class="panel-input"
                                           placeholder="其他责任人">
                                </div>
                            </div>
                        </div>
                    </div>

                    <div class="form-section">
                            <div class="form-item">
                                <label>追加备注:</label>
                                <div class="checkbox-item">
                                    <input type="checkbox" id="cuofa-panel-remarks" checked>
                                    <label for="cuofa-panel-remarks">追加备注</label>
                                </div>
                            </div>
                    </div>

                    <div class="btn-group">
                        <button class="btn-send" id="cuofa-panel-send-btn">发送到错发表</button>
                    </div>
                    <div class="bottom-labels">
                        <span id="cuofa-shop-bottom" class="bottom-label"></span>
                    </div>
                </div>
            `;
        }

        // 生成私域面板HTML
        static generateSiyuPanel() {
            const vekzOptionsHTML = CONFIG.SIYU.VEKZ_OPTIONS.map((item, index) => `
                <div class="radio-item btn-style">
                    <input type="radio" name="siyu-vekz" id="siyu-vekz-${index}" value="${item}"
                           class="panel-radio" ${index === 0 ? 'checked' : ''}>
                    <label for="siyu-vekz-${index}">${item}</label>
                </div>
            `).join('');

            const typeOptionsHTML = CONFIG.SIYU.TYPE_OPTIONS.map((item, index) => `
                <div class="radio-item btn-style">
                    <input type="radio" name="siyu-type" id="siyu-type-${index}" value="${item}"
                           class="panel-radio" ${index === 0 ? 'checked' : ''}>
                    <label for="siyu-type-${index}">${item}</label>
                </div>
            `).join('');

            const qShortcutsHTML = CONFIG.SIYU.Q_SHORTCUTS.map((item, index) => `
                <span class="shortcut" data-value="${item}" id="siyu-q-shortcut-${index}">${item}</span>
            `).join('');

            return `
                <div id="siyu-panel">
                    <div class="panel-header">
                        <div class="panel-title" id="siyu-panel-title"></div>
                        <span class="panel-type">私域面板</span>
                        <span class="panel-close">&times;</span>
                    </div>
                    <div class="panel-body">
                        <div class="form-section">
                            <div class="form-item">
                                <label>款式:</label>
                                <div class="radio-group" id="siyu-product-style-group"></div>
                            </div>
                        </div>

                        <div class="form-section">
                            <div class="form-item">
                                <label for="siyu-q">金额:</label>
                                <input type="number" id="siyu-q" class="panel-input"
                                       placeholder="请输入金额" step="0.01" min="0">
                                <div class="shortcut-group">${qShortcutsHTML}</div>
                            </div>

                            <div class="form-item">
                                <label>折扣:</label>
                                <div class="radio-group">${vekzOptionsHTML}</div>
                            </div>

                            <div class="form-item flex-row">
                                <div class="flex-item">
                                    <label for="siyu-uull">数量:</label>
                                    <input type="number" id="siyu-uull" class="panel-input">
                                </div>
                                <div class="flex-item">
                                    <label>类型:</label>
                                    <div class="radio-group">${typeOptionsHTML}</div>
                                </div>
                            </div>
                        </div>
                    </div>
                    <div class="btn-group">
                        <button class="btn-send" id="siyu-panel-send-btn">发送到私域表格</button>
                    </div>
                </div>
            `;
        }

        // 生成快递丢件面板HTML
        static generateDiuJianPanel() {
            const qkklOptionsHTML = CONFIG.DIUJIAN.QKKL_OPTIONS.map((item, index) => `
                <div class="radio-item btn-style">
                    <input type="radio" name="diujian-qkkl" id="diujian-qkkl-${index}"
                           value="${item}" class="panel-radio" ${index === 0 ? 'checked' : ''}>
                    <label for="diujian-qkkl-${index}">${item}</label>
                </div>
            `).join('');

            return `
                <div id="diujian-panel">
                    <div class="panel-header">
                        <div class="panel-title" id="diujian-panel-title"></div>
                        <span class="panel-type">快递丢件</span>
                        <span class="panel-close">&times;</span>
                    </div>

                    <div class="panel-body">
                        <div class="form-section">
                            <div class="flex-row">
                                <div class="flex-item">
                                    <label>原因:</label>
                                    <div class="radio-group">${qkklOptionsHTML}</div>
                                </div>
                                <div class="flex-item">
                                    <label for="diujian-jxvi">登记金额:</label>
                                    <input type="number" id="diujian-jxvi" class="panel-input"
                                           placeholder="自动获取" step="0.01" min="0">
                                </div>
                            </div>
                        </div>
                    </div>

                    <div class="form-section">
                        <div class="flex-row">
                            <div class="flex-item">
                                <span>运单号:</span>
                                <span id="diujian-logistics-company" class="bottom-label"></span>
                                <span id="diujian-tracking-number" class="bottom-label"></span>
                            </div>
                        </div>
                    </div>

                    <div class="form-section">
                        <div class="form-item">
                            <label>追加备注:</label>
                            <div class="checkbox-item">
                                <input type="checkbox" id="diujian-append-remarks" checked>
                                <label for="diujian-append-remarks">追加备注</label>
                            </div>
                        </div>
                    </div>

                    <div class="btn-group">
                        <button class="btn-send" id="diujian-panel-send-btn">发送到快递丢件表</button>
                    </div>
                </div>
            `;
        }

        // 注入所有面板HTML到页面
        static injectHTML() {
            const html = this.generateAfterSalesPanel() + this.generateWandanPanel() + this.generateCuofaPanel() + this.generateSiyuPanel() + this.generateDiuJianPanel();
            document.body.insertAdjacentHTML('beforeend', html);
        }
    }

    // ==================== 应用入口点 ====================
    function init() {

        //获取当前页面的完整URL地址
        const currentHref = window.location.href;
        if (!currentHref.includes(CONFIG.TARGET_DOMAINS.W) && !currentHref.includes(CONFIG.TARGET_DOMAINS.WWW)) {
            console.log("非目标页面，脚本不执行");
            console.log("========== 订单推送脚本加载结束 ==========");
            return;
        }
        console.log("售后自用插件初始化开始");

        // 将面板的HTML结构动态添加到页面中
        HTMLGenerator.injectHTML();

        // Initialize panel manager
        const panelManager = new PanelManager();

        console.log("售后自用插件初始化完成");
    }

    // Start the plugin when DOM is ready
    if (document.readyState === 'loading') {
        document.addEventListener('DOMContentLoaded', init);
    } else {
        init();
    }

})();
