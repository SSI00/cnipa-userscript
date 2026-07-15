// ==UserScript==
// @name         CNIPA 专利信息批量查询
// @namespace    http://tampermonkey.net/
// @version      1.8
// @description  国知局专利信息批量查询：申请人/代理机构/最近缴费人/最近缴费种类/法律状态/案件状态，支持 Excel 上传导出、失败重查
// @author       CNIPA_Fee_Collector
// @license      MIT
// @match        https://cpquery.cponline.cnipa.gov.cn/chinesepatent/*
// @run-at       document-start
// @require      https://cdn.sheetjs.com/xlsx-0.20.3/package/dist/xlsx.full.min.js
// @grant        none
// ==/UserScript==

(function () {
    'use strict';

    window.__cnipaAuth = { token: null, userType: null };

    // ---------- hook XHR 和 fetch，捕获 Authorization ----------
    (function hookXHR() {
        const origSetHeader = XMLHttpRequest.prototype.setRequestHeader;
        XMLHttpRequest.prototype.setRequestHeader = function (header, value) {
            const h = header.toLowerCase();
            if (h === 'authorization' && value.startsWith('Bearer ')) {
                window.__cnipaAuth.token = value.substring(7);
            }
            if (h === 'usertype') {
                window.__cnipaAuth.userType = value;
            }
            return origSetHeader.call(this, header, value);
        };
    })();

    (function hookFetch() {
        const origFetch = window.fetch;
        window.fetch = function (input, init) {
            try {
                const headers = (init && init.headers) || (input && input.headers);
                if (headers) {
                    let auth, userType;
                    if (headers instanceof Headers) {
                        auth = headers.get('authorization');
                        userType = headers.get('usertype');
                    } else if (typeof headers === 'object') {
                        auth = headers['authorization'] || headers['Authorization'];
                        userType = headers['usertype'] || headers['userType'];
                    }
                    if (auth && auth.startsWith('Bearer ')) {
                        window.__cnipaAuth.token = auth.substring(7);
                    }
                    if (userType) {
                        window.__cnipaAuth.userType = userType;
                    }
                }
            } catch (e) {}
            return origFetch.apply(this, arguments);
        };
    })();

    // ---------- 申请号格式清洗 ----------
    function normalizeAppNo(raw) {
        let s = String(raw == null ? '' : raw).trim().replace(/^cn\s*/i, '').replace(/[\s.]/g, '');
        if (!/^\d{12}[\dXx]$/.test(s)) return null;
        return s.toUpperCase();
    }

    // ---------- 字段配置（key → 显示名、所属 API、提取函数） ----------
    // 每个字段返回字符串；未查到返回 ''
    const FIELDS = {
        applicant: {
            label: '申请人',
            api: 'sqxx',
            extract(data) {
                const list = data.data && data.data.shenqingren && data.data.shenqingren.shenqingrenList;
                return (list && list.length) ? (list[0].shenqingrxm || '').trim() : '';
            }
        },
        agency: {
            label: '代理机构',
            api: 'sqxx',
            extract(data) {
                const d = data.data && data.data.dailijg;
                return (d && d.isShow && d.dailijgList && d.dailijgList.length) ? (d.dailijgList[0].dailijgdm || '').trim() : '';
            }
        },
        payer: {
            label: '最近缴费人',
            api: 'fyxx',
            extract(data) {
                const list = data.data && data.data.yijiaofei && data.data.yijiaofei.svYijfList;
                return (list && list.length) ? (list[0].yijiaofjfrxm || '').trim() : '';
            }
        },
        feeType: {
            label: '最近缴费种类',
            api: 'fyxx',
            extract(data) {
                const list = data.data && data.data.yijiaofei && data.data.yijiaofei.svYijfList;
                return (list && list.length) ? (list[0].yijiaofjfzlmc || '').trim() : '';
            }
        },
        legalStatus: {
            label: '法律状态',
            api: 'sqxx',
            extract(data) {
                const z = data.data && data.data.zhuluxmxx;
                if (!z || !z.isShow) return '未公开';
                const v = ((z.zhuluxmxx || {}).falvzt || '').trim();
                return (v && v !== '--') ? v : '未公开';
            }
        },
        caseStatus: {
            label: '案件状态',
            api: 'sqxx',
            extract(data) {
                const z = data.data && data.data.zhuluxmxx;
                if (!z || !z.isShow) return '';
                const v = ((z.zhuluxmxx || {}).anjianywzt || '').trim();
                return (v && v !== '--') ? v : '';
            }
        }
    };
    const FIELD_ORDER_DEFAULT = ['applicant', 'agency', 'payer', 'feeType', 'legalStatus', 'caseStatus'];

    // ---------- 全局状态 ----------
    const state = {
        fieldOrder: FIELD_ORDER_DEFAULT.slice(),   // 当前字段顺序（含未勾选）
        checked: new Set(FIELD_ORDER_DEFAULT),      // 勾选的字段
        rows: [],   // [{original, cleaned, results:{fieldKey:value}, status:'pending'|'ok'|'fail'|'skip', note}]
        running: false,
        paused: false
    };

    const APIS = {
        sqxx: '/api/view/gn/sqxx',
        fyxx: '/api/view/gn/fyxx'
    };

    // ---------- 更新日志（新版本追加到最前面） ----------
    const CHANGELOG = [
        { version: '1.8', date: '2026-07-15', items: ['修复折叠后白色背景板残留的问题，折叠后只剩标题栏'] },
        { version: '1.7', date: '2026-07-15', items: ['按钮体系统一为 3 种样式（主按钮蓝底/次按钮蓝边/文字按钮）', 'tab 切换改为下划线选中样式', '选择文件按钮美化，隐藏浏览器原生样式'] },
        { version: '1.6', date: '2026-07-15', items: ['状态文字单独一行显示，不再换行或撑宽窗口'] },
        { version: '1.5', date: '2026-07-15', items: ['新增暂停/继续功能', '按钮防换行，宽度自适应', '窗口大小可拖拽调整（最小 460×400）', '导出文件名改为 CNIPA查询结果_条数_时间戳.xlsx'] },
        { version: '1.4', date: '2026-07-15', items: ['结果预览改为智能滚动，不再强制跳到底部'] },
        { version: '1.3', date: '2026-07-15', items: ['面板滚动条统一为蓝色加宽样式'] },
        { version: '1.2', date: '2026-07-15', items: ['按钮蓝色统一为国知局标题蓝 #3664D1'] },
        { version: '1.1', date: '2026-07-15', items: ['新增重查失败项功能', '结果预览滚动条加宽'] },
        { version: '1.0', date: '2026-07-15', items: ['多选查询字段（申请人/代理机构/最近缴费人/最近缴费种类/法律状态/案件状态），支持拖拽排序', '支持粘贴和上传 xlsx 两种输入方式，内置模板下载', '查询结果导出 xlsx', '窗口可拖动'] },
    ];

    function init() {
        injectStyle();
        const panel = document.createElement('div');
        panel.id = 'cnipa-panel';
        panel.innerHTML = `
            <div id="cnipa-header">
                <b>CNIPA 专利信息批量查询 v1.8</b>
                <span id="cnipa-header-btns">
                    <span id="cnipa-changelog-btn" title="更新日志">ⓘ</span>
                    <span id="cnipa-collapse">—</span>
                </span>
            </div>
            <div id="cnipa-body">
                <div id="cnipa-auth-status" class="auth-wait">⏳ 等待获取登录态... 请先在页面上<b>手动搜索一次</b></div>

                <div class="section-title">查询字段（勾选 + 拖动排序）：</div>
                <ul id="cnipa-fields"></ul>

                <div class="section-title">申请号来源：</div>
                <div class="input-tabs">
                    <button id="tab-paste" class="tab active">粘贴</button>
                    <button id="tab-upload" class="tab">上传 Excel</button>
                </div>
                <div id="pane-paste">
                    <div id="cnipa-input" contenteditable="true" placeholder="每行一个申请号，支持 CN202610662164.3 / 202610662164.3 / 2026106621643"></div>
                </div>
                <div id="pane-upload" style="display:none;">
                    <div class="upload-row">
                        <button id="cnipa-download-tpl" class="btn-text">下载模板</button>
                        <span class="hint">模板为 xlsx，A 列填申请号</span>
                    </div>
                    <label for="cnipa-file" class="btn btn-file">选择文件</label>
                    <input type="file" id="cnipa-file" accept=".xlsx" />
                    <div id="cnipa-file-info" class="hint"></div>
                </div>

                <div class="btn-row">
                    <button id="cnipa-start" class="btn btn-primary">开始查询</button>
                    <button id="cnipa-pause" class="btn" disabled>暂停</button>
                    <button id="cnipa-retry-failed" class="btn" disabled>重查失败项</button>
                    <button id="cnipa-export" class="btn btn-primary" disabled>导出 Excel</button>
                </div>
                <div id="cnipa-status"></div>
                <div id="cnipa-progress"><div id="cnipa-progress-bar"></div></div>

                <div class="section-title">结果预览：</div>
                <div id="cnipa-output-wrap"><table id="cnipa-output"></table></div>
            </div>
            <div id="cnipa-resize-handle"></div>
            <div id="cnipa-changelog-overlay" style="display:none;">
                <div id="cnipa-changelog-box">
                    <div id="cnipa-changelog-head">
                        <b>更新日志</b>
                        <span id="cnipa-changelog-close">✕</span>
                    </div>
                    <div id="cnipa-changelog-list"></div>
                </div>
            </div>
        `;
        document.body.appendChild(panel);

        buildFieldList();
        bindTabs();
        bindDrag();
        bindActions();
        watchAuth();
    }

    // ---------- 样式 ----------
    function injectStyle() {
        const css = `
        #cnipa-panel { position:fixed; top:70px; right:20px; width:460px; z-index:999999;
            min-width:460px; min-height:400px;
            background:#fff; border:1px solid #bbb; border-radius:8px; overflow:hidden;
            box-shadow:0 4px 20px rgba(0,0,0,.25); font-size:13px; font-family:"Microsoft YaHei",sans-serif; color:#222; }
        /* 折叠状态：只剩标题栏，去掉最小高度和背景板 */
        #cnipa-panel.collapsed { min-height:0; height:auto !important; width:460px !important; }
        #cnipa-panel.collapsed #cnipa-resize-handle { display:none; }
        #cnipa-panel.collapsed #cnipa-header { border-radius:8px; }
        #cnipa-resize-handle { position:absolute; right:0; bottom:0; width:18px; height:18px;
            cursor:nwse-resize; z-index:10;
            background:linear-gradient(135deg, transparent 50%, #3664D1 50%); }
        #cnipa-header { padding:9px 12px; background:#3664D1; color:#fff; border-radius:8px 8px 0 0;
            cursor:move; display:flex; justify-content:space-between; align-items:center; user-select:none; }
        #cnipa-collapse, #cnipa-changelog-btn { cursor:pointer; font-size:16px; width:22px; height:22px; line-height:20px;
            text-align:center; border-radius:4px; transition:background .15s; display:inline-block; }
        #cnipa-collapse:hover, #cnipa-changelog-btn:hover { background:rgba(255,255,255,.2); }
        #cnipa-changelog-btn { font-size:14px; margin-right:4px; }
        /* 更新日志浮层 */
        #cnipa-changelog-overlay { position:absolute; top:0; left:0; right:0; bottom:0;
            background:rgba(0,0,0,.3); z-index:20; display:flex; align-items:center; justify-content:center; }
        #cnipa-changelog-box { background:#fff; border-radius:8px; width:85%; max-height:80%;
            display:flex; flex-direction:column; box-shadow:0 4px 20px rgba(0,0,0,.3); }
        #cnipa-changelog-head { padding:10px 14px; background:#3664D1; color:#fff; border-radius:8px 8px 0 0;
            display:flex; justify-content:space-between; align-items:center; }
        #cnipa-changelog-close { cursor:pointer; width:22px; height:22px; line-height:20px; text-align:center;
            border-radius:4px; transition:background .15s; }
        #cnipa-changelog-close:hover { background:rgba(255,255,255,.2); }
        #cnipa-changelog-list { padding:10px 14px; overflow-y:auto; font-size:12px; color:#333; }
        #cnipa-changelog-list .cl-version { font-weight:bold; color:#3664D1; margin:8px 0 3px; font-size:13px; }
        #cnipa-changelog-list .cl-version:first-child { margin-top:0; }
        #cnipa-changelog-list .cl-date { color:#999; font-weight:normal; font-size:11px; margin-left:6px; }
        #cnipa-changelog-list ul { margin:0 0 4px; padding-left:18px; }
        #cnipa-changelog-list li { margin:2px 0; line-height:1.5; }
        #cnipa-body { padding:10px 12px; max-height:70vh; overflow-y:auto; }
        .section-title { color:#555; margin:10px 0 5px; font-weight:bold; }
        .auth-wait { color:#d32f2f; font-size:12px; margin-bottom:4px; }
        .auth-ok { color:#388e3c; font-size:12px; margin-bottom:4px; }
        #cnipa-fields { list-style:none; margin:0; padding:0; border:1px solid #ddd; border-radius:6px; overflow:hidden; }
        #cnipa-fields li { display:flex; align-items:center; padding:6px 8px; background:#fafafa; border-bottom:1px solid #eee; cursor:grab; }
        #cnipa-fields li:last-child { border-bottom:none; }
        #cnipa-fields li.dragging { opacity:.4; }
        #cnipa-fields li .drag-handle { color:#999; margin-right:8px; cursor:grab; user-select:none; }
        #cnipa-fields li label { flex:1; cursor:pointer; user-select:none; }
        /* tab 切换：现代下划线选中样式 */
        .input-tabs { display:flex; gap:16px; margin-bottom:8px; border-bottom:1px solid #e0e0e0; }
        .tab { padding:5px 2px; border:none; background:none; cursor:pointer; font-size:13px; color:#666;
            border-bottom:2px solid transparent; margin-bottom:-1px; transition:color .15s; }
        .tab:hover { color:#3664D1; }
        .tab.active { color:#3664D1; border-bottom-color:#3664D1; font-weight:bold; }
        #cnipa-input { width:100%; box-sizing:border-box; min-height:80px; max-height:130px; overflow-y:auto;
            border:1px solid #ccc; border-radius:4px; padding:6px; background:#fff;
            font-family:Consolas,monospace; font-size:13px; outline:none; }
        #cnipa-input:empty:before { content:attr(placeholder); color:#aaa; }
        .upload-row { display:flex; align-items:center; gap:8px; margin-bottom:8px; }
        .hint { color:#888; font-size:12px; }
        #cnipa-file { display:none; }
        /* ===== 统一按钮体系 ===== */
        /* 基础：所有按钮同高 28px、字号 13px、圆角 4px、不换行 */
        .btn { display:inline-block; height:28px; line-height:26px; padding:0 14px; box-sizing:border-box;
            font-size:13px; border-radius:4px; cursor:pointer; white-space:nowrap; flex-shrink:0;
            background:#fff; color:#3664D1; border:1px solid #3664D1; transition:background .15s,color .15s,border-color .15s; }
        .btn:hover:not(:disabled) { background:#eef3fd; }
        .btn:disabled { color:#9e9e9e; border-color:#d0d0d0; background:#f7f7f7; cursor:not-allowed; }
        /* 主按钮：蓝底白字，用于关键动作（开始查询、导出 Excel） */
        .btn-primary { background:#3664D1; color:#fff; border-color:#3664D1; }
        .btn-primary:hover:not(:disabled) { background:#2a4fad; border-color:#2a4fad; }
        .btn-primary:disabled { background:#c5cfe8; border-color:#c5cfe8; color:#fff; }
        /* 文字按钮：无边框蓝字，用于低频辅助操作（下载模板） */
        .btn-text { border:none; background:none; color:#3664D1; padding:0 4px; height:auto; line-height:1.5; }
        .btn-text:hover:not(:disabled) { background:none; text-decoration:underline; }
        /* 文件选择按钮（label 伪装） */
        .btn-file { line-height:26px; }
        .btn-row { margin:10px 0 6px; display:flex; align-items:center; gap:8px; flex-wrap:wrap; }
        /* 状态文字单独一行，不换行，超出省略号 */
        #cnipa-status { display:block; color:#666; font-size:12px; margin:0 0 5px;
            white-space:nowrap; overflow:hidden; text-overflow:ellipsis; }
        #cnipa-progress { height:6px; background:#eee; border-radius:3px; margin-bottom:8px; }
        #cnipa-progress-bar { height:100%; width:0; background:#3664D1; border-radius:3px; transition:width .2s; }
        #cnipa-output-wrap { max-height:220px; overflow:auto; border:1px solid #ddd; border-radius:4px; }
        #cnipa-output { border-collapse:collapse; width:100%; font-size:12px; }
        #cnipa-output th, #cnipa-output td { border:1px solid #e0e0e0; padding:4px 6px; text-align:left; white-space:nowrap; }
        #cnipa-output th { background:#f5f5f5; position:sticky; top:0; }
        #cnipa-output tr.fail td { color:#c62828; }
        #cnipa-output tr.skip td { color:#f57c00; }
        /* 加宽结果预览区滚动条，方便点按；蓝色系与标题一致 */
        #cnipa-output-wrap::-webkit-scrollbar { height:14px; width:14px; }
        #cnipa-output-wrap::-webkit-scrollbar-track { background:#e8edf9; border-radius:7px; }
        #cnipa-output-wrap::-webkit-scrollbar-thumb { background:#3664D1; border-radius:7px; border:2px solid #e8edf9; }
        #cnipa-output-wrap::-webkit-scrollbar-thumb:hover { background:#2a4fad; }
        #cnipa-output-wrap { scrollbar-width:thick; scrollbar-color:#3664D1 #e8edf9; }
        /* 面板 body 滚动条也用同款蓝色样式 */
        #cnipa-body::-webkit-scrollbar { height:14px; width:14px; }
        #cnipa-body::-webkit-scrollbar-track { background:#e8edf9; border-radius:7px; }
        #cnipa-body::-webkit-scrollbar-thumb { background:#3664D1; border-radius:7px; border:2px solid #e8edf9; }
        #cnipa-body::-webkit-scrollbar-thumb:hover { background:#2a4fad; }
        #cnipa-body { scrollbar-width:thick; scrollbar-color:#3664D1 #e8edf9; }
        `;
        const style = document.createElement('style');
        style.textContent = css;
        document.head.appendChild(style);
    }

    // ---------- 字段多选 + 拖拽排序 ----------
    function buildFieldList() {
        const ul = document.getElementById('cnipa-fields');
        ul.innerHTML = '';
        state.fieldOrder.forEach(key => {
            const li = document.createElement('li');
            li.draggable = true;
            li.dataset.key = key;
            li.innerHTML = `<span class="drag-handle">☰</span><label><input type="checkbox" data-key="${key}" ${state.checked.has(key) ? 'checked' : ''}> ${FIELDS[key].label}</label>`;
            ul.appendChild(li);
        });
        ul.querySelectorAll('input[type=checkbox]').forEach(cb => {
            cb.onchange = () => {
                const k = cb.dataset.key;
                if (cb.checked) state.checked.add(k);
                else state.checked.delete(k);
            };
        });
        // HTML5 拖拽排序
        let dragEl = null;
        ul.querySelectorAll('li').forEach(li => {
            li.addEventListener('dragstart', e => {
                dragEl = li;
                li.classList.add('dragging');
                e.dataTransfer.effectAllowed = 'move';
            });
            li.addEventListener('dragend', () => {
                li.classList.remove('dragging');
                dragEl = null;
                // 更新 fieldOrder
                state.fieldOrder = Array.from(ul.querySelectorAll('li')).map(l => l.dataset.key);
            });
            li.addEventListener('dragover', e => {
                e.preventDefault();
                if (!dragEl || dragEl === li) return;
                const rect = li.getBoundingClientRect();
                const after = (e.clientY - rect.top) > rect.height / 2;
                ul.insertBefore(dragEl, after ? li.nextSibling : li);
            });
        });
    }

    function getSelectedFields() {
        return state.fieldOrder.filter(k => state.checked.has(k));
    }

    // ---------- 输入方式切换 ----------
    function bindTabs() {
        const tp = document.getElementById('tab-paste');
        const tu = document.getElementById('tab-upload');
        const pp = document.getElementById('pane-paste');
        const pu = document.getElementById('pane-upload');
        tp.onclick = () => { tp.classList.add('active'); tu.classList.remove('active'); pp.style.display = ''; pu.style.display = 'none'; };
        tu.onclick = () => { tu.classList.add('active'); tp.classList.remove('active'); pu.style.display = ''; pp.style.display = 'none'; };
    }

    // ---------- 窗口拖动 ----------
    function bindDrag() {
        const panel = document.getElementById('cnipa-panel');
        const header = document.getElementById('cnipa-header');
        let sx, sy, ox, oy, dragging = false;
        header.onmousedown = e => {
            if (e.target.id === 'cnipa-collapse' || e.target.id === 'cnipa-changelog-btn') return;
            dragging = true; sx = e.clientX; sy = e.clientY;
            const r = panel.getBoundingClientRect(); ox = r.left; oy = r.top;
            e.preventDefault();
        };
        document.onmousemove = e => {
            if (!dragging) return;
            panel.style.left = (ox + e.clientX - sx) + 'px';
            panel.style.top = (oy + e.clientY - sy) + 'px';
            panel.style.right = 'auto';
        };
        document.onmouseup = () => dragging = false;

        document.getElementById('cnipa-collapse').onclick = function () {
            const body = document.getElementById('cnipa-body');
            const panel = document.getElementById('cnipa-panel');
            const collapsed = body.style.display === 'none';
            body.style.display = collapsed ? '' : 'none';
            panel.classList.toggle('collapsed', !collapsed);
            this.textContent = collapsed ? '—' : '+';
            // 折叠时清掉拖拽设置的固定高度，展开时恢复 body 滚动高度
            if (!collapsed) {
                panel.style.height = 'auto';
            } else {
                body.style.maxHeight = '70vh';
            }
        };

        // 窗口大小拖拽（右下角手柄）
        const handle = document.getElementById('cnipa-resize-handle');
        let rsx, rsy, rsw, rsh, resizing = false;
        handle.onmousedown = e => {
            resizing = true; rsx = e.clientX; rsy = e.clientY;
            rsw = panel.offsetWidth; rsh = panel.offsetHeight;
            e.preventDefault();
            e.stopPropagation();
        };
        document.addEventListener('mousemove', e => {
            if (!resizing) return;
            const newW = Math.max(460, rsw + e.clientX - rsx);
            const newH = Math.max(400, rsh + e.clientY - rsy);
            panel.style.width = newW + 'px';
            panel.style.height = newH + 'px';
            // body 高度随窗口调整
            const body = document.getElementById('cnipa-body');
            body.style.maxHeight = (newH - 50) + 'px';
        });
        document.addEventListener('mouseup', () => resizing = false);
    }

    // ---------- 登录态监听 ----------
    function watchAuth() {
        const el = document.getElementById('cnipa-auth-status');
        setInterval(() => {
            if (window.__cnipaAuth.token) {
                el.innerHTML = '✅ 登录态已获取，可以开始查询';
                el.className = 'auth-ok';
            } else {
                el.innerHTML = '⏳ 等待获取登录态... 请先在页面上<b>手动搜索一次</b>';
                el.className = 'auth-wait';
            }
        }, 500);
    }

    // ---------- 模板下载 ----------
    function downloadTemplate() {
        const ws = XLSX.utils.aoa_to_sheet([['申请号']]);
        const wb = XLSX.utils.book_new();
        XLSX.utils.book_append_sheet(wb, ws, 'Sheet1');
        XLSX.writeFile(wb, '查询模板.xlsx');
    }

    // ---------- Excel 上传解析 ----------
    function parseUpload(file) {
        return new Promise((resolve, reject) => {
            const reader = new FileReader();
            reader.onload = e => {
                try {
                    const data = new Uint8Array(e.target.result);
                    const wb = XLSX.read(data, { type: 'array' });
                    const ws = wb.Sheets[wb.SheetNames[0]];
                    const aoa = XLSX.utils.sheet_to_json(ws, { header: 1, defval: '' });
                    // 第一行表头，从第二行开始取 A 列
                    const list = [];
                    for (let i = 1; i < aoa.length; i++) {
                        const v = aoa[i][0];
                        if (v !== '' && v != null) list.push(String(v).trim());
                    }
                    resolve(list);
                } catch (err) { reject(err); }
            };
            reader.onerror = reject;
            reader.readAsArrayBuffer(file);
        });
    }

    // ---------- 调 API（带重试） ----------
    async function callApi(apiKey, appNo, retryCount = 0) {
        const auth = window.__cnipaAuth;
        if (!auth.token) throw new Error('no-auth');
        try {
            const resp = await fetch(APIS[apiKey], {
                method: 'POST',
                headers: {
                    'Content-Type': 'application/json;charset=utf-8',
                    'Accept': 'application/json, text/plain, */*',
                    'Authorization': 'Bearer ' + auth.token,
                    'userType': auth.userType || ''
                },
                body: JSON.stringify({ zhuanlisqh: appNo }),
                credentials: 'include'
            });
            const text = await resp.text();
            if (resp.status === 401) {
                window.__cnipaAuth.token = null;
                throw new Error('auth-expired');
            }
            let data;
            try { data = JSON.parse(text); } catch (e) {
                if (retryCount === 0) {
                    await new Promise(r => setTimeout(r, 800));
                    return callApi(apiKey, appNo, retryCount + 1);
                }
                throw new Error(`HTTP ${resp.status} 非JSON`);
            }
            if (data.code !== 200) throw new Error(`code=${data.code}`);
            return data;
        } catch (e) {
            if (e.message === 'auth-expired' || e.message === 'no-auth') throw e;
            if (retryCount === 0) {
                await new Promise(r => setTimeout(r, 800));
                return callApi(apiKey, appNo, retryCount + 1);
            }
            throw e;
        }
    }

    // ---------- 处理单个申请号 ----------
    async function processRow(row, fields) {
        // 需要调哪些 API
        const needSqxx = fields.some(f => FIELDS[f].api === 'sqxx');
        const needFyxx = fields.some(f => FIELDS[f].api === 'fyxx');
        let sqxxData = null, fyxxData = null;
        if (needSqxx) sqxxData = await callApi('sqxx', row.cleaned);
        if (needFyxx) fyxxData = await callApi('fyxx', row.cleaned);
        fields.forEach(f => {
            const src = FIELDS[f].api === 'sqxx' ? sqxxData : fyxxData;
            try {
                row.results[f] = src ? FIELDS[f].extract(src) : '';
            } catch (e) {
                row.results[f] = '';
            }
        });
    }

    // ---------- 渲染结果表格 ----------
    function renderOutput(fields) {
        const table = document.getElementById('cnipa-output');
        const head = ['申请号', ...fields.map(f => FIELDS[f].label), '备注'];
        let html = '<thead><tr>' + head.map(h => `<th>${h}</th>`).join('') + '</tr></thead><tbody>';
        state.rows.forEach(r => {
            const cls = r.status === 'fail' ? 'fail' : (r.status === 'skip' ? 'skip' : '');
            const cells = [r.original, ...fields.map(f => r.results[f] || ''), r.note || ''];
            html += `<tr class="${cls}">` + cells.map(c => `<td>${escapeHtml(c)}</td>`).join('') + '</tr>';
        });
        html += '</tbody>';
        table.innerHTML = html;
    }

    function escapeHtml(s) {
        return String(s == null ? '' : s).replace(/[&<>"]/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' }[c]));
    }

    // ---------- 导出 Excel ----------
    function exportExcel(fields) {
        const head = ['申请号', ...fields.map(f => FIELDS[f].label)];
        const aoa = [head];
        state.rows.forEach(r => {
            aoa.push([r.original, ...fields.map(f => r.results[f] || '')]);
        });
        const ws = XLSX.utils.aoa_to_sheet(aoa);
        const wb = XLSX.utils.book_new();
        XLSX.utils.book_append_sheet(wb, ws, '查询结果');

        // 文件名：CNIPA查询结果_条数_YYMMDDHHMMSS.xlsx
        // 条数 = 已成功查询的行数（status 为 ok）
        const okCount = state.rows.filter(r => r.status === 'ok').length;
        const now = new Date();
        const pad = n => String(n).padStart(2, '0');
        const ts = String(now.getFullYear()).slice(2) + pad(now.getMonth() + 1) + pad(now.getDate())
                 + pad(now.getHours()) + pad(now.getMinutes()) + pad(now.getSeconds());
        const filename = `CNIPA查询结果_${okCount}条_${ts}.xlsx`;
        XLSX.writeFile(wb, filename);
    }

    // ---------- 绑定按钮动作 ----------
    function bindActions() {
        document.getElementById('cnipa-download-tpl').onclick = downloadTemplate;

        document.getElementById('cnipa-file').onchange = async e => {
            const file = e.target.files[0];
            const info = document.getElementById('cnipa-file-info');
            if (!file) { info.textContent = ''; return; }
            try {
                const list = await parseUpload(file);
                info.textContent = `已读取 ${list.length} 个申请号（${file.name}）`;
                info.dataset.appnos = JSON.stringify(list);
            } catch (err) {
                info.textContent = `解析失败: ${err.message}`;
                delete info.dataset.appnos;
            }
        };

        document.getElementById('cnipa-start').onclick = startQuery;
        document.getElementById('cnipa-pause').onclick = togglePause;
        document.getElementById('cnipa-retry-failed').onclick = retryFailed;
        document.getElementById('cnipa-export').onclick = () => exportExcel(getSelectedFields());

        // 更新日志浮层
        const overlay = document.getElementById('cnipa-changelog-overlay');
        document.getElementById('cnipa-changelog-btn').onclick = () => {
            const list = document.getElementById('cnipa-changelog-list');
            list.innerHTML = CHANGELOG.map(v =>
                `<div class="cl-version">v${v.version}<span class="cl-date">${v.date}</span></div><ul>` +
                v.items.map(i => `<li>${i}</li>`).join('') + '</ul>'
            ).join('');
            overlay.style.display = 'flex';
        };
        document.getElementById('cnipa-changelog-close').onclick = () => overlay.style.display = 'none';
        overlay.onclick = e => { if (e.target === overlay) overlay.style.display = 'none'; };
    }

    // ---------- 暂停 / 继续 ----------
    function togglePause() {
        if (!state.running) return;
        state.paused = !state.paused;
        const btn = document.getElementById('cnipa-pause');
        btn.textContent = state.paused ? '继续' : '暂停';
        if (state.paused) {
            document.getElementById('cnipa-status').textContent = '已暂停';
        }
        updateButtons();   // 暂停/继续时刷新导出按钮可用状态
    }

    // ---------- 更新按钮状态 ----------
    function updateButtons() {
        const hasRows = state.rows.length > 0;
        const hasFailed = state.rows.some(r => r.status === 'fail');
        // 导出：查询中且未暂停（正在跑，数据在变）时禁用；暂停或未在查询且有数据时可用
        document.getElementById('cnipa-export').disabled = !hasRows || (state.running && !state.paused);
        document.getElementById('cnipa-retry-failed').disabled = !hasFailed || state.running;
        document.getElementById('cnipa-start').disabled = state.running;
        // 暂停按钮：查询中可用，否则置灰；非查询中恢复"暂停"文字
        const pauseBtn = document.getElementById('cnipa-pause');
        pauseBtn.disabled = !state.running;
        if (!state.running) {
            state.paused = false;
            pauseBtn.textContent = '暂停';
        }
    }

    // ---------- 批量执行（开始查询和重查失败共用） ----------
    async function runBatch(rowsToProcess, fields) {
        const statusEl = document.getElementById('cnipa-status');
        const bar = document.getElementById('cnipa-progress-bar');
        state.running = true;
        updateButtons();
        renderOutput(fields);

        const total = rowsToProcess.length;
        let stopped = false;
        for (let i = 0; i < total; i++) {
            if (stopped) break;
            // 暂停检查：暂停时在此等待，直到继续
            while (state.paused && state.running) {
                await new Promise(r => setTimeout(r, 200));
            }
            if (!state.running) break;
            const row = rowsToProcess[i];
            statusEl.textContent = `${i + 1}/${total} ${row.cleaned}`;
            try {
                await processRow(row, fields);
                row.status = 'ok';
                row.note = '';
            } catch (e) {
                if (e.message === 'auth-expired' || e.message === 'no-auth') {
                    row.status = 'fail';
                    row.note = '登录态过期';
                    statusEl.textContent = '登录态过期，已停止';
                    alert('登录态已过期！\n\n请在页面上重新搜索一次，然后重新开始查询。');
                    stopped = true;
                } else {
                    row.status = 'fail';
                    row.note = e.message;
                }
            }
            bar.style.width = ((i + 1) / total * 100) + '%';
            // 渲染前记录用户是否本来就在底部附近（30px 容差）
            const wrap = document.getElementById('cnipa-output-wrap');
            const wasAtBottom = wrap.scrollTop + wrap.clientHeight >= wrap.scrollHeight - 30;
            renderOutput(fields);
            // 只有用户本来就在底部时才跟随滚动，否则保持用户当前位置
            if (wasAtBottom) wrap.scrollTop = wrap.scrollHeight;
            await new Promise(r => setTimeout(r, 400));
        }

        if (!stopped) statusEl.textContent = `完成 ${total} 条`;
        state.running = false;
        updateButtons();
    }

    // ---------- 开始查询 ----------
    async function startQuery() {
        if (state.running) return;
        const fields = getSelectedFields();
        if (!fields.length) { alert('请至少勾选一个查询字段'); return; }
        if (!window.__cnipaAuth.token) {
            alert('还没有获取到登录态！\n\n请先在国知局页面的搜索框里随便搜一个申请号，等面板顶部变成"✅ 登录态已获取"后再点开始。');
            return;
        }

        // 收集申请号
        let rawList = [];
        const uploadInfo = document.getElementById('cnipa-file-info');
        const uploadActive = document.getElementById('pane-upload').style.display !== 'none';
        if (uploadActive) {
            if (!uploadInfo.dataset.appnos) { alert('请先上传 Excel 文件'); return; }
            rawList = JSON.parse(uploadInfo.dataset.appnos);
        } else {
            const input = document.getElementById('cnipa-input').innerText.trim();
            if (!input) { alert('请先粘贴申请号'); return; }
            rawList = input.split('\n').map(s => s.trim()).filter(Boolean);
        }
        if (!rawList.length) { alert('没有有效的申请号'); return; }

        // 构建行
        state.rows = rawList.map(orig => {
            const cleaned = normalizeAppNo(orig);
            return {
                original: orig,
                cleaned: cleaned,
                results: {},
                status: cleaned ? 'pending' : 'skip',
                note: cleaned ? '' : '申请号格式无法识别'
            };
        });

        // 只处理状态为 pending 的（skip 的格式错误行不查）
        const toProcess = state.rows.filter(r => r.status === 'pending');
        await runBatch(toProcess, fields);
    }

    // ---------- 重查失败项 ----------
    async function retryFailed() {
        if (state.running) return;
        const fields = getSelectedFields();
        if (!fields.length) { alert('请至少勾选一个查询字段'); return; }
        if (!window.__cnipaAuth.token) {
            alert('还没有获取到登录态！\n\n请先在页面上手动搜索一次，等面板顶部变成"✅ 登录态已获取"后再点。');
            return;
        }
        const failedRows = state.rows.filter(r => r.status === 'fail');
        if (!failedRows.length) { alert('没有失败项需要重查'); return; }
        await runBatch(failedRows, fields);
    }

    if (document.readyState === 'loading') {
        document.addEventListener('DOMContentLoaded', init);
    } else {
        init();
    }
})();
