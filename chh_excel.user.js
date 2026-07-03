// ==UserScript==
// @name         CHH摸鱼助手(excel)
// @namespace    http://tampermonkey.net/
// @version      1.0
// @description  把 Chiphell 论坛伪装成 Excel 表格。公式栏变成可点击的层级目录，底部 Sheet 实现翻页。
// @author       Gemini (ported from S1摸鱼助手 by Gemini)
// @match        *://www.chiphell.com/forum-*
// @match        *://www.chiphell.com/thread-*
// @match        *://www.chiphell.com/forum.php
// @grant        GM_addStyle
// @run-at       document-end
// @license MIT

// ==/UserScript==

(function() {
    'use strict';

    let isWorkMode = localStorage.getItem('chh_excel_mode') === 'true';

    // ==========================================
    // 配色 (Excel 经典 Office Green — 摸鱼伪装的灵魂色,
    //       不要适配论坛, 越像真 Excel 越摸鱼)
    // ==========================================
    const COLOR_PRIMARY  = '#217346';   // Excel 经典绿
    const COLOR_PRIMARY_DARK = '#104e2b'; // 深绿 (hover)
    const COLOR_HOVER_BG = '#e6f2ea';    // 浅绿 hover 背景
    const COLOR_TEXT     = '#000';
    const COLOR_BORDER   = '#e1e1e1';

    // ==========================================
    // 1. Excel 界面 HTML
    // ==========================================
    const excelHeaderHTML = `
        <div id="excel-fake-header">
            <div class="excel-title-bar">
                <span class="excel-icon">X</span>
                <span class="excel-filename">Hardware_Discussion_2026.xlsx - Excel</span>
            </div>
            <div class="excel-menu-bar">
                <span>文件</span>
                <span class="active">开始</span>
                <span>插入</span>
                <span>页面布局</span>
                <span>公式</span>
                <span>数据</span>
                <span>审阅</span>
                <span>视图</span>
            </div>
            <div class="excel-ribbon">
                <div class="ribbon-group">
                    <div class="ribbon-btn">📋 粘贴</div>
                </div>
                <div class="ribbon-divider"></div>
                <div class="ribbon-group font-settings">
                    <select><option>DengXian</option><option>Calibri</option></select>
                    <select><option>11</option></select>
                    <div class="font-actions"><b>B</b> <i>I</i> <u>U</u></div>
                </div>
                <div class="ribbon-divider"></div>
                <div class="ribbon-group" style="color:#999; font-size:12px; margin-left:10px;">
                    <span style="margin-right:10px">自动换行</span>
                    <span>合并后居中</span>
                </div>
            </div>
            <div class="excel-formula-bar">
                <span class="name-box">A1</span>
                <span class="fx">fx</span>
                <div id="formula-bar-content"></div>
            </div>
            <div class="excel-col-headers">
                <div class="row-idx-blank"></div>
                <div style="flex:1; min-width:120px; border-right:1px solid #d4d4d4;">A (用户/ID)</div>
                <div style="flex:8; border-right:1px solid #d4d4d4;">B (主题 / 内容)</div>
            </div>
        </div>
        <div id="excel-sheet-bar-container">
            <div class="sheet-nav-arrows">
                <span>◀</span><span>▶</span>
            </div>
            <div id="excel-sheets-scroll">
                </div>
            <div class="sheet-add-btn">+</div>
        </div>
        <div id="excel-footer-bar">
            <span>就绪</span>
            <span style="float:right; margin-right:20px;">-------+ -- 100%</span>
        </div>
    `;

    const headerDiv = document.createElement('div');
    headerDiv.innerHTML = excelHeaderHTML;
    document.body.prepend(headerDiv);

    // ==========================================
    // 2. CSS 样式 (配色替换为 Chiphell 暗红系)
    // ==========================================
    const css = `
        /* 切换按钮 (右上角 — 摸鱼模式开关, 留位给图按钮) */
        #excel-toggle-btn {
            position: fixed; top: 5px; right: 85px; padding: 5px 10px;
            background: ${COLOR_PRIMARY}; color: white; opacity: 1; z-index: 999999;
            cursor: pointer; font-size: 12px; font-weight: bold; border-radius: 4px;
            box-shadow: 0 2px 5px rgba(0,0,0,0.2); font-family: sans-serif;
        }
        #excel-toggle-btn:hover { background: ${COLOR_PRIMARY_DARK}; }

        /* 全局图片切换按钮 (右上角) */
        #excel-img-toggle-btn {
            position: fixed; top: 5px; right: 5px; padding: 5px 10px;
            background: #444; color: white; opacity: 1; z-index: 999999;
            cursor: pointer; font-size: 12px; font-weight: bold; border-radius: 4px;
            box-shadow: 0 2px 5px rgba(0,0,0,0.2); font-family: sans-serif;
        }
        #excel-img-toggle-btn:hover { background: #222; }
        #excel-img-toggle-btn.all-shown { background: ${COLOR_PRIMARY}; }
        #excel-img-toggle-btn.all-shown:hover { background: ${COLOR_PRIMARY_DARK}; }

        /* === 图片隐藏 (移植自 NGA 摸鱼体验) === */
        body.excel-mode .t_f img.hld__img-hidden { display: none !important; }
        .excel-img-toggle {
            display: inline-block; margin: 2px 0; padding: 2px 10px;
            background: #f8f8f8; border: 1px solid #ccc; color: #444;
            font-size: 10pt; border-radius: 2px; cursor: pointer;
            font-family: 'DengXian', sans-serif; line-height: 1.4;
        }
        .excel-img-toggle:hover {
            background: ${COLOR_HOVER_BG}; border-color: ${COLOR_PRIMARY}; color: ${COLOR_PRIMARY};
        }

        /* === 全局重置 === */
        body.excel-mode {
            background: #fff !important;
            font-family: 'DengXian', 'Calibri', sans-serif !important;
            font-size: 11pt !important;
            color: ${COLOR_TEXT} !important;
            margin: 0 !important;
            padding-top: 165px !important;
            padding-bottom: 60px !important;
            overflow-x: hidden !important;
        }

        /* === 隐藏 Chiphell 界面元素 === */
        body.excel-mode #toptb,
        body.excel-mode #hd,
        body.excel-mode #nv_ph,
        body.excel-mode #scbar,
        body.excel-mode #ft,
        body.excel-mode #scrolltop,
        body.excel-mode #pt,
        body.excel-mode .bm.bml,
        body.excel-mode .bm.bmw.fl,
        body.excel-mode .pgbtn,
        body.excel-mode .focus,
        body.excel-mode .a_mu, body.excel-mode .ad,
        body.excel-mode #f_pst,
        body.excel-mode .th,
        body.excel-mode tbody[id^="separatorline"],
        body.excel-mode tbody[id^="stickthread"],
        body.excel-mode #pgt,
        body.excel-mode .pgs,
        body.excel-mode #fd_page_bottom
        { display: none !important; }

        /* === 列表页样式 === */
        body.excel-mode .wp, body.excel-mode #ct, body.excel-mode .mn, body.excel-mode .tl, body.excel-mode .bm_c {
            width: 100% !important; margin: 0 !important; padding: 0 !important; border: none !important; min-width: 0 !important;
        }
        body.excel-mode #threadlisttableid {
            width: 100% !important; border-collapse: collapse !important; table-layout: fixed !important; margin-top: -1px !important;
        }
        body.excel-mode #threadlisttableid tbody { display: table-row-group !important; border: none !important; }
        body.excel-mode #threadlisttableid tr { display: table-row !important; height: 22px !important; }
        body.excel-mode #threadlisttableid td, body.excel-mode #threadlisttableid th {
            padding: 2px 5px !important; border: 1px solid ${COLOR_BORDER} !important; height: 22px !important; line-height: 22px !important;
            font-size: 11pt !important; font-weight: normal !important; background: #fff !important; color: ${COLOR_TEXT} !important;
            white-space: nowrap !important; overflow: hidden !important; text-align: left !important;
        }
        body.excel-mode #threadlisttableid tr:hover td, body.excel-mode #threadlisttableid tr:hover th {
            background: ${COLOR_HOVER_BG} !important; outline: 1px solid ${COLOR_PRIMARY}; z-index: 10;
        }
        body.excel-mode .xst, body.excel-mode a { color: ${COLOR_TEXT} !important; text-decoration: none !important; font-family: 'DengXian'; }
        body.excel-mode #threadlisttableid td.icn,
        body.excel-mode #threadlisttableid td.o,
        body.excel-mode .tps, body.excel-mode .fico-image, body.excel-mode .fico-attachment
        { display: none !important; }
        body.excel-mode th.common em, body.excel-mode th.common em a { color: #999 !important; margin-right: 5px; }

        /* === 帖子正文页样式 === */
        body.excel-mode #postlist, body.excel-mode .plhin, body.excel-mode .pls, body.excel-mode .plc, body.excel-mode .bm {
            background: #fff !important; background-color: #fff !important; background-image: none !important; border: none !important;
        }
        body.excel-mode table.plhin { width: 100% !important; border-collapse: collapse !important; margin-bottom: -1px !important; }

        /* 左侧用户栏 */
        body.excel-mode .pls {
            width: 120px !important; border: 1px solid ${COLOR_BORDER} !important; padding: 5px !important; vertical-align: top !important;
            font-size: 11pt !important; color: #333 !important; text-align: left !important;
        }
        body.excel-mode .pls .avatar, body.excel-mode .pls .tns, body.excel-mode .pls p, body.excel-mode .pls dl, body.excel-mode .pls .md_ctrl, body.excel-mode .pls .o
        { display: none !important; }
        body.excel-mode .pls .pi { padding: 0 !important; margin: 0 !important; border: none !important; text-align: left !important; overflow: hidden; height: auto !important; }
        body.excel-mode .pls .authi { font-weight: normal !important; color: ${COLOR_TEXT} !important; }

        /* 右侧内容栏 */
        body.excel-mode .plc {
            border: 1px solid ${COLOR_BORDER} !important; padding: 5px 10px !important; vertical-align: top !important; width: auto !important;
        }
        body.excel-mode .t_f {
            font-size: 11pt !important; font-family: 'DengXian', sans-serif !important; line-height: 1.4 !important; color: ${COLOR_TEXT} !important;
        }
        body.excel-mode .t_f img {
            display: block !important; max-width: 98% !important; height: auto !important; opacity: 0.8; margin: 5px 0 !important;
        }

        /* 操作按钮 (可见) */
        body.excel-mode .sign, body.excel-mode .modact, body.excel-mode .a_ga { display: none !important; }
        body.excel-mode .pob { padding: 2px 0 !important; background: #fff !important; border: none !important; }
        body.excel-mode .po { text-align: right !important; padding: 0 5px 0 0 !important; margin: 0 !important; }
        body.excel-mode .po a {
            display: inline-block !important; border: 1px solid #ccc !important; padding: 1px 5px !important; margin-left: 5px !important;
            background: #f8f8f8 !important; color: #444 !important; text-decoration: none !important; font-size: 10pt !important; border-radius: 2px;
        }
        body.excel-mode .po a:hover { background: ${COLOR_HOVER_BG} !important; border-color: ${COLOR_PRIMARY} !important; }

        /* === Excel UI 头部样式 === */
        #excel-fake-header { display: none; }
        body.excel-mode #excel-fake-header { display: block; position: fixed; top: 0; left: 0; width: 100%; z-index: 9999; background: #f3f2f1; border-bottom: 1px solid #ccc; }
        .excel-title-bar { background: ${COLOR_PRIMARY}; color: white; padding: 5px 10px; font-size: 12px; display: flex; align-items: center; }
        .excel-icon { background: white; color: ${COLOR_PRIMARY}; padding: 0 4px; margin-right: 10px; font-weight: bold; border-radius: 2px; font-size: 10px;}
        .excel-menu-bar { background: ${COLOR_PRIMARY}; color: #eee; display: flex; font-size: 13px; padding-top: 5px;}
        .excel-menu-bar span { padding: 5px 12px; cursor: pointer; }
        .excel-menu-bar span.active { background: #f3f2f1; color: ${COLOR_PRIMARY}; border-radius: 4px 4px 0 0; }
        .excel-ribbon { height: 50px; background: #f3f2f1; display: flex; align-items: center; padding: 0 10px; border-bottom: 1px solid #d4d4d4; }
        .ribbon-divider { height: 30px; width: 1px; background: #ccc; margin: 0 10px; }
        .font-settings select { height: 20px; font-size: 11px; border: 1px solid #ccc; }
        .font-actions { margin-top: 2px; font-size: 12px; }
        .font-actions * { margin-right: 5px; cursor: pointer; padding: 0 2px;}
        .excel-formula-bar { display: flex; align-items: center; padding: 4px; background: white; }
        .name-box { width: 40px; border-right: 1px solid #ccc; text-align: center; font-size: 11px; color: #333; margin-right: 8px;}
        .fx { color: #ccc; font-weight: bold; margin-right: 8px; font-style: italic; }

        /* 公式栏 DIV 样式 */
        #formula-bar-content {
            width: 100%; height: 20px; line-height: 20px; outline: none;
            font-family: 'DengXian'; font-size: 11pt; color: #333;
            overflow: hidden; white-space: nowrap;
        }
        #formula-bar-content a { color: #333; text-decoration: none; margin: 0 2px; }
        #formula-bar-content a:hover { text-decoration: underline; color: ${COLOR_PRIMARY}; }
        #formula-bar-content em { color: #999; font-style: normal; margin: 0 2px; }
        #formula-bar-content .nvhm {
            background: none !important; text-indent: 0 !important; width: auto !important;
            float: none !important; font-family: 'DengXian' !important;
        }

        .excel-col-headers { display: flex; background: #f3f2f1; border-bottom: 1px solid #d4d4d4; font-size: 11px; color: #666; text-align: center; height: 20px; line-height: 20px;}
        .excel-col-headers div { border-right: 1px solid #d4d4d4; }
        .row-idx-blank { width: 35px; background: #e6e6e6; }

        /* === 底部 Sheet 栏样式 === */
        #excel-sheet-bar-container { display: none; }
        body.excel-mode #excel-sheet-bar-container {
            display: flex; position: fixed; bottom: 22px; left: 0; width: 100%; height: 30px;
            background: #f3f2f1; border-top: 1px solid #d4d4d4; z-index: 10000; align-items: center; padding-left: 5px;
        }
        .sheet-nav-arrows span { color: #999; padding: 0 5px; cursor: pointer; font-size: 12px; }
        .sheet-nav-arrows span:hover { color: #333; }

        #excel-sheets-scroll {
            display: flex; overflow-x: auto; margin-left: 10px; scrollbar-width: none;
        }
        #excel-sheets-scroll::-webkit-scrollbar { display: none; }

        .sheet-tab {
            padding: 4px 15px; margin-right: 1px; font-size: 12px; cursor: pointer; white-space: nowrap;
            color: ${COLOR_TEXT}; text-decoration: none !important; border-right: 1px solid #d4d4d4;
        }
        .sheet-tab:hover { background: #e1e1e1; }
        .sheet-tab.active {
            background: #fff; color: ${COLOR_PRIMARY}; font-weight: bold; border-bottom: 2px solid #fff; position: relative; top: 1px;
        }
        .sheet-add-btn { margin-left: 10px; color: #666; cursor: pointer; font-size: 16px; width: 20px; text-align: center;}

        /* 底部就绪状态栏 */
        #excel-footer-bar { display: none; }
        body.excel-mode #excel-footer-bar {
            display: block; position: fixed; bottom: 0; left: 0; width: 100%; height: 22px;
            background: #f3f2f1; border-top: 1px solid #d4d4d4; color: #666; font-size: 11px;
            line-height: 22px; padding-left: 10px; z-index: 10000;
        }
    `;
    GM_addStyle(css);

    // ==========================================
    // 3. JS 逻辑
    // ==========================================
    const btn = document.createElement('div');
    btn.id = 'excel-toggle-btn';
    btn.innerHTML = isWorkMode ? '退出摸鱼' : '摸鱼模式';
    btn.title = 'Ctrl+Shift+X 切换模式';
    btn.addEventListener('click', toggleMode);
    document.body.appendChild(btn);

    // 全局图片切换按钮 (移植自 NGA)
    const imgBtn = document.createElement('div');
    imgBtn.id = 'excel-img-toggle-btn';
    imgBtn.innerHTML = '🖼️ 显示图片';
    imgBtn.title = '一键显示/隐藏当前帖子所有图片';
    imgBtn.addEventListener('click', toggleAllImages);
    document.body.appendChild(imgBtn);

    document.addEventListener('keydown', function(e) {
        if (e.ctrlKey && e.shiftKey && e.key.toUpperCase() === 'X') {
             e.preventDefault();
             toggleMode();
        }
    });

    if (isWorkMode) {
        enableExcel();
    }

    function toggleMode() {
        isWorkMode = !isWorkMode;
        localStorage.setItem('chh_excel_mode', isWorkMode);
        btn.innerHTML = isWorkMode ? '退出摸鱼' : '摸鱼模式';
        if (isWorkMode) enableExcel();
        else disableExcel();
    }

    function enableExcel() {
        document.body.classList.add('excel-mode');
        document.title = "Hardware_Discussion_2026.xlsx - Excel";

        forceLinksToSelf();
        initBreadcrumb();
        initSheetPagination();
        updateFormulaBar();
        initImageHider();
    }

    function disableExcel() {
        document.body.classList.remove('excel-mode');
        document.title = "Chiphell";
    }

    function forceLinksToSelf() {
        const links = document.querySelectorAll('a.xst, a.s, #ct a[target="_blank"]');
        links.forEach(a => {
            if (a.target === '_blank') {
                a.removeAttribute('target');
                a.target = '_self';
            }
        });
    }

    // === 核心功能1：带链接的面包屑导航 ===
    function initBreadcrumb() {
        const pt = document.getElementById('pt');
        const formulaContent = document.getElementById('formula-bar-content');

        if (pt && formulaContent) {
            const clone = pt.querySelector('.z').cloneNode(true);

            const homeLink = clone.querySelector('.nvhm');
            if(homeLink) {
                homeLink.innerText = "Chiphell";
            }

            formulaContent.innerHTML = clone.innerHTML;
            formulaContent.dataset.defaultHtml = clone.innerHTML;

            const links = formulaContent.querySelectorAll('a');
            links.forEach(a => a.target = "_self");
        }
    }

    // === 核心功能2：Sheet 翻页 ===
    function initSheetPagination() {
        const sheetContainer = document.getElementById('excel-sheets-scroll');
        sheetContainer.innerHTML = '';

        const s1Pager = document.querySelector('#fd_page_bottom .pg') || document.querySelector('.pgs .pg') || document.querySelector('.pg');

        if (s1Pager) {
            const elements = s1Pager.querySelectorAll('a, strong, label');

            elements.forEach(el => {
                if (el.tagName === 'LABEL') return;

                const sheetTab = document.createElement(el.tagName === 'A' ? 'a' : 'div');
                sheetTab.className = 'sheet-tab';

                let text = el.innerText;
                if (el.classList.contains('prev')) text = '◀ 上一页';
                if (el.classList.contains('nxt')) text = '下一页 ▶';
                if (text.includes('...')) text = '...';

                sheetTab.innerText = text;

                if (el.tagName === 'STRONG') {
                    sheetTab.classList.add('active');
                } else {
                    sheetTab.href = el.href;
                    sheetTab.target = '_self';
                }
                sheetContainer.appendChild(sheetTab);
            });
        } else {
            const sheet1 = document.createElement('div');
            sheet1.className = 'sheet-tab active';
            sheet1.innerText = 'Sheet1';
            sheetContainer.appendChild(sheet1);
        }
    }

    function updateFormulaBar() {
        const container = document.getElementById('formula-bar-content');

        document.addEventListener('mouseover', function(e) {
            if (!isWorkMode) return;
            let text = '';
            let target = e.target;

            if (target.tagName === 'A' && (target.closest('#threadlisttableid') || target.closest('.plhin'))) text = target.innerText;
            else if (target.tagName === 'TD' || target.tagName === 'TH') text = target.innerText;

            if (text && text.trim() !== "") {
                container.innerText = text.trim();
            }
        });

        document.addEventListener('mouseout', function(e) {
            if (!isWorkMode) return;

            if (!e.relatedTarget || (!e.relatedTarget.closest('#threadlisttableid') && !e.relatedTarget.closest('.plhin'))) {
                const defaultHtml = container.dataset.defaultHtml;
                if (defaultHtml) {
                    container.innerHTML = defaultHtml;
                    container.querySelectorAll('a').forEach(a => a.target = "_self");
                }
            }
        });
    }

    // ==========================================
    // 4. 图片隐藏 (移植自 NGA 摸鱼体验, 默认隐藏)
    // ==========================================
    function initImageHider() {
        // 只处理帖子正文 .t_f 里带 class="zoom" 的图片 (Discuz! 标准附件)
        // 自动跳过: 头像 (user_avatar)、UI 图标 (boardlogo 等)、表情
        const images = document.querySelectorAll('.t_f img.zoom:not([data-hld-injected])');

        images.forEach(img => {
            // 标记已注入, 防止重复 (Ctrl+Shift+X 切换时)
            img.dataset.hldInjected = '1';

            // 默认隐藏该图片
            img.classList.add('hld__img-hidden');

            // 在 img 后插入逐图切换按钮
            const btn = document.createElement('button');
            btn.className = 'excel-img-toggle';
            btn.type = 'button';
            btn.textContent = '📷 显示图片';
            btn.addEventListener('click', function() {
                const isHidden = img.classList.contains('hld__img-hidden');
                if (isHidden) {
                    img.classList.remove('hld__img-hidden');
                    btn.textContent = '🚫 隐藏图片';
                } else {
                    img.classList.add('hld__img-hidden');
                    btn.textContent = '📷 显示图片';
                }
                updateImgToggleBtn();
            });
            img.parentNode.insertBefore(btn, img.nextSibling);
        });

        // 同步全局按钮状态
        updateImgToggleBtn();
    }

    function toggleAllImages() {
        const images = document.querySelectorAll('.t_f img.hld__img-hidden');
        const allBtns = document.querySelectorAll('.excel-img-toggle');
        const isShowingAll = images.length === 0; // 当前全部显示 → 下一步全部隐藏

        if (isShowingAll) {
            // 当前显示 → 全部隐藏
            document.querySelectorAll('.t_f img.zoom').forEach(img => {
                img.classList.add('hld__img-hidden');
            });
            allBtns.forEach(b => { b.textContent = '📷 显示图片'; });
        } else {
            // 当前隐藏 → 全部显示
            document.querySelectorAll('.t_f img.zoom').forEach(img => {
                img.classList.remove('hld__img-hidden');
            });
            allBtns.forEach(b => { b.textContent = '🚫 隐藏图片'; });
        }
        updateImgToggleBtn();
    }

    function updateImgToggleBtn() {
        const btn = document.getElementById('excel-img-toggle-btn');
        if (!btn) return;
        const hiddenCount = document.querySelectorAll('.t_f img.hld__img-hidden').length;
        const totalCount = document.querySelectorAll('.t_f img.zoom').length;

        if (hiddenCount === 0 && totalCount > 0) {
            // 全部显示 → 按钮提示 "全部隐藏"
            btn.innerHTML = '🖼️ 全部隐藏';
            btn.classList.add('all-shown');
        } else {
            // 有隐藏 → 按钮提示 "全部显示"
            btn.innerHTML = '🖼️ 全部显示';
            btn.classList.remove('all-shown');
        }
    }
})();