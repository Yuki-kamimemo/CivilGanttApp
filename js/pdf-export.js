// ---------------------------------------------------
// PDF / PNG エクスポートロジック (ヘッドレスブラウザ版)
// Python (Playwright) にHTMLを送信してPDFを生成する
// ---------------------------------------------------

// ---------------------------------------------------
// 設定値収集
// ---------------------------------------------------
function collectPdfSettings() {
    return {
        printStart: document.getElementById('modal-print-start').value,
        printEnd: document.getElementById('modal-print-end').value,
        zoom: parseFloat(document.getElementById('modal-print-zoom').value) || 1.0,
        paperSize: document.getElementById('modal-print-paper').value || 'a4-landscape',
        showNotes: document.getElementById('modal-print-show-notes').checked,
        showDaily: document.getElementById('modal-print-show-daily').checked,
        showStamp: document.getElementById('modal-print-show-stamp').checked,
        stamp1: document.getElementById('modal-print-stamp1').value || '現場代理人',
        stamp2: document.getElementById('modal-print-stamp2').value || '監理技術者'
    };
}

// ---------------------------------------------------
// ローディング表示制御
// ---------------------------------------------------
function showPdfLoading(text) {
    const overlay = document.getElementById('pdf-loading-overlay');
    const textEl = overlay.querySelector('.pdf-loading-text');
    if (textEl) textEl.textContent = text || '処理中...';
    overlay.classList.add('visible');
}

function hidePdfLoading() {
    document.getElementById('pdf-loading-overlay').classList.remove('visible');
}

// ---------------------------------------------------
// ハンコ枠の操作
// ---------------------------------------------------
function getStampTitles() {
    const boxes = document.querySelectorAll('.stamp-title');
    return Array.from(boxes).map(el => el.textContent);
}

function setStampTitles(title1, title2) {
    const boxes = document.querySelectorAll('.stamp-title');
    if (boxes[0]) boxes[0].textContent = title1;
    if (boxes[1]) boxes[1].textContent = title2;
}

function restoreStampTitles(originals) {
    const boxes = document.querySelectorAll('.stamp-title');
    originals.forEach((text, i) => { if (boxes[i]) boxes[i].textContent = text; });
}

// ---------------------------------------------------
// 印刷用HTML構築
// ---------------------------------------------------
async function buildPrintableHtml(settings) {
    // 1. 現在の表示状態を退避
    const prePdfState = {
        displayStart: state.displayStart,
        displayEnd: state.displayEnd,
        zoomRatio: state.zoomRatio,
        viewRange: state.viewRange
    };

    // 2. 印刷設定を適用して一度レンダリング（高さ測定用）
    state.displayStart = settings.printStart;
    state.displayEnd = settings.printEnd;
    state.zoomRatio = settings.zoom;
    state.viewRange = 'custom';
    renderAll();

    const rc = document.getElementById('right-container');
    const dnr = document.getElementById('daily-notes-right');
    if (rc) rc.scrollLeft = 0;
    if (dnr) dnr.scrollLeft = 0;

    await new Promise(r => setTimeout(r, 100));

    // 3. ★ ページサイズ・高さ・ganttZoom の計算
    const MM_TO_PX = 96 / 25.4;
    const isA3 = settings.paperSize === 'a3-landscape';
    const paperWidthMm  = isA3 ? 420 : 297;
    const paperHeightMm = isA3 ? 297 : 210;
    const printableWidthPx  = (paperWidthMm  - 20) * MM_TO_PX; // 左右余白各10mm
    const printableHeightPx = (paperHeightMm - 20) * MM_TO_PX; // 上下余白各10mm

    const PROJ_HEADER_PX = 70;
    const STAMP_PX       = settings.showStamp ? 90 : 0;
    const BUFFER_PX      = 20;

    const dailyNotesEl = document.querySelector('.daily-notes-wrapper');
    const DAILY_NOTES_PRINT_PX = settings.showDaily && dailyNotesEl
        ? dailyNotesEl.offsetHeight : 0;

    const calHeaderEl = document.getElementById('calendar-header');
    const calHeaderH  = calHeaderEl ? calHeaderEl.offsetHeight : 32;
    const chartAreaEl = document.getElementById('chart-area');
    const actualGanttHeight = calHeaderH + (chartAreaEl ? chartAreaEl.scrollHeight : 400);

    // gantt-wrapper と daily-notes-wrapper の両方に同じ ganttZoom を適用するため
    // 合計高さに対して利用可能スペースの比率を計算
    const availSpace       = printableHeightPx - PROJ_HEADER_PX - STAMP_PX - BUFFER_PX;
    const totalZoomedHeight = actualGanttHeight + (settings.showDaily ? DAILY_NOTES_PRINT_PX : 0);

    let ganttZoom = 1.0;
    if (totalZoomedHeight > availSpace && availSpace > 0) {
        ganttZoom = Math.max(0.3, availSpace / totalZoomedHeight);
    }

    // 4. ★ 横方向ページ分割の計算
    // 左パネル実幅を取得（ganttZoom適用後の視覚幅 = leftPaneWidth * ganttZoom）
    const leftPaneEl   = document.querySelector('.left-pane');
    const leftPaneWidth = leftPaneEl ? leftPaneEl.offsetWidth : 620;

    // 右パネルの印刷可能幅（ganttZoom後）
    const rightPaneAvailPx = printableWidthPx - leftPaneWidth * ganttZoom;

    // 現在のセル幅（ganttZoom前） — renderAll()内で設定済み
    const rawCellWidth = state.viewScale === 'day' ? CELL_WIDTH_DAY : CELL_WIDTH_MONTH;

    // 1ページに収まるセル数
    const cellsPerPage = Math.max(1, Math.floor(rightPaneAvailPx / (rawCellWidth * ganttZoom)));

    // 日付範囲を各ページに分割
    const pages = [];
    if (state.viewScale === 'day') {
        let cur = new Date(settings.printStart);
        const end = new Date(settings.printEnd);
        while (cur <= end) {
            const pageEnd = new Date(cur);
            pageEnd.setDate(pageEnd.getDate() + cellsPerPage - 1);
            if (pageEnd > end) pageEnd.setTime(end.getTime());
            pages.push({ start: formatDate(cur), end: formatDate(pageEnd) });
            cur.setDate(cur.getDate() + cellsPerPage);
        }
    } else {
        // 月単位：開始月から終了月まで
        let curMonthStart = new Date(
            new Date(settings.printStart).getFullYear(),
            new Date(settings.printStart).getMonth(), 1
        );
        const endMonthStart = new Date(
            new Date(settings.printEnd).getFullYear(),
            new Date(settings.printEnd).getMonth(), 1
        );
        while (curMonthStart <= endMonthStart) {
            // ページ末の月の初日
            const pageEndMonthFirst = new Date(
                curMonthStart.getFullYear(),
                curMonthStart.getMonth() + cellsPerPage - 1, 1
            );
            const actualEndMonthFirst = pageEndMonthFirst <= endMonthStart
                ? pageEndMonthFirst : endMonthStart;
            // その月の最終日
            const pageEndDate = new Date(
                actualEndMonthFirst.getFullYear(),
                actualEndMonthFirst.getMonth() + 1, 0
            );
            pages.push({ start: formatDate(curMonthStart), end: formatDate(pageEndDate) });
            curMonthStart = new Date(
                curMonthStart.getFullYear(),
                curMonthStart.getMonth() + cellsPerPage, 1
            );
        }
    }

    // 5. スタイルシートの読み込み
    let styleText = '';
    try {
        const styleSheetResponse = await fetch('style.css');
        if (styleSheetResponse.ok) styleText = await styleSheetResponse.text();
    } catch (e) {
        console.warn('style.cssのフェッチに失敗:', e);
    }

    // 6. ハンコ枠の一時的な更新
    const stampContainer = document.querySelector('.stamp-container');
    const origStampTitles = getStampTitles();
    if (settings.showStamp) setStampTitles(settings.stamp1, settings.stamp2);

    const stampHtml = settings.showStamp && stampContainer ? stampContainer.outerHTML : '';
    const projectInfoHtml = `
        <div class="print-project-header" style="display:flex; align-items:stretch; margin-bottom:8px; border:1.5px solid #adb5bd; border-radius:4px; overflow:hidden; font-size:13px; line-height:1.3;">
            <div style="display:flex; align-items:center; background:#e9ecef; padding:5px 10px; font-weight:bold; color:#495057; white-space:nowrap; border-right:1px solid #adb5bd;">工事名</div>
            <div style="display:flex; align-items:center; padding:5px 12px; font-weight:bold; flex:1; border-right:1.5px solid #adb5bd;">${state.projectName || ''}</div>
            <div style="display:flex; align-items:center; background:#e9ecef; padding:5px 10px; font-weight:bold; color:#495057; white-space:nowrap; border-right:1px solid #adb5bd;">事業者名</div>
            <div style="display:flex; align-items:center; padding:5px 12px; font-weight:bold; border-right:1.5px solid #adb5bd;">${state.companyName || ''}</div>
            <div style="display:flex; align-items:center; background:#e9ecef; padding:5px 10px; font-weight:bold; color:#495057; white-space:nowrap; border-right:1px solid #adb5bd;">全体工期</div>
            <div style="display:flex; align-items:center; padding:5px 12px; font-weight:bold; white-space:nowrap;">${state.projectStart || ''} ～ ${state.projectEnd || ''}</div>
        </div>`;

    // 7. ★ 各ページのHTMLを生成して縦に積み重ねる
    // render.js の renderAll() が「CELL_WIDTH_DAY = 45 * state.zoomRatio」でセル幅をリセットするため、
    // zoomRatio を調整することでカレンダー幅を制御する
    const origZoomRatio = state.zoomRatio; // = settings.zoom
    const BASE_CELL_WIDTH = state.viewScale === 'day' ? 45 : 100; // render.js の基本セル幅

    let allPagesHtml = '';
    for (let i = 0; i < pages.length; i++) {
        const pageInfo  = pages[i];
        const isLastPage = i === pages.length - 1;

        // このページの実際のセル数（日数 or 月数）を計算
        let actualCells;
        if (state.viewScale === 'day') {
            const startD = new Date(pageInfo.start);
            const endD   = new Date(pageInfo.end);
            actualCells  = Math.round((endD - startD) / 86400000) + 1;
        } else {
            const startD = new Date(pageInfo.start);
            const endD   = new Date(pageInfo.end);
            actualCells  = (endD.getFullYear() * 12 + endD.getMonth())
                         - (startD.getFullYear() * 12 + startD.getMonth()) + 1;
        }

        // カレンダーが印刷可能幅より狭い場合は zoomRatio を上げて伸張する
        // （renderAll が CELL_WIDTH = BASE * zoomRatio でセル幅を決定するため）
        state.zoomRatio = origZoomRatio;
        if (actualCells > 0 && rightPaneAvailPx > 0) {
            const actualZoomedWidth = actualCells * rawCellWidth * ganttZoom;
            if (actualZoomedWidth < rightPaneAvailPx) {
                // 伸張後の zoomRatio: BASE * zoomRatio * ganttZoom * actualCells = rightPaneAvailPx
                state.zoomRatio = rightPaneAvailPx / (actualCells * BASE_CELL_WIDTH * ganttZoom);
            }
        }

        state.displayStart = pageInfo.start;
        state.displayEnd   = pageInfo.end;
        renderAll();
        if (rc) rc.scrollLeft = 0;
        if (dnr) dnr.scrollLeft = 0;
        await new Promise(r => setTimeout(r, 80));

        // ★追加：エディタ本体が備考欄（または日報セル）に居座っていると
        // PDFにエディタUIが混入するため、キャプチャ直前だけ一時的に切り離す
        const em = window.editorManager;
        let originalParent = null;
        let originalHtml = '';
        if (em) {
            // 現在のHTMLを保持してエディタを抜く
            if (em.activeContainer) {
                originalParent = em.activeContainer;
                originalHtml = em.quill.root.innerHTML;
                em.detachEditor();
                originalParent.innerHTML = `<div class="ql-editor ql-editor-content" style="padding:0;">${originalHtml}</div>`;
            }
        }

        // ★追加：備考欄の高さチェックと自動圧縮（A3横サイズへの適合）
        const notesContainer = document.getElementById('project-notes-editor-container');
        if (notesContainer && settings.showNotes) {
            const LIMIT_HEIGHT = 880; // A3横の安全な高さ（px）
            const contentHeight = notesContainer.scrollHeight;
            
            // もし内容が制限を超えていたら、縮小して収める
            if (contentHeight > LIMIT_HEIGHT) {
                const scale = (LIMIT_HEIGHT / contentHeight).toFixed(3);
                notesContainer.style.transform = `scale(${scale})`;
                notesContainer.style.transformOrigin = 'top center';
                // 縮小後の描画崩れを防ぐため、元のコンテナの高さは維持する
                notesContainer.style.height = `${contentHeight}px`; 
            } else {
                notesContainer.style.transform = '';
                notesContainer.style.height = '';
            }
        }

        // ★スペーサーをPDFキャプチャから除外（水平スクロールバー高さ分の余白が混入するのを防ぐ）
        const spacer = document.getElementById('left-hscroll-spacer');
        if (spacer) spacer.style.display = 'none';

        const pageMainHtml = document.querySelector('.main-container').outerHTML;
        const pageBreak = isLastPage ? '' : 'page-break-after: always;';

        // ★スペーサーを元に戻す
        if (spacer) spacer.style.display = '';

        allPagesHtml += `
<div style="padding:10px; ${pageBreak}">
    ${projectInfoHtml}
    <div style="display:flex; justify-content:flex-end; margin-bottom:5px;">
        ${stampHtml}
    </div>
    ${pageMainHtml}
</div>`;

        // ★追加：キャプチャが終わったのでエディタを元の位置に戻す
        if (em && originalParent) {
            if (originalParent === em.defaultContainer) {
                em.dockToDefault();
            } else {
                // インライン編集中の場合はその場所に戻す
                em.openEditor(originalParent, em.onSaveCallback, originalParent.style.writingMode === 'vertical-rl');
            }
        }
    }

    // 8. 退避した状態を復元
    restoreStampTitles(origStampTitles);
    state.displayStart = prePdfState.displayStart;
    state.displayEnd   = prePdfState.displayEnd;
    state.zoomRatio    = prePdfState.zoomRatio;
    state.viewRange    = prePdfState.viewRange;
    renderAll();

    // 9. 印刷用CSS
    const printSpecificCss = `
        @page {
            size: ${isA3 ? 'A3' : 'A4'} landscape;
            margin: 10mm;
        }
        body {
            background-color: white !important;
            margin: 0; padding: 0;
            overflow: visible !important;
        }
        .main-container {
            height: auto !important;
            overflow: visible !important;
            border: 1px solid #adb5bd !important;
            border-radius: 4px !important;
        }
        .gantt-wrapper {
            height: auto !important;
            overflow: visible !important;
            width: max-content !important;
            display: flex !important;
            align-items: flex-start !important; /* 左ペインが右側に引っ張られて伸びるのを防ぐ */
            zoom: ${ganttZoom.toFixed(4)} !important;
        }
        .left-pane {
            overflow: visible !important;
            height: auto !important;
            min-height: 0 !important;
            flex: none !important;
            resize: none !important; /* リサイズハンドルを消す */
            border-bottom: 1px solid #dee2e6 !important; /* 下端の線を確実に描画 */
        }
        /* table-container の overflow:auto がPDF時に行ずれを起こすため無効化 */
        .table-container, #left-container {
            overflow: visible !important;
            height: auto !important;
            min-height: 0 !important;
            flex: none !important; /* flexによる自動拡張を無効化 */
        }
        /* スクロールバー補正スペーサーをPDFから除外 */
        #left-hscroll-spacer {
            display: none !important;
        }
        /* テーブル行の高さがコンテンツによって伸びるのを防ぎ、同期を安定させる */
        .task-row {
            break-inside: avoid !important;
        }
        .task-row td {
            overflow: hidden !important;
            box-sizing: border-box !important;
            line-height: 1.2 !important;
        }
        /* th の position:sticky がPDFレンダリング時にレイアウトズレを起こすため無効化 */
        th {
            position: static !important;
        }
        .right-pane {
            overflow: visible !important;
            height: auto !important;
            min-height: 0 !important; /* 空白を詰めるために必須 */
            width: max-content !important;
        }
        #right-container {
            overflow: visible !important;
            height: auto !important;
            min-height: 0 !important; /* 空白を詰めるために必須 */
            width: max-content !important;
        }
        #calendar-header {
            overflow: visible !important;
            width: max-content !important;
        }
        #chart-area {
            overflow: visible !important;
            height: auto !important;
            min-height: 0 !important; /* 空白を詰めるために必須 */
        }
        .daily-notes-wrapper {
            display: ${settings.showDaily ? 'flex' : 'none'} !important;
            height: ${settings.showDaily ? DAILY_NOTES_PRINT_PX + 'px' : '0'} !important;
            max-height: ${settings.showDaily ? DAILY_NOTES_PRINT_PX + 'px' : '0'} !important;
            min-height: 0 !important;
            flex: none !important;
            overflow: hidden !important;
            zoom: ${ganttZoom.toFixed(4)} !important;
        }
        #daily-notes-right {
            overflow: hidden !important;
            height: 100% !important;
            width: max-content !important;
        }
        /* タブ名: ganttZoom で縮小される分を逆補正して常に読みやすいサイズにする */
        #daily-notes-left {
            width: ${leftPaneWidth}px !important;
            display: flex !important;
            align-items: center !important;
            justify-content: center !important;
            overflow: visible !important;
            flex-shrink: 0 !important;
        }
        .daily-notes-content {
            overflow: visible !important;
        }
        .daily-tab-actions {
            display: flex !important;
            align-items: center !important;
            gap: ${(8 / ganttZoom).toFixed(1)}px !important;
            padding: ${(8 / ganttZoom).toFixed(1)}px ${(16 / ganttZoom).toFixed(1)}px !important;
            border-radius: ${(6 / ganttZoom).toFixed(1)}px !important;
            border: ${(1 / ganttZoom).toFixed(1)}px solid #ced4da !important;
            background: white !important;
        }
        .daily-tab-actions button {
            display: none !important;
        }
        #current-daily-tab-name {
            font-size: ${(13 / ganttZoom).toFixed(1)}px !important;
            font-weight: bold !important;
            white-space: nowrap !important;
            color: #212529 !important;
            display: block !important;
        }
        .notes-pane { display: ${settings.showNotes ? 'flex' : 'none'} !important; height: auto !important; overflow: hidden !important; }
        #project-notes-editor-container { overflow: hidden !important; }
        #project-notes-editor-container p,
        #project-notes-editor-container h1,
        #project-notes-editor-container h2,
        #project-notes-editor-container h3 { margin: 0; padding: 0; }
        .print-hide { display: none !important; }
    `;

    return `<!DOCTYPE html>
<html lang="ja">
<head>
    <meta charset="UTF-8">
    <title>Print Preview</title>
    <style>
        ${styleText}
        ${printSpecificCss}
    </style>
</head>
<body class="print-capture-mode">
    ${allPagesHtml}
</body>
</html>`;
}

// ---------------------------------------------------
// プレビュー表示
// ---------------------------------------------------
window.updatePdfPreview = async function () {
    showPdfLoading('プレビュー生成中...');
    try {
        const settings = collectPdfSettings();
        const htmlString = await buildPrintableHtml(settings);

        let iframe = document.getElementById('pdf-preview-iframe');
        if (!iframe) {
            console.warn('プレビュー用の iframe が見つかりません。動的に生成します。');
            const previewArea = document.getElementById('pdf-preview-area');
            if (previewArea) {
                previewArea.innerHTML = '<iframe id="pdf-preview-iframe" style="flex:1; width:100%; border:1px solid #ccc; background: white;" title="PDF Preview"></iframe>';
                iframe = document.getElementById('pdf-preview-iframe');
            }
        }

        if (iframe) {
            iframe.srcdoc = htmlString;
        } else {
            console.error('プレビューエリア自体が見つかりません。');
        }

    } catch (e) {
        console.error('プレビュー生成エラー:', e);
        alert('プレビューの生成に失敗しました:\n' + e.message);
    } finally {
        hidePdfLoading();
    }
};

// ヘッドレスブラウザ方式ではページングの概念が「プレビュー上のスクロール」に変わるためダミー化
window.navigatePdfPage = function (_delta) {
    console.log('ページナビゲーション機能は iframe プレビューでは不要(スクロールで確認)');
};

// ---------------------------------------------------
// ファイル名生成
// ---------------------------------------------------
function generateDefaultFilename(ext) {
    const name = (state.projectName || '工程表').replace(/[\\/:*?"<>|]/g, '_');
    const now = new Date();
    const ymd = `${now.getFullYear()}${String(now.getMonth() + 1).padStart(2, '0')}${String(now.getDate()).padStart(2, '0')}`;
    return `${name}_${ymd}.${ext}`;
}

// ---------------------------------------------------
// PDF保存 (Python APIへの送信)
// ---------------------------------------------------
window.executePdfSave = async function () {
    showPdfLoading('PDF出力エンジンへ送信中...');
    try {
        const settings = collectPdfSettings();
        const htmlString = await buildPrintableHtml(settings);
        const defaultName = generateDefaultFilename('pdf');

        if (window.pywebview && window.pywebview.api && window.pywebview.api.save_pdf_file) {
            // HTML文字列と設定（余白や用紙サイズ）を直接Pythonバックエンドに渡す
            const result = await window.pywebview.api.save_pdf_file(htmlString, settings, defaultName);
            if (result) {
                if (window.showToast) window.showToast('PDFを保存しました');
            }
        } else {
            alert('PDF出力バックエンド(Playwright等)に接続できません。アプリから実行してください。');
        }
    } catch (e) {
        console.error('PDF保存エラー:', e);
        alert('PDFの保存に失敗しました:\n' + e.message);
    } finally {
        hidePdfLoading();
    }
};

// ---------------------------------------------------
// PNG保存 (Python APIへの送信)
// ---------------------------------------------------
window.executePngSave = async function () {
    showPdfLoading('PNG出力エンジンへ送信中...');
    try {
        const settings = collectPdfSettings();
        const htmlString = await buildPrintableHtml(settings);
        const defaultName = generateDefaultFilename('png');

        if (window.pywebview && window.pywebview.api && window.pywebview.api.save_png_file) {
            const result = await window.pywebview.api.save_png_file(htmlString, defaultName);
            if (result) {
                if (window.showToast) window.showToast('PNGを保存しました');
            }
        } else {
            alert('PNG出力バックエンドに接続できません。アプリから実行してください。');
        }
    } catch (e) {
        console.error('PNG保存エラー:', e);
        alert('PNGの保存に失敗しました:\n' + e.message);
    } finally {
        hidePdfLoading();
    }
};
