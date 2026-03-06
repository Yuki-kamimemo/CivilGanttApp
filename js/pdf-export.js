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

    // 2. 印刷設定を適用してレンダリング
    state.displayStart = settings.printStart;
    state.displayEnd = settings.printEnd;
    state.zoomRatio = settings.zoom;
    state.viewRange = 'custom'; // 印刷時は「全体表示」や指定範囲で固定
    renderAll();

    // スクロール位置を先頭に戻してからHTMLを取得
    const rc = document.getElementById('right-container');
    const dnr = document.getElementById('daily-notes-right');
    if (rc) rc.scrollLeft = 0;
    if (dnr) dnr.scrollLeft = 0;

    // UI安定化のため少し待機
    await new Promise(r => setTimeout(r, 100));

    // ★ 印刷高さ圧縮の計算（DOM測定はUI安定後に実施）
    const MM_TO_PX = 96 / 25.4;
    const isA3 = settings.paperSize === 'a3-landscape';
    const paperHeightMm = isA3 ? 297 : 210;
    const printableHeightPx = (paperHeightMm - 20) * MM_TO_PX; // 20mm = 上下余白合計
    const PROJ_HEADER_PX = 70;   // プロジェクト情報行
    const STAMP_PX = settings.showStamp ? 90 : 0;  // ハンコ枠
    // 日別備考の高さは作業画面の実際の高さを使用（固定値ではなくDOMから取得）
    const dailyNotesEl = document.querySelector('.daily-notes-wrapper');
    const DAILY_NOTES_PRINT_PX = settings.showDaily && dailyNotesEl
        ? dailyNotesEl.offsetHeight
        : 0;
    const BUFFER_PX = 20;        // 余白バッファ
    const availForGantt = printableHeightPx - PROJ_HEADER_PX - STAMP_PX -
        (settings.showDaily ? DAILY_NOTES_PRINT_PX : 0) - BUFFER_PX;

    // カレンダーヘッダー + 全チャート行の実際の高さを測定
    const calHeaderEl = document.getElementById('calendar-header');
    const calHeaderH = calHeaderEl ? calHeaderEl.offsetHeight : 32;
    const chartAreaEl = document.getElementById('chart-area');
    const actualGanttHeight = calHeaderH + (chartAreaEl ? chartAreaEl.scrollHeight : 400);

    let ganttZoom = 1.0;
    if (actualGanttHeight > availForGantt && availForGantt > 0) {
        ganttZoom = Math.max(0.3, availForGantt / actualGanttHeight);
    }

    // 3. ハンコ枠の一時的な更新
    const stampContainer = document.querySelector('.stamp-container');
    const origStampTitles = getStampTitles();
    if (settings.showStamp) {
        setStampTitles(settings.stamp1, settings.stamp2);
    }

    // 4. クローンを作成して印刷用HTMLを組み立てる
    const mainContainerHtml = document.querySelector('.main-container').outerHTML;
    const stampHtml = settings.showStamp && stampContainer ? stampContainer.outerHTML : '';

    const projectInfoHtml = `
        <div class="print-project-header" style="display: flex; justify-content: space-between; font-size: 14px; margin-bottom: 10px; font-weight: bold;">
            <div>工事名: ${state.projectName || ''}</div>
            <div>事業者名: ${state.companyName || ''}</div>
            <div>全体工期: ${state.projectStart || ''} ～ ${state.projectEnd || ''}</div>
        </div>
    `;

    // 抽出するCSS (現在のstyle.cssを読み込むか、必要最低限のCSSをインライン化する)
    // ここでは、Playwrightがレンダリングする際に同じCSSを参照できるように
    // 抽出または絶対パスリンクを生成します。
    // 今回は簡易的にローカルのスタイルシートの内容を取得して注入します (もしfetchが使えれば)
    let styleText = '';
    try {
        const styleSheetResponse = await fetch('style.css');
        if (styleSheetResponse.ok) {
            styleText = await styleSheetResponse.text();
        }
    } catch (e) {
        console.warn('style.cssのフェッチに失敗しました (iframe環境でのみ動作が変わる可能性があります): ', e);
    }

    // 印刷時専用の微調整CSSを追加
    // html2canvas用だった print-capture-mode とは異なり、本物の印刷用CSS
    const printSpecificCss = `
        @page {
            size: ${settings.paperSize === 'a3-landscape' ? 'A3' : 'A4'} landscape;
            margin: 10mm;
        }
        body {
            background-color: white !important;
            margin: 0;
            padding: 0;
            overflow: visible !important;
        }
        .main-container {
            height: auto !important;
            overflow: visible !important;
            /* flex-direction は変更しない（左右ペインの横並びを維持） */
        }
        .gantt-wrapper {
            height: auto !important;
            overflow: visible !important;
            width: max-content !important;
            zoom: ${ganttZoom.toFixed(4)} !important;
        }
        .left-pane {
            overflow: visible !important;
            height: auto !important;
        }
        .right-pane {
            overflow: visible !important;
            height: auto !important;
            width: max-content !important;
        }
        #right-container {
            overflow: visible !important;
            height: auto !important;
            width: max-content !important;
        }
        #calendar-header {
            overflow: visible !important;
            width: max-content !important;
        }
        #chart-area {
            overflow: visible !important;
            height: auto !important;
        }
        .daily-notes-wrapper {
            display: ${settings.showDaily ? 'flex' : 'none'} !important;
            height: ${settings.showDaily ? DAILY_NOTES_PRINT_PX + 'px' : '0'} !important;
            max-height: ${settings.showDaily ? DAILY_NOTES_PRINT_PX + 'px' : '0'} !important;
            min-height: 0 !important;
            flex: none !important;
            overflow: hidden !important;
        }
        #daily-notes-right {
            overflow: hidden !important;
            height: 100% !important;
            width: max-content !important;
        }
        /* 備考エリアの表示切り替え */
        .notes-pane { display: ${settings.showNotes ? 'flex' : 'none'} !important; height: auto !important; }

        .print-hide { display: none !important; }
    `;

    const htmlString = `
<!DOCTYPE html>
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
    <div style="padding: 10px;">
        ${projectInfoHtml}
        <div style="display:flex; justify-content:flex-end; margin-bottom: 5px;">
            ${stampHtml}
        </div>
        ${mainContainerHtml}
    </div>
</body>
</html>`;

    // 5. 退避した状態を復元
    restoreStampTitles(origStampTitles);
    state.displayStart = prePdfState.displayStart;
    state.displayEnd = prePdfState.displayEnd;
    state.zoomRatio = prePdfState.zoomRatio;
    state.viewRange = prePdfState.viewRange;
    renderAll();

    return htmlString;
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
window.navigatePdfPage = function (delta) {
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
