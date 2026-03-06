// ---------------------------------------------------
// PDF / PNG エクスポートロジック
// html2canvas + jsPDF を使用
// ---------------------------------------------------

const PAPER_SIZES = {
    'a4-landscape': { widthMm: 297, heightMm: 210 },
    'a3-landscape': { widthMm: 420, heightMm: 297 }
};
const PDF_MARGIN_MM = 10;
const CAPTURE_SCALE = 2;

// プレビュー用ページ管理
let pdfPageCanvases = [];
let pdfCurrentPage = 0;

// ---------------------------------------------------
// 設定値収集
// ---------------------------------------------------
function collectPdfSettings() {
    return {
        printStart:  document.getElementById('modal-print-start').value,
        printEnd:    document.getElementById('modal-print-end').value,
        zoom:        parseFloat(document.getElementById('modal-print-zoom').value) || 1.0,
        paperSize:   document.getElementById('modal-print-paper').value || 'a4-landscape',
        showNotes:   document.getElementById('modal-print-show-notes').checked,
        showDaily:   document.getElementById('modal-print-show-daily').checked,
        showStamp:   document.getElementById('modal-print-show-stamp').checked,
        stamp1:      document.getElementById('modal-print-stamp1').value || '現場代理人',
        stamp2:      document.getElementById('modal-print-stamp2').value || '監理技術者'
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
// フレーム待機ユーティリティ
// ---------------------------------------------------
function waitFrames(n) {
    return new Promise(resolve => {
        let count = 0;
        function tick() {
            if (++count >= n) resolve();
            else requestAnimationFrame(tick);
        }
        requestAnimationFrame(tick);
    });
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
// メインキャプチャ処理
// ---------------------------------------------------
async function captureGanttToCanvas(settings) {
    // 現在の状態を退避
    const prePdfState = {
        displayStart: state.displayStart,
        displayEnd:   state.displayEnd,
        zoomRatio:    state.zoomRatio,
        viewRange:    state.viewRange
    };

    // 設定を state に反映してレンダリング
    state.displayStart = settings.printStart;
    state.displayEnd   = settings.printEnd;
    state.zoomRatio    = settings.zoom;
    state.viewRange    = 'custom';
    renderAll();

    // 描画安定待機
    await waitFrames(2);

    // notes-pane / daily-notes の表示制御
    const notesPane = document.getElementById('notes-pane');
    if (notesPane) {
        notesPane.style.display = settings.showNotes ? '' : 'none';
    }
    const dailyNotesWrapper = document.querySelector('.daily-notes-wrapper');
    if (dailyNotesWrapper) {
        dailyNotesWrapper.style.display = settings.showDaily ? '' : 'none';
    }

    // 日次備考欄の寸法を退避（print-capture-mode 適用前）
    const dnRightEl = document.getElementById('daily-notes-right');
    const dnGridEl  = document.getElementById('daily-notes-grid');
    const dnRightH  = dnRightEl ? dnRightEl.offsetHeight : 0;
    const dnRightW  = (() => {
        const rc = document.getElementById('right-container');
        return rc ? rc.scrollWidth : (dnRightEl ? dnRightEl.scrollWidth : 0);
    })();

    // キャプチャモード付与
    document.body.classList.add('print-capture-mode');

    // 日次備考欄の高さ・幅を明示的にセット（height:auto による高さ崩壊を防ぐ）
    if (settings.showDaily && dnRightEl && dnRightH > 0) {
        dnRightEl.style.setProperty('height', dnRightH + 'px', 'important');
        dnRightEl.style.setProperty('width',  dnRightW + 'px', 'important');
        if (dnGridEl) {
            dnGridEl.style.setProperty('height', dnRightH + 'px', 'important');
        }
    }

    // DOM寸法固定（executePrint と同パターン）
    const mainContainer = document.querySelector('.main-container');
    const leftBlock = mainContainer ? mainContainer.firstElementChild : null;
    const projectNotes = document.getElementById('project-notes');
    const leftPane = document.getElementById('left-pane');
    const dailyNotesLeft = document.getElementById('daily-notes-left');
    const PRINT_NOTES_MARGIN = 33;
    const COLLAPSED_WIDTH = 350;
    const EXPANDED_WIDTH  = 520;

    if (leftBlock && notesPane) {
        const targetHeight = leftBlock.offsetHeight;
        notesPane.style.setProperty('height', targetHeight + 'px', 'important');
        if (projectNotes) {
            projectNotes.style.setProperty('height', `calc(${targetHeight}px - ${PRINT_NOTES_MARGIN}px)`, 'important');
        }
    }

    if (leftPane && dailyNotesLeft) {
        const targetWidth = leftPane.classList.contains('collapsed-view') ? COLLAPSED_WIDTH : EXPANDED_WIDTH;
        dailyNotesLeft.style.setProperty('width',     targetWidth + 'px', 'important');
        dailyNotesLeft.style.setProperty('min-width', targetWidth + 'px', 'important');
        dailyNotesLeft.style.setProperty('max-width', targetWidth + 'px', 'important');
    }

    // ハンコ枠の制御
    const stampContainer = document.querySelector('.stamp-container');
    const origStampTitles = getStampTitles();
    if (stampContainer) {
        stampContainer.style.display = settings.showStamp ? '' : 'none';
    }
    if (settings.showStamp) {
        setStampTitles(settings.stamp1, settings.stamp2);
    }

    // フォントロード完了を保証
    await document.fonts.ready;

    // 左ペイン幅を記録
    const leftPaneEl = document.getElementById('left-pane');
    const leftPaneWidthPx = leftPaneEl ? leftPaneEl.offsetWidth * CAPTURE_SCALE : 0;

    // キャプチャ実行
    let totalCanvas;
    try {
        totalCanvas = await html2canvas(document.body, {
            scale:           CAPTURE_SCALE,
            backgroundColor: '#ffffff',
            allowTaint:      true,
            useCORS:         true,
            logging:         false,
            ignoreElements:  (el) => {
                return el.classList.contains('modal-overlay') ||
                       el.id === 'color-palette-popup' ||
                       el.classList.contains('context-menu') ||
                       el.classList.contains('print-hide');
            }
        });
    } finally {
        // --- 後始末（必ず実行） ---
        document.body.classList.remove('print-capture-mode');

        // DOM寸法固定解除
        if (notesPane) notesPane.style.removeProperty('height');
        if (projectNotes) projectNotes.style.removeProperty('height');
        if (dailyNotesLeft) {
            dailyNotesLeft.style.removeProperty('width');
            dailyNotesLeft.style.removeProperty('min-width');
            dailyNotesLeft.style.removeProperty('max-width');
            if (leftPane) {
                const currentWidth = leftPane.offsetWidth;
                dailyNotesLeft.style.width    = currentWidth + 'px';
                dailyNotesLeft.style.minWidth = currentWidth + 'px';
            }
        }

        // 日次備考欄の明示的寸法を解除
        if (dnRightEl) {
            dnRightEl.style.removeProperty('height');
            dnRightEl.style.removeProperty('width');
        }
        if (dnGridEl) dnGridEl.style.removeProperty('height');

        // notes-pane / daily-notes を元に戻す
        if (notesPane) notesPane.style.display = '';
        if (dailyNotesWrapper) dailyNotesWrapper.style.display = '';

        // ハンコ枠を元に戻す
        if (stampContainer) stampContainer.style.display = '';
        restoreStampTitles(origStampTitles);

        // state 復元
        state.displayStart = prePdfState.displayStart;
        state.displayEnd   = prePdfState.displayEnd;
        state.zoomRatio    = prePdfState.zoomRatio;
        state.viewRange    = prePdfState.viewRange;
        renderAll();

        // スクロール位置同期
        await waitFrames(1);
        const rightContainer   = document.getElementById('right-container');
        const dailyNotesRight  = document.getElementById('daily-notes-right');
        if (rightContainer && dailyNotesRight) {
            dailyNotesRight.scrollLeft = rightContainer.scrollLeft;
        }
    }

    return { totalCanvas, leftPaneWidthPx };
}

// ---------------------------------------------------
// ページスライス・合成
// ---------------------------------------------------
function sliceAndCompositePages(totalCanvas, leftPaneWidthPx, settings) {
    const paper = PAPER_SIZES[settings.paperSize] || PAPER_SIZES['a4-landscape'];
    const dpi   = 96;

    const pageWidthPx  = Math.round((paper.widthMm  - PDF_MARGIN_MM * 2) / 25.4 * dpi * CAPTURE_SCALE);
    const pageHeightPx = Math.round((paper.heightMm - PDF_MARGIN_MM * 2) / 25.4 * dpi * CAPTURE_SCALE);

    const rightPaneFullWidth = totalCanvas.width - leftPaneWidthPx;
    const rightPanePerPage   = pageWidthPx - leftPaneWidthPx;

    const pages = [];
    let rightOffset = 0;
    let isFirst = true;

    while (rightOffset < rightPaneFullWidth || isFirst) {
        const canvas = document.createElement('canvas');
        canvas.width  = pageWidthPx;
        canvas.height = pageHeightPx;
        const ctx = canvas.getContext('2d');

        // 白背景
        ctx.fillStyle = '#ffffff';
        ctx.fillRect(0, 0, pageWidthPx, pageHeightPx);

        if (isFirst) {
            // 1ページ目: そのまま左上からコピー
            ctx.drawImage(
                totalCanvas,
                0, 0, pageWidthPx, pageHeightPx,
                0, 0, pageWidthPx, pageHeightPx
            );
        } else {
            // 2ページ目以降: 左ペイン + 右ペイン（右オフセット分ずらす）
            // 左ペイン
            ctx.drawImage(
                totalCanvas,
                0, 0, leftPaneWidthPx, pageHeightPx,
                0, 0, leftPaneWidthPx, pageHeightPx
            );
            // 右ペイン（右方向にオフセット）
            const srcRightWidth = Math.min(rightPanePerPage, rightPaneFullWidth - rightOffset);
            ctx.drawImage(
                totalCanvas,
                leftPaneWidthPx + rightOffset, 0, srcRightWidth, pageHeightPx,
                leftPaneWidthPx, 0, srcRightWidth, pageHeightPx
            );
        }

        pages.push(canvas);

        if (isFirst) {
            // 1ページ目で右ペインが収まる範囲をオフセットに反映
            rightOffset = rightPanePerPage - leftPaneWidthPx;
            // 補正: 1ページ目は leftPaneWidthPx の分だけ右ペインが少ない
            rightOffset = pageWidthPx - leftPaneWidthPx;
        } else {
            rightOffset += rightPanePerPage;
        }

        isFirst = false;

        if (rightOffset >= rightPaneFullWidth) break;
    }

    return pages;
}

// ---------------------------------------------------
// PDF構築
// ---------------------------------------------------
function buildPdfFromPages(pageCanvases, paperSize) {
    const paper = PAPER_SIZES[paperSize] || PAPER_SIZES['a4-landscape'];
    const formatName = paperSize.startsWith('a3') ? 'a3' : 'a4';
    const { jsPDF } = window.jspdf;

    const pdf = new jsPDF({
        orientation: 'landscape',
        unit:        'mm',
        format:      formatName
    });

    const printableWidthMm  = paper.widthMm  - PDF_MARGIN_MM * 2;
    const printableHeightMm = paper.heightMm - PDF_MARGIN_MM * 2;

    pageCanvases.forEach((canvas, i) => {
        if (i > 0) pdf.addPage();
        const imgData = canvas.toDataURL('image/jpeg', 0.92);
        pdf.addImage(imgData, 'JPEG', PDF_MARGIN_MM, PDF_MARGIN_MM, printableWidthMm, printableHeightMm);
    });

    return pdf;
}

// ---------------------------------------------------
// ファイル名生成
// ---------------------------------------------------
function generateDefaultFilename(ext) {
    const name = (state.projectName || '工程表').replace(/[\\/:*?"<>|]/g, '_');
    const now  = new Date();
    const ymd  = `${now.getFullYear()}${String(now.getMonth()+1).padStart(2,'0')}${String(now.getDate()).padStart(2,'0')}`;
    return `${name}_${ymd}.${ext}`;
}

// ---------------------------------------------------
// プレビュー表示
// ---------------------------------------------------
window.updatePdfPreview = async function () {
    showPdfLoading('プレビュー生成中...');
    try {
        const settings = collectPdfSettings();
        const { totalCanvas, leftPaneWidthPx } = await captureGanttToCanvas(settings);
        pdfPageCanvases = sliceAndCompositePages(totalCanvas, leftPaneWidthPx, settings);
        pdfCurrentPage  = 0;
        renderPreviewPage();
    } catch (e) {
        console.error('プレビュー生成エラー:', e);
        alert('プレビューの生成に失敗しました:\n' + e.message);
    } finally {
        hidePdfLoading();
    }
};

function renderPreviewPage() {
    const area  = document.getElementById('pdf-preview-area');
    const label = document.getElementById('pdf-preview-page-label');
    if (!pdfPageCanvases.length) return;

    const total   = pdfPageCanvases.length;
    const current = pdfCurrentPage + 1;
    if (label) label.textContent = `ページ: ${current} / ${total}`;

    area.innerHTML = '';
    const src = pdfPageCanvases[pdfCurrentPage];
    // 表示用に縮小コピー
    const displayCanvas = document.createElement('canvas');
    const maxW = area.clientWidth  - 32 || 800;
    const maxH = area.clientHeight - 32 || 500;
    const scale = Math.min(maxW / src.width, maxH / src.height, 1);
    displayCanvas.width  = Math.round(src.width  * scale);
    displayCanvas.height = Math.round(src.height * scale);
    const ctx = displayCanvas.getContext('2d');
    ctx.drawImage(src, 0, 0, displayCanvas.width, displayCanvas.height);
    area.appendChild(displayCanvas);
}

window.navigatePdfPage = function (delta) {
    if (!pdfPageCanvases.length) return;
    pdfCurrentPage = Math.max(0, Math.min(pdfPageCanvases.length - 1, pdfCurrentPage + delta));
    renderPreviewPage();
};

// ---------------------------------------------------
// PDF保存
// ---------------------------------------------------
window.executePdfSave = async function () {
    showPdfLoading('PDF生成中...');
    try {
        const settings = collectPdfSettings();
        const { totalCanvas, leftPaneWidthPx } = await captureGanttToCanvas(settings);
        const pages = sliceAndCompositePages(totalCanvas, leftPaneWidthPx, settings);
        const pdf   = buildPdfFromPages(pages, settings.paperSize);

        // data URI として取得
        const dataUri = pdf.output('datauristring');
        const defaultName = generateDefaultFilename('pdf');

        if (window.pywebview && window.pywebview.api && window.pywebview.api.save_pdf_file) {
            const result = await window.pywebview.api.save_pdf_file(dataUri, defaultName);
            if (result) {
                if (window.showToast) window.showToast('PDFを保存しました');
            }
        } else {
            // フォールバック: ブラウザダウンロード
            pdf.save(defaultName);
        }
    } catch (e) {
        console.error('PDF保存エラー:', e);
        alert('PDFの保存に失敗しました:\n' + e.message);
    } finally {
        hidePdfLoading();
    }
};

// ---------------------------------------------------
// PNG保存
// ---------------------------------------------------
window.executePngSave = async function () {
    showPdfLoading('PNG生成中...');
    try {
        const settings = collectPdfSettings();
        const { totalCanvas } = await captureGanttToCanvas(settings);

        const dataUri     = totalCanvas.toDataURL('image/png');
        const defaultName = generateDefaultFilename('png');

        if (window.pywebview && window.pywebview.api && window.pywebview.api.save_image_file) {
            const result = await window.pywebview.api.save_image_file(dataUri, defaultName);
            if (result) {
                if (window.showToast) window.showToast('PNGを保存しました');
            }
        } else {
            // フォールバック: ブラウザダウンロード
            const a = document.createElement('a');
            a.href     = dataUri;
            a.download = defaultName;
            a.click();
        }
    } catch (e) {
        console.error('PNG保存エラー:', e);
        alert('PNGの保存に失敗しました:\n' + e.message);
    } finally {
        hidePdfLoading();
    }
};
