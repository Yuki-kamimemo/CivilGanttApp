// ---------------------------------------------------
// モーダル・コンテキストメニュー操作群
// ---------------------------------------------------
window.openHolidayModal = function () {
    document.getElementById('modal-sun').checked = state.holidays.sundays;
    document.getElementById('modal-sat').checked = state.holidays.saturdays;
    document.getElementById('modal-national').checked = state.holidays.nationalHolidays;
    window.renderCustomHolidays();
    document.getElementById('holiday-modal').style.display = 'flex';
};
window.closeHolidayModal = function () { document.getElementById('holiday-modal').style.display = 'none'; };

window.renderCustomHolidays = function () {
    const list = document.getElementById('custom-holiday-list'); list.innerHTML = '';
    state.holidays.custom.sort().forEach(date => {
        list.innerHTML += `<li>${date} <button type="button" onclick="window.removeCustomHoliday('${date}')">×</button></li>`;
    });
};
window.addCustomHoliday = function () {
    const val = document.getElementById('custom-holiday-input').value;
    if (val && !state.holidays.custom.includes(val)) {
        state.holidays.custom.push(val);
        window.renderCustomHolidays();
        document.getElementById('custom-holiday-input').value = '';
    }
};
window.removeCustomHoliday = function (date) {
    state.holidays.custom = state.holidays.custom.filter(d => d !== date);
    window.renderCustomHolidays();
};
window.saveHolidaySettings = function () {
    state.holidays.sundays = document.getElementById('modal-sun').checked;
    state.holidays.saturdays = document.getElementById('modal-sat').checked;
    state.holidays.nationalHolidays = document.getElementById('modal-national').checked;
    window.closeHolidayModal();
    window.saveStateToHistory();
    renderAll();
};

window.openPeriodModal = function (taskId, periodId) {
    const task = state.tasks.find(t => t.id === taskId);
    if (!task) return;
    const period = task.periods.find(p => p.pid === periodId);
    if (!period) return;

    editModalTaskId = taskId;
    editModalPeriodId = periodId;

    document.getElementById('modal-period-start').value = period.start || '';
    document.getElementById('modal-period-end').value = period.end || '';
    document.getElementById('modal-period-days').value = calcDiffDays(period.start, period.end) || '';
    document.getElementById('modal-period-progress').value = period.progress || 0;
    document.getElementById('modal-period-color').value = period.color || '#3b82f6';

    document.getElementById('period-modal').style.display = 'flex';
};

window.closePeriodModal = function () {
    document.getElementById('period-modal').style.display = 'none';
    editModalTaskId = null;
    editModalPeriodId = null;
};

window.savePeriodModal = function () {
    if (!editModalTaskId || !editModalPeriodId) return;

    const task = state.tasks.find(t => t.id === editModalTaskId);
    const period = task.periods.find(p => p.pid === editModalPeriodId);

    const newStart = document.getElementById('modal-period-start').value;
    const newEnd = document.getElementById('modal-period-end').value;
    let newProgress = parseInt(document.getElementById('modal-period-progress').value) || 0;
    const newColor = document.getElementById('modal-period-color').value;

    const oldStart = period.start;
    let diffWorkDays = 0;

    const snappedStart = snapToWorkDay(newStart, 1);
    const snappedEnd = snapToWorkDay(newEnd, -1);

    if (oldStart && snappedStart) {
        diffWorkDays = getWorkDayShift(oldStart, snappedStart);
    }

    period.start = snappedStart;
    period.end = snappedEnd;
    if (newProgress < 0) newProgress = 0;
    if (newProgress > 100) newProgress = 100;
    period.progress = newProgress;
    period.color = newColor;

    if (period.start && period.end) {
        checkAndExtendProjectDates(period.start, period.end);
    }

    if (diffWorkDays !== 0) shiftDependentTasks(period.pid, diffWorkDays);

    window.closePeriodModal();
    window.saveStateToHistory();
    renderAll();
};

// ---------------------------------------------------
// カラーパレット
// ---------------------------------------------------
const PALETTE_COLORS = [
    '#000000', '#343a40', '#6c757d', '#adb5bd', '#ffffff',
    '#dc3545', '#fd7e14', '#ffc107', '#198754', '#0d6efd'
];

window.showColorPalette = function (event, prop) {
    event.stopPropagation();
    const popup = document.getElementById('color-palette-popup');
    // 同じボタンを再クリックしたら閉じる
    if (popup.dataset.prop === prop && popup.style.display !== 'none') {
        popup.style.display = 'none';
        return;
    }
    popup.dataset.prop = prop;

    const grid = document.getElementById('color-swatches-grid');
    grid.innerHTML = '';

    // 塗りつぶし色のみ「なし」スウォッチを先頭に追加
    if (prop === 'backgroundColor') {
        const clearBtn = document.createElement('button');
        clearBtn.type = 'button';
        clearBtn.className = 'color-swatch transparent-swatch';
        clearBtn.title = '塗りつぶしなし (transparent)';
        clearBtn.onclick = (e) => {
            e.stopPropagation();
            window.handleFormatChange(prop, 'transparent');
            popup.style.display = 'none';
        };
        grid.appendChild(clearBtn);
    }

    PALETTE_COLORS.forEach(color => {
        const btn = document.createElement('button');
        btn.type = 'button';
        btn.className = 'color-swatch';
        btn.style.backgroundColor = color;
        btn.title = color;
        btn.onclick = (e) => {
            e.stopPropagation();
            // カラーピッカーの値も同期
            const inputId = { color: 'toolbar-text-color', backgroundColor: 'toolbar-bg-color', borderColor: 'toolbar-border-color' }[prop];
            const input = document.getElementById(inputId);
            if (input) input.value = color;
            window.handleFormatChange(prop, color);
            popup.style.display = 'none';
        };
        grid.appendChild(btn);
    });

    // ポップアップの表示位置をボタンの直下に設定
    const rect = event.currentTarget.getBoundingClientRect();
    popup.style.left = rect.left + 'px';
    popup.style.top = (rect.bottom + 4) + 'px';
    popup.style.display = 'block';
};

// パレット外クリックで閉じる
document.addEventListener('click', () => {
    const popup = document.getElementById('color-palette-popup');
    if (popup) popup.style.display = 'none';
});

// ---------------------------------------------------
// 書式設定（ツールバー）
// ---------------------------------------------------
let savedSelectionRange = null;
let savedSelectionNode = null;

document.addEventListener('selectionchange', () => {
    const sel = window.getSelection();
    if (sel.rangeCount > 0) {
        const node = sel.anchorNode;
        const el = node.nodeType === 3 ? node.parentNode : node;
        if (el && el.isContentEditable) {
            savedSelectionRange = sel.getRangeAt(0).cloneRange();
            savedSelectionNode = el.closest('[contenteditable="true"]') || el;
        }
    }
});

// -------------------------------------------------------
// ツールバー・パレットのクリックでcontentEditableのフォーカスを
// 失わないようにする（選択テキストへの書式適用を可能にする）
// -------------------------------------------------------
const _formatToolbar = document.querySelector('.format-toolbar');
if (_formatToolbar) {
    _formatToolbar.addEventListener('mousedown', (e) => {
        // input・select は通常のフォーカス動作が必要なので対象外
        if (e.target.tagName !== 'INPUT' && e.target.tagName !== 'SELECT') {
            e.preventDefault();
        }
    });
}
const _colorPalettePopup = document.getElementById('color-palette-popup');
if (_colorPalettePopup) {
    _colorPalettePopup.addEventListener('mousedown', (e) => e.preventDefault());
}

window.updateFormatToolbar = function () {
    const textOnlyGroup = document.getElementById('text-only-formats');
    let target = null;

    if (selectedItem) {
        if (selectedItem.type === 'text') {
            target = state.texts.find(t => t.id === selectedItem.textId);
        } else if (selectedItem.type === 'cell') {
            const task = state.tasks.find(t => t.id === selectedItem.taskId);
            if (task && task.styles) target = task.styles[selectedItem.field];
        } else if (selectedItem.type === 'notes') {
            target = state.notesStyle;
        } else if (selectedItem.type === 'daily_notes') {
            target = state.dailyNotesStyle;
        }
    }

    if (target || (selectedItem && ['text', 'cell', 'notes', 'daily_notes'].includes(selectedItem.type))) {
        if (textOnlyGroup) {
            textOnlyGroup.style.opacity = '1';
            textOnlyGroup.style.pointerEvents = 'auto';
        }
        const t = target || {};
        document.getElementById('toolbar-font-family').value = t.fontFamily || state.globalFontFamily || "'Segoe UI', Tahoma, Geneva, Verdana, sans-serif";
        document.getElementById('toolbar-font-size').value = t.fontSize || state.globalFontSize || 13;

        const textColor = t.color || '#212529';
        document.getElementById('toolbar-text-color').value = textColor;
        const previewColor = document.getElementById('preview-color');
        if (previewColor) { previewColor.style.backgroundImage = ''; previewColor.style.backgroundColor = textColor; }

        const bgColor = (t.backgroundColor && t.backgroundColor !== 'transparent') ? t.backgroundColor : null;
        document.getElementById('toolbar-bg-color').value = bgColor || '#ffffff';
        const previewBg = document.getElementById('preview-backgroundColor');
        if (previewBg) {
            if (bgColor) {
                previewBg.style.backgroundImage = ''; previewBg.style.backgroundColor = bgColor;
            } else {
                previewBg.style.backgroundColor = 'white';
                previewBg.style.backgroundImage = 'linear-gradient(45deg,#ccc 25%,transparent 25%),linear-gradient(-45deg,#ccc 25%,transparent 25%),linear-gradient(45deg,transparent 75%,#ccc 75%),linear-gradient(-45deg,transparent 75%,#ccc 75%)';
                previewBg.style.backgroundSize = '6px 6px';
                previewBg.style.backgroundPosition = '0 0,0 3px,3px -3px,-3px 0';
            }
        }

        document.getElementById('toolbar-border-style').value = t.borderStyle || 'solid';
        document.getElementById('toolbar-border-width').value = t.borderWidth !== undefined ? t.borderWidth : 1;
        const borderColor = t.borderColor || '#6c757d';
        document.getElementById('toolbar-border-color').value = borderColor;
        const previewBorder = document.getElementById('preview-borderColor');
        if (previewBorder) { previewBorder.style.backgroundImage = ''; previewBorder.style.backgroundColor = borderColor; }

        document.getElementById('btn-bold').classList.toggle('active', t.fontWeight === 'bold');

        const isDailyNotes = (selectedItem && selectedItem.type === 'daily_notes');
        const align = t.textAlign || (isDailyNotes ? 'center' : 'left');
        document.getElementById('btn-align-left').classList.toggle('active', align === 'left');
        document.getElementById('btn-align-center').classList.toggle('active', align === 'center');
        document.getElementById('btn-align-right').classList.toggle('active', align === 'right');

        const wm = t.writingMode || (isDailyNotes ? 'vertical-rl' : 'horizontal-tb');
        const btnWm = document.getElementById('btn-writing-mode');
        if (btnWm) btnWm.classList.toggle('active', wm === 'vertical-rl');

        const valign = t.verticalAlign || 'flex-start';
        const btnValignTop = document.getElementById('btn-valign-top');
        const btnValignMiddle = document.getElementById('btn-valign-middle');
        const btnValignBottom = document.getElementById('btn-valign-bottom');
        if (btnValignTop) btnValignTop.classList.toggle('active', valign === 'flex-start');
        if (btnValignMiddle) btnValignMiddle.classList.toggle('active', valign === 'center');
        if (btnValignBottom) btnValignBottom.classList.toggle('active', valign === 'flex-end');
    } else {
        if (textOnlyGroup) {
            textOnlyGroup.style.opacity = '0.5';
            textOnlyGroup.style.pointerEvents = 'none';
        }
        document.getElementById('toolbar-font-family').value = state.globalFontFamily || "'Segoe UI', Tahoma, Geneva, Verdana, sans-serif";
        document.getElementById('toolbar-font-size').value = state.globalFontSize || 13;
        document.querySelectorAll('#text-only-formats .format-btn').forEach(btn => btn.classList.remove('active'));
    }
};

window.handleFormatChange = function (prop, value) {
    const selection = window.getSelection();
    let activeEl = document.activeElement;

    if (savedSelectionRange && savedSelectionNode) {
        // chart-text-box はカラーピッカー等で blur すると contentEditable=false になるため
        // 一時的に再有効化して選択テキストへの書式適用を可能にする
        const isChartTextBox = savedSelectionNode.classList &&
            savedSelectionNode.classList.contains('chart-text-box');
        if (savedSelectionNode.isContentEditable || isChartTextBox) {
            if (!activeEl || !activeEl.isContentEditable) {
                try {
                    if (isChartTextBox && !savedSelectionNode.isContentEditable) {
                        savedSelectionNode.contentEditable = 'true';
                    }
                    savedSelectionNode.focus();
                    selection.removeAllRanges();
                    selection.addRange(savedSelectionRange);
                    activeEl = savedSelectionNode;
                } catch (e) {
                    // rangeが無効な場合はクリアして通常パスへ
                    if (isChartTextBox && savedSelectionNode) {
                        savedSelectionNode.contentEditable = 'false';
                    }
                    savedSelectionRange = null;
                    savedSelectionNode = null;
                }
            }
        }
    }

    if (activeEl && activeEl.isContentEditable) {
        if (prop === 'fontWeight') {
            document.execCommand('bold', false, null);
        } else if (prop === 'color') {
            // execCommand('foreColor') はブロック境界で視覚的な改行を生じることがあるため
            // 選択範囲を <span> でラップして color スタイルを直接適用する
            const sel = window.getSelection();
            if (sel && sel.rangeCount > 0 && !sel.isCollapsed) {
                const range = sel.getRangeAt(0);
                const span = document.createElement('span');
                span.style.color = value;
                try {
                    range.surroundContents(span);
                } catch (err) {
                    // 複数ノードにまたがる選択の場合は execCommand にフォールバック
                    document.execCommand('foreColor', false, value);
                }
            } else {
                document.execCommand('foreColor', false, value);
            }
        } else if (prop === 'backgroundColor') {
            if (!document.execCommand('hiliteColor', false, value)) {
                document.execCommand('backColor', false, value);
            }
        } else if (prop === 'fontFamily') {
            document.execCommand('fontName', false, value);
        } else if (prop === 'fontSize') {
            // execCommand('fontSize') のサイズ指定は非標準のため、
            // 選択範囲を <span> でラップして font-size を直接適用する
            const sel = window.getSelection();
            if (sel && sel.rangeCount > 0 && !sel.isCollapsed) {
                const range = sel.getRangeAt(0);
                const span = document.createElement('span');
                span.style.fontSize = value + 'px';
                try {
                    range.surroundContents(span);
                } catch (err) {
                    // 複数ノードにまたがる選択の場合はフォールバック
                    document.execCommand('fontSize', false, '7');
                    const fonts = activeEl.getElementsByTagName('font');
                    for (let i = fonts.length - 1; i >= 0; i--) {
                        const f = fonts[i];
                        // size属性が7、またはstyleがxxx-largeの要素を対象にする
                        if (f.getAttribute('size') === '7' || f.style.fontSize === 'xxx-large') {
                            f.removeAttribute('size');
                            f.style.fontSize = value + 'px';
                        }
                    }
                }
            } else {
                // テキスト未選択時はテキストボックス全体のスタイルを変更
                activeEl.style.fontSize = value + 'px';
            }
        } else if (prop === 'textAlign') {
            if (value === 'left') document.execCommand('justifyLeft', false, null);
            else if (value === 'center') document.execCommand('justifyCenter', false, null);
            else if (value === 'right') document.execCommand('justifyRight', false, null);
        }

        if (activeEl.classList.contains('chart-text-box')) {
            const txt = state.texts.find(t => t.id === activeEl.dataset.id);
            if (txt) txt.text = activeEl.innerHTML;
        } else if (activeEl.id === 'project-notes') {
            state.notes = activeEl.innerHTML;
        }

        if (selection.rangeCount > 0) {
            savedSelectionRange = selection.getRangeAt(0).cloneRange();
        }

        window.saveStateToHistory();
        return;
    }

    let target = null;
    if (selectedItem) {
        if (selectedItem.type === 'text') {
            target = state.texts.find(t => t.id === selectedItem.textId);
        } else if (selectedItem.type === 'cell') {
            const task = state.tasks.find(t => t.id === selectedItem.taskId);
            if (task) {
                if (!task.styles) task.styles = {};
                if (!task.styles[selectedItem.field]) task.styles[selectedItem.field] = {};
                target = task.styles[selectedItem.field];
            }
        } else if (selectedItem.type === 'notes') {
            if (!state.notesStyle) state.notesStyle = {};
            target = state.notesStyle;
        } else if (selectedItem.type === 'daily_notes') {
            if (!state.dailyNotesStyle) state.dailyNotesStyle = {};
            target = state.dailyNotesStyle;
        }
    }

    if (target) {
        if (prop === 'fontWeight') {
            target.fontWeight = (target.fontWeight === 'bold') ? 'normal' : 'bold';
        } else if (prop === 'writingMode') {
            target.writingMode = (target.writingMode === 'vertical-rl') ? 'horizontal-tb' : 'vertical-rl';
        } else {
            target[prop] = value;
        }

        if (selectedItem && selectedItem.type === 'text') {
            const txt = state.texts.find(t => t.id === selectedItem.textId);
            if (txt && txt.text) {
                const tempDiv = document.createElement('div');
                tempDiv.innerHTML = txt.text;
                const elements = tempDiv.querySelectorAll('*');
                elements.forEach(el => {
                    if (prop === 'color' || prop === 'backgroundColor') { el.style[prop] = ''; el.removeAttribute('color'); el.removeAttribute('bgcolor'); }
                    if (prop === 'fontSize') { el.style.fontSize = ''; el.removeAttribute('size'); }
                    if (prop === 'fontFamily') { el.style.fontFamily = ''; el.removeAttribute('face'); }
                });
                txt.text = tempDiv.innerHTML;
            }
        }
    } else {
        if (prop === 'fontFamily') state.globalFontFamily = value;
        if (prop === 'fontSize') state.globalFontSize = parseInt(value) || 13;
    }

    window.saveStateToHistory();
    renderAll();
    window.updateFormatToolbar();
};

window.clearSavedSelection = function () {
    savedSelectionRange = null;
    savedSelectionNode = null;
};

window.selectInput = function (type, taskId = null, field = null) {
    selectedItem = { type: type, taskId: taskId, field: field };
    // 新しい要素が選択されたとき、以前のcontentEditable編集セッションの
    // 選択範囲を消去する。これにより書式が誤った要素に適用されるのを防ぐ。
    savedSelectionRange = null;
    savedSelectionNode = null;
    window.updateFormatToolbar();
};

// ---------------------------------------------------
// 独自モーダル制御
// ---------------------------------------------------
let textInputCallback = null;
let confirmCallback = null;

window.openTextInputModal = function (title, defaultValue, callback) {
    document.getElementById('text-input-title').textContent = title;
    const inputField = document.getElementById('text-input-field');
    inputField.value = defaultValue;
    textInputCallback = callback;
    document.getElementById('text-input-modal').style.display = 'flex';
    inputField.focus();
};
window.closeTextInputModal = function () {
    document.getElementById('text-input-modal').style.display = 'none';
    textInputCallback = null;
};
window.saveTextInputModal = function () {
    const val = document.getElementById('text-input-field').value;
    if (textInputCallback) textInputCallback(val);
    window.closeTextInputModal();
};

window.openConfirmModal = function (title, message, callback) {
    document.getElementById('confirm-title').textContent = title;
    document.getElementById('confirm-message').textContent = message;
    confirmCallback = callback;
    document.getElementById('confirm-modal').style.display = 'flex';
};
window.closeConfirmModal = function () {
    document.getElementById('confirm-modal').style.display = 'none';
    confirmCallback = null;
};
window.executeConfirmModal = function () {
    if (confirmCallback) confirmCallback();
    window.closeConfirmModal();
};

// ---------------------------------------------------
// 印刷制御
// ---------------------------------------------------
let prePrintState = null;

window.openPrintModal = function () {
    document.getElementById('modal-print-start').value = state.displayStart;
    document.getElementById('modal-print-end').value = state.displayEnd;
    document.getElementById('modal-print-zoom').value = state.zoomRatio || 1.0;
    document.getElementById('print-modal').style.display = 'flex';
};

window.closePrintModal = function () {
    document.getElementById('print-modal').style.display = 'none';
};

window.executePrint = function () {
    prePrintState = {
        displayStart: state.displayStart,
        displayEnd: state.displayEnd,
        zoomRatio: state.zoomRatio,
        viewRange: state.viewRange
    };

    state.displayStart = document.getElementById('modal-print-start').value;
    state.displayEnd = document.getElementById('modal-print-end').value;
    state.zoomRatio = parseFloat(document.getElementById('modal-print-zoom').value) || 1.0;
    state.viewRange = 'custom';

    window.closePrintModal();
    renderAll();

    setTimeout(() => {
        const mainContainer = document.querySelector('.main-container');
        const leftBlock = mainContainer.firstElementChild;
        const notesPane = document.getElementById('notes-pane');
        const projectNotes = document.getElementById('project-notes');

        const targetHeight = leftBlock.offsetHeight;
        if (notesPane) notesPane.style.setProperty('height', targetHeight + 'px', 'important');
        if (projectNotes) projectNotes.style.setProperty('height', 'calc(' + targetHeight + 'px - 33px)', 'important');

        const leftPane = document.getElementById('left-pane');
        const dailyNotesLeft = document.getElementById('daily-notes-left');
        if (leftPane && dailyNotesLeft) {
            const targetWidth = leftPane.classList.contains('collapsed-view') ? 350 : 520;
            dailyNotesLeft.style.setProperty('width', targetWidth + 'px', 'important');
            dailyNotesLeft.style.setProperty('min-width', targetWidth + 'px', 'important');
            dailyNotesLeft.style.setProperty('max-width', targetWidth + 'px', 'important');
        }

        window.print();

        setTimeout(() => {
            if (notesPane) notesPane.style.removeProperty('height');
            if (projectNotes) projectNotes.style.removeProperty('height');

            if (dailyNotesLeft) {
                dailyNotesLeft.style.removeProperty('width');
                dailyNotesLeft.style.removeProperty('min-width');
                dailyNotesLeft.style.removeProperty('max-width');

                // ★追加: レイアウト崩れを防ぐため、上の表の幅に合わせて強制的に再設定する
                if (leftPane) {
                    const currentWidth = leftPane.offsetWidth;
                    dailyNotesLeft.style.width = currentWidth + 'px';
                    dailyNotesLeft.style.minWidth = currentWidth + 'px';
                }
            }

            state.displayStart = prePrintState.displayStart;
            state.displayEnd = prePrintState.displayEnd;
            state.zoomRatio = prePrintState.zoomRatio;
            state.viewRange = prePrintState.viewRange;
            renderAll();

            // ★追加: 印刷画面から戻った際に、左右のスクロール位置を同期させる
            setTimeout(() => {
                const rightContainer = document.getElementById('right-container');
                const dailyNotesRight = document.getElementById('daily-notes-right');
                if (rightContainer && dailyNotesRight) {
                    dailyNotesRight.scrollLeft = rightContainer.scrollLeft;
                }
            }, 50);
        }, 500);
    }, 500);
};

// ---------------------------------------------------
// コンテキストメニュー制御
// ---------------------------------------------------
window.showContextMenu = function (e, title, type, data) {
    e.preventDefault();
    const menu = document.getElementById('custom-context-menu');
    document.getElementById('ctx-menu-title').textContent = title;
    const actions = document.getElementById('ctx-menu-actions');
    actions.innerHTML = '';

    selectedItem = { type, ...data };

    if (type === 'bar') {
        actions.innerHTML = `
            <div class="menu-action-item" onclick="window.handleContextAction('edit_period')">詳細設定</div>
            <div style="border-top: 1px solid #eee; margin: 4px 0;"></div>
            <div class="menu-action-item delete-action" onclick="window.handleContextAction('delete_period')">この期間を削除</div>
        `;
    } else if (type === 'arrow') {
        actions.innerHTML = `<div class="menu-action-item delete-action" onclick="window.handleContextAction('delete_arrow')">依存関係を削除</div>`;
    } else if (type === 'task') {
        const pasteDisabled = copiedTask ? '' : 'disabled';
        const taskIndex = state.tasks.findIndex(t => t.id === data.taskId);
        const task = state.tasks[taskIndex];
        const canMergeAbove = taskIndex > 0;

        let mergeActions = '';
        if (canMergeAbove) {
            const koshuMergeText = task.mergeAboveKoshu ? "工種の結合を解除" : "工種を上の行と結合";
            const shubetsuMergeText = task.mergeAboveShubetsu ? "種別の結合を解除" : "種別を上の行と結合";
            mergeActions = `
                <div class="menu-action-item" onclick="window.handleContextAction('toggle_merge_koshu')">${koshuMergeText}</div>
                <div class="menu-action-item" onclick="window.handleContextAction('toggle_merge_shubetsu')">${shubetsuMergeText}</div>
                <div style="border-top: 1px solid #eee; margin: 4px 0;"></div>
            `;
        }

        actions.innerHTML = `
            <div class="menu-action-item" onclick="window.handleContextAction('copy_task')">行をコピー</div>
            <div class="menu-action-item ${pasteDisabled}" onclick="if(!this.classList.contains('disabled')) window.handleContextAction('paste_task')">下に貼り付け</div>
            <div style="border-top: 1px solid #eee; margin: 4px 0;"></div>
            ${mergeActions}
            <div class="menu-action-item delete-action" onclick="window.handleContextAction('delete_task')">この行を削除</div>
        `;
    } else if (type === 'text') {
        actions.innerHTML = `
            <div class="menu-action-item" onclick="window.handleContextAction('edit_text')">テキストを編集</div>
            <div style="border-top: 1px solid #eee; margin: 4px 0;"></div>
            <div class="menu-action-item delete-action" onclick="window.handleContextAction('delete_text')">削除</div>
        `;
    }

    menu.style.display = 'block';
    menu.style.left = e.clientX + 'px';
    menu.style.top = e.clientY + 'px';
};

window.handleContextAction = function (action) {
    if (!selectedItem) return;

    if (action === 'edit_period') {
        window.openPeriodModal(selectedItem.taskId, selectedItem.periodId);
    } else if (action === 'delete_period') {
        const task = state.tasks.find(t => t.id === selectedItem.taskId);
        if (task && task.periods.length > 1) {
            window.handleRemovePeriod(selectedItem.taskId, selectedItem.periodId);
        } else {
            alert('最後の1つの期間は削除できません。\n行ごと削除するには行の右クリックから「この行を削除」を使用してください。');
        }
    } else if (action === 'delete_arrow') {
        const task = state.tasks.find(t => t.id === selectedItem.taskId);
        const period = task ? task.periods.find(p => p.pid === selectedItem.periodId) : null;
        if (period) {
            let deps = period.dep ? period.dep.toString().split(',').map(s => s.trim()) : [];
            deps = deps.filter(d => d !== selectedItem.predPid);
            window.handlePeriodChange(selectedItem.taskId, selectedItem.periodId, 'dep', deps.join(', '));
        }
    } else if (action === 'copy_task') {
        const task = state.tasks.find(t => t.id === selectedItem.taskId);
        if (task) copiedTask = JSON.parse(JSON.stringify(task));
    } else if (action === 'paste_task') {
        if (!copiedTask) return;
        const targetIndex = state.tasks.findIndex(t => t.id === selectedItem.taskId);
        if (targetIndex !== -1) {
            const newTask = JSON.parse(JSON.stringify(copiedTask));
            newTask.id = generateId();
            newTask.periods.forEach(p => p.pid = generateId());
            newTask.mergeAboveKoshu = false;
            newTask.mergeAboveShubetsu = false;
            state.tasks.splice(targetIndex + 1, 0, newTask);
            window.saveStateToHistory();
            renderAll();
        }
    } else if (action === 'delete_task') {
        const targetId = selectedItem.taskId;
        window.openConfirmModal('行の削除', 'この行（タスク）を完全に削除しますか？', function () {
            const targetIndex = state.tasks.findIndex(t => t.id === targetId);
            if (targetIndex !== -1 && targetIndex < state.tasks.length - 1) {
                state.tasks[targetIndex + 1].mergeAboveKoshu = false;
                state.tasks[targetIndex + 1].mergeAboveShubetsu = false;
            }
            state.tasks = state.tasks.filter(t => t.id !== targetId);
            window.saveStateToHistory();
            renderAll();
        });
    } else if (action === 'toggle_merge_koshu') {
        const task = state.tasks.find(t => t.id === selectedItem.taskId);
        if (task) {
            task.mergeAboveKoshu = !task.mergeAboveKoshu;
            window.saveStateToHistory(); renderAll();
        }
    } else if (action === 'toggle_merge_shubetsu') {
        const task = state.tasks.find(t => t.id === selectedItem.taskId);
        if (task) {
            task.mergeAboveShubetsu = !task.mergeAboveShubetsu;
            window.saveStateToHistory(); renderAll();
        }
    } else if (action === 'edit_text') {
        const txt = state.texts.find(t => t.id === selectedItem.textId);
        if (txt) {
            window.openTextInputModal('テキストの文字を編集', txt.text, function (val) {
                if (val) { txt.text = val; window.saveStateToHistory(); renderAll(); }
            });
        }
    } else if (action === 'delete_text') {
        const targetTextId = selectedItem.textId;
        window.openConfirmModal('削除', 'このテキストを削除しますか？', function () {
            state.texts = state.texts.filter(t => t.id !== targetTextId);
            window.saveStateToHistory(); renderAll();
        });
    }

    selectedItem = null;
    document.getElementById('custom-context-menu').style.display = 'none';
};

window.handleRowContextMenu = function (e, taskId, taskNo, koshu) {
    if (e.target.tagName === 'INPUT' && e.target.type !== 'color') return;
    const title = `行アクション (No.${taskNo} ${koshu || '名称未設定'})`;
    window.showContextMenu(e, title, 'task', { taskId });
};

// --- デバッグ用：書式設定状態の確認 ---
window.debugFormatState = function () {
    console.group('[書式デバッグ]');
    console.log('selectedItem:', JSON.stringify(selectedItem));
    console.log('savedSelectionNode:', savedSelectionNode ?
        (savedSelectionNode.id || savedSelectionNode.className || savedSelectionNode.tagName) : 'null');
    console.log('savedSelectionRange:', savedSelectionRange ? '設定あり' : 'null');

    if (selectedItem && selectedItem.type === 'cell') {
        const task = state.tasks.find(t => t.id === selectedItem.taskId);
        const styles = task && task.styles ? task.styles[selectedItem.field] : null;
        console.log('セルスタイル:', JSON.stringify(styles));
    } else if (selectedItem && selectedItem.type === 'text') {
        const txt = state.texts.find(t => t.id === selectedItem.textId);
        console.log('テキストボックススタイル:', txt ?
            { color: txt.color, fontWeight: txt.fontWeight, backgroundColor: txt.backgroundColor } : 'null');
    }
    console.groupEnd();
};

window.testFormatFix = function () {
    console.group('[書式修正テスト]');
    let passed = 0, failed = 0;

    // テスト1: selectInput後にsavedSelectionRangeが消去されるか
    savedSelectionRange = { collapsed: false }; // ダミーの選択範囲
    savedSelectionNode = document.getElementById('project-notes');
    window.selectInput('cell', 'test_id', 'koshu');
    if (savedSelectionRange === null && savedSelectionNode === null) {
        console.log('✓ テスト1 PASS: selectInput後にsavedSelectionRangeが消去された');
        passed++;
    } else {
        console.error('✗ テスト1 FAIL: savedSelectionRangeが消去されていない');
        failed++;
    }

    // テスト2: selectedItemがcellの場合、cell stylesに書式が保存されるか
    const testTaskId = state.tasks[0] && state.tasks[0].id;
    if (testTaskId) {
        window.selectInput('cell', testTaskId, 'koshu');
        const beforeColor = (state.tasks[0].styles && state.tasks[0].styles.koshu &&
            state.tasks[0].styles.koshu.color) || null;
        const testColor = '#ff0000';
        savedSelectionRange = null; savedSelectionNode = null;
        window.handleFormatChange('color', testColor);
        const afterColor = state.tasks[0].styles && state.tasks[0].styles.koshu &&
            state.tasks[0].styles.koshu.color;
        if (afterColor === testColor) {
            console.log('✓ テスト2 PASS: セルの文字色が正しく保存された');
            passed++;
        } else {
            console.error('✗ テスト2 FAIL: セルの文字色が保存されていない。値:', afterColor);
            failed++;
        }
        // テスト後のデータを元に戻す
        if (state.tasks[0].styles && state.tasks[0].styles.koshu) {
            state.tasks[0].styles.koshu.color = beforeColor;
        }
    } else {
        console.warn('スキップ: タスクが存在しない');
    }

    console.log(`結果: ${passed}件成功 / ${failed}件失敗`);
    console.groupEnd();
    return failed === 0;
};
