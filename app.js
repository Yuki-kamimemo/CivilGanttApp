// ---------------------------------------------------
// 定数・グローバル変数
// ---------------------------------------------------
let CELL_WIDTH_DAY = 45;
let CELL_WIDTH_MONTH = 100;

window.nationalHolidays = {};

// コピーされたタスクを保持する変数
let copiedTask = null;

async function fetchNationalHolidays() {
    try {
        const res = await fetch('https://holidays-jp.github.io/api/v1/date.json');
        if (res.ok) { 
            window.nationalHolidays = await res.json(); 
            renderAll(); 
        }
    } catch (e) { 
        console.error('祝日データの取得に失敗しました。', e); 
    }
}

// ---------------------------------------------------
// 1. データモデルと履歴管理 (State)
// ---------------------------------------------------
let state = {
    projectName: '', 
    companyName: '', 
    viewRange: 'month', 
    viewScale: 'day', 
    
    projectStart: '', 
    projectEnd: '',   
    displayStart: '', 
    displayEnd: '',   
    
    notes: '', 
    notesCollapsed: false, 
    
    dailyNoteTabs: [ 
        { id: 'tab_general', name: '作業全般・天候' }, 
        { id: 'tab_safety', name: '安全管理・行事' } 
    ],
    activeDailyNoteTab: 'tab_general', 
    dailyNotesData: { 'tab_general': {}, 'tab_safety': {} },
    
    holidays: { sundays: true, saturdays: false, nationalHolidays: true, custom: [] },
    tasks: [],
    texts: [],
    shapes: [],
    autoCreateBar: true // 自動作成のオンオフ状態
};

// 描画ツールの状態管理
let currentTool = 'pointer'; // 初期状態は「選択・操作」ツール

window.setTool = function(toolName) {
    currentTool = toolName;
    
    // 全てのメニューから「✓」を外し、選ばれたものだけに「✓」をつける
    const tools = ['pointer', 'text'];
    tools.forEach(t => {
        const el = document.getElementById('menu-tool-' + t);
        if (el) {
            el.textContent = (t === toolName ? '✓ ' : '　 ') + el.textContent.substring(2);
        }
    });
    
    // 描画ツールの時は、カレンダー上のカーソルを十字にする
    const chartArea = document.getElementById('chart-area');
    chartArea.style.cursor = (toolName === 'pointer') ? 'default' : 'crosshair';
};

// バー作成モードの管理
let barCreationMode = 'normal'; // 初期状態は「予定」

window.setBarCreationMode = function(mode) {
    barCreationMode = mode;
    
    const btnNormal = document.getElementById('mode-normal');
    const btnNew = document.getElementById('mode-new-change');
    const btnCurrent = document.getElementById('mode-current-change');
    
    if (btnNormal && btnNew && btnCurrent) {
        [btnNormal, btnNew, btnCurrent].forEach(btn => {
            btn.classList.remove('active');
            btn.style.fontWeight = 'normal';
            btn.style.borderColor = 'transparent';
            btn.style.backgroundColor = 'transparent';
        });
        
        if (mode === 'normal') {
            btnNormal.style.fontWeight = 'bold';
            btnNormal.style.borderColor = '#0d6efd';
            btnNormal.style.backgroundColor = '#cce5ff';
        } else if (mode === 'new_change') {
            btnNew.style.fontWeight = 'bold';
            btnNew.style.borderColor = '#dc3545';
            btnNew.style.backgroundColor = '#f8d7da';
        } else if (mode === 'current_change') {
            btnCurrent.style.fontWeight = 'bold';
            btnCurrent.style.borderColor = '#dc3545';
            btnCurrent.style.backgroundColor = '#f8d7da';
        }
    }
};

let selectedItem = null; 
let editModalTaskId = null;
let editModalPeriodId = null;

// 履歴管理（Undo/Redo）
let stateHistory = [];
let historyIndex = -1;
const MAX_HISTORY = 50; 

window.saveStateToHistory = function() {
    if (historyIndex < stateHistory.length - 1) {
        stateHistory = stateHistory.slice(0, historyIndex + 1);
    }
    stateHistory.push(JSON.parse(JSON.stringify(state)));
    if (stateHistory.length > MAX_HISTORY) {
        stateHistory.shift(); 
    } else {
        historyIndex++;
    }
    updateUndoRedoUI();
}

window.undo = function() { 
    if (historyIndex > 0) { 
        historyIndex--; 
        state = JSON.parse(JSON.stringify(stateHistory[historyIndex])); 
        renderAll(); 
        updateUndoRedoUI(); 
    } 
}

window.redo = function() { 
    if (historyIndex < stateHistory.length - 1) { 
        historyIndex++; 
        state = JSON.parse(JSON.stringify(stateHistory[historyIndex])); 
        renderAll(); 
        updateUndoRedoUI(); 
    } 
}

function updateUndoRedoUI() {
    const undoBtn = document.getElementById('menu-undo'); 
    const redoBtn = document.getElementById('menu-redo');
    if (undoBtn) undoBtn.classList.toggle('disabled', historyIndex <= 0);
    if (redoBtn) redoBtn.classList.toggle('disabled', historyIndex >= stateHistory.length - 1);
}

// ---------------------------------------------------
// キーボード・マウス操作のイベント設定
// ---------------------------------------------------
document.addEventListener('keydown', (e) => {
    const isInputting = e.target.tagName === 'INPUT' || e.target.tagName === 'TEXTAREA' || e.target.isContentEditable;
    
    // 保存 (Ctrl+S) は入力中でも有効
    if ((e.ctrlKey || e.metaKey) && e.key.toLowerCase() === 's') {
        e.preventDefault();
        window.handleOverwriteSave();
        return;
    }

    if (isInputting) return;
    
    // Undo / Redo
    if ((e.ctrlKey || e.metaKey) && e.key.toLowerCase() === 'z' && !e.shiftKey) { 
        e.preventDefault(); window.undo(); 
    }
    if (((e.ctrlKey || e.metaKey) && e.key.toLowerCase() === 'y') || ((e.ctrlKey || e.metaKey) && e.shiftKey && e.key.toLowerCase() === 'z')) { 
        e.preventDefault(); window.redo(); 
    }
    
    // 削除
    if (e.key === 'Delete' && selectedItem) {
        if (selectedItem.type === 'bar') window.handleContextAction('delete_period');
        else if (selectedItem.type === 'arrow') window.handleContextAction('delete_arrow');
        else if (selectedItem.type === 'task') window.handleContextAction('delete_task');
    }
});

document.addEventListener('mousedown', (e) => {
    const menu = document.getElementById('custom-context-menu');
    if (e.target.closest('.context-menu')) return;
    
    if (e.target.tagName === 'INPUT' || e.target.tagName === 'TEXTAREA' || e.target.closest('[contenteditable="true"]')) {
        if (menu) menu.style.display = 'none';
        return;
    }
    
    // ▼▼▼ ここを追加：ボタンなどをクリックした時は無駄な再描画をしない ▼▼▼
    if (e.target.closest('button') || e.target.closest('.menubar')) {
        if (menu) menu.style.display = 'none';
        return;
    }
    // ▲▲▲ ここまで ▲▲▲
    
    if (!e.target.closest('.task-bar') && !e.target.closest('path') && !e.target.closest('.chart-text-box') && !e.target.closest('#format-toolbar')) { 
        // 選択中のアイテムがある時だけ再描画するように条件を厳しくする
        if (selectedItem) {
            selectedItem = null; 
            if (window.updateFormatToolbar) window.updateFormatToolbar();
            renderChart(); 
        }
    }
    if (menu) menu.style.display = 'none';
});

// ---------------------------------------------------
// 汎用ヘルパー関数群
// ---------------------------------------------------
const generateId = () => Math.random().toString(36).substr(2, 9);
const formatDate = (dateObj) => {
    if (!dateObj) return '';
    const y = dateObj.getFullYear();
    const m = String(dateObj.getMonth() + 1).padStart(2, '0');
    const d = String(dateObj.getDate()).padStart(2, '0');
    return `${y}-${m}-${d}`;
};
const formatShortDate = (dateStr) => {
    if(!dateStr) return '';
    const parts = dateStr.split('-');
    return `${parseInt(parts[1])}/${parseInt(parts[2])}`;
};

const isHoliday = (dateObj) => {
    if (!dateObj) return false;
    const dateStr = formatDate(dateObj);
    if (state.holidays.nationalHolidays && window.nationalHolidays && window.nationalHolidays[dateStr]) return true;
    const day = dateObj.getDay();
    if (state.holidays.sundays && day === 0) return true;
    if (state.holidays.saturdays && day === 6) return true;
    if (state.holidays.custom.includes(dateStr)) return true;
    return false;
};

const snapToWorkDay = (dateStr, direction = 1) => {
    if (!dateStr) return '';
    let cur = new Date(dateStr);
    while (isHoliday(cur)) cur.setDate(cur.getDate() + direction);
    return formatDate(cur);
};

const shiftWorkDays = (dateStr, shiftCount) => {
    if (!dateStr) return '';
    let cur = new Date(dateStr);
    if (shiftCount === 0) return snapToWorkDay(dateStr, 1);
    let step = shiftCount > 0 ? 1 : -1;
    let remain = Math.abs(shiftCount);
    while (remain > 0) {
        cur.setDate(cur.getDate() + step);
        if (!isHoliday(cur)) remain--;
    }
    return formatDate(cur);
};

const getWorkDayShift = (oldStartStr, newStartStr) => {
    if (!oldStartStr || !newStartStr) return 0;
    let d1 = new Date(oldStartStr); 
    let d2 = new Date(newStartStr);
    if (d1.getTime() === d2.getTime()) return 0;
    let step = d1 < d2 ? 1 : -1;
    let count = 0; 
    let cur = new Date(d1);
    while (cur.getTime() !== d2.getTime()) {
        cur.setDate(cur.getDate() + step);
        if (!isHoliday(cur)) count += step;
    }
    return count;
};

const calcDiffDays = (startStr, endStr) => {
    if (!startStr || !endStr) return '';
    let start = new Date(startStr); 
    let end = new Date(endStr);
    if (start > end) return 0;
    let count = 0; 
    let cur = new Date(start);
    while (cur <= end) {
        if (!isHoliday(cur)) count++;
        cur.setDate(cur.getDate() + 1);
    }
    return count;
};

const calcEndDate = (startStr, daysNum) => {
    if (!startStr || !daysNum || daysNum < 1) return '';
    let cur = new Date(startStr); 
    let count = 0;
    while (true) {
        if (!isHoliday(cur)) count++;
        if (count >= daysNum) break;
        cur.setDate(cur.getDate() + 1);
    }
    return formatDate(cur);
};

function dateToPx(dateStr) {
    if (!dateStr || !state.displayStart) return 0;
    const date = new Date(dateStr); 
    const dStart = new Date(state.displayStart);
    
    if (state.viewScale === 'day') { 
        return ((date.getTime() - dStart.getTime()) / (1000 * 60 * 60 * 24)) * CELL_WIDTH_DAY; 
    } else if (state.viewScale === 'month') {
        const monthsDiff = (date.getFullYear() - dStart.getFullYear()) * 12 + (date.getMonth() - dStart.getMonth());
        const daysInMonth = new Date(date.getFullYear(), date.getMonth() + 1, 0).getDate();
        return (monthsDiff + ((date.getDate() - 1) / daysInMonth)) * CELL_WIDTH_MONTH;
    }
}

function pxToDate(px) {
    const dStart = new Date(state.displayStart);
    if (state.viewScale === 'day') { 
        dStart.setDate(dStart.getDate() + Math.round(px / CELL_WIDTH_DAY)); 
        return formatDate(dStart); 
    } else if (state.viewScale === 'month') {
        const months = px / CELL_WIDTH_MONTH; 
        const mInt = Math.floor(months); 
        const mFrac = months - mInt;
        let targetMonth = new Date(dStart.getFullYear(), dStart.getMonth() + mInt, 1);
        const daysInMonth = new Date(targetMonth.getFullYear(), targetMonth.getMonth() + 1, 0).getDate();
        targetMonth.setDate(1 + Math.round(mFrac * daysInMonth)); 
        return formatDate(targetMonth);
    }
}

// 期間の自動延長
window.checkAndExtendProjectDates = function(startStr, endStr) {
    if (!startStr && !endStr) return;
    let changed = false;

    if (startStr && (!state.projectStart || new Date(startStr) < new Date(state.projectStart))) {
        state.projectStart = startStr;
        changed = true;
    }
    if (endStr && (!state.projectEnd || new Date(endStr) > new Date(state.projectEnd))) {
        state.projectEnd = endStr;
        changed = true;
    }

    if (changed && state.viewRange === 'custom') {
        state.displayStart = state.projectStart;
        state.displayEnd = state.projectEnd;
    }
};

// 後続タスクのシフト（玉突き式に全ての後続タスクをシフトする）
window.shiftDependentTasks = function(sourcePid, diffWorkDays) {
    if (!diffWorkDays || diffWorkDays === 0) return;
    
    // 無限ループ（矢印が循環している場合など）を防ぐために、すでに動かしたIDを記憶する
    const shiftedPids = new Set();
    shiftedPids.add(sourcePid);

    // 繋がりを最後までたどって、連動して動かすための「再帰関数」
    const shiftRecursive = (currentPid) => {
        state.tasks.forEach(task => {
            task.periods.forEach(period => {
                if (period.dep) {
                    const deps = period.dep.toString().split(',').map(s => s.trim());
                    
                    // currentPid（直前に動かしたバー）に依存していて、まだ動かしていない場合
                    if (deps.includes(currentPid) && !shiftedPids.has(period.pid)) { 
                        shiftedPids.add(period.pid); // 動かしたリストに追加
                        
                        // 日付を同じ日数分ずらす
                        if (period.start) period.start = shiftWorkDays(period.start, diffWorkDays);
                        if (period.end) period.end = shiftWorkDays(period.end, diffWorkDays);
                        
                        // 連動して動いた結果、プロジェクト全体の期間をはみ出したら自動で広げる
                        checkAndExtendProjectDates(period.start, period.end);
                        
                        // ★さらにこのタスクに依存している「その次のタスク」も連動させる
                        shiftRecursive(period.pid);
                    }
                }
            });
        });
    };

    // 最初に動かしたバーを起点として、連動処理をスタート
    shiftRecursive(sourcePid);
};

window.toggleTableDetail = function() {
    const leftPane = document.getElementById('left-pane'); 
    const btn = document.getElementById('table-toggle-btn');
    leftPane.classList.toggle('collapsed-view');
    btn.textContent = leftPane.classList.contains('collapsed-view') ? '▶' : '◀';
    btn.title = leftPane.classList.contains('collapsed-view') ? '詳細列を展開する' : '詳細列を折りたたむ';
    renderChart();
};


// ---------------------------------------------------
// 2. 初期化処理・連動処理
// ---------------------------------------------------
document.addEventListener('DOMContentLoaded', () => {
    fetchNationalHolidays();
    
    const today = new Date(); 
    const nextMonth = new Date(today); 
    nextMonth.setMonth(nextMonth.getMonth() + 1);
    
    state.projectStart = formatDate(today); 
    state.projectEnd = formatDate(nextMonth);
    state.displayStart = state.projectStart;
    state.displayEnd = state.projectEnd;

    state.tasks.push({ 
        id: generateId(), no: 1, koshu: "", shubetsu: "", saibetsu: "", collapsed: false, 
        mergeAboveKoshu: false, mergeAboveShubetsu: false,
        periods: [ { pid: generateId(), dep: "", start: "", end: "", progress: 0, color: "#3b82f6", displayRow: 0 } ] 
    });
    
    document.getElementById('view-range-selector').value = state.viewRange; 
    window.handleViewRangeChange(false); 

    // スクロールと幅の同期処理
    const rightContainer = document.getElementById('right-container'); 
    const dailyNotesRight = document.getElementById('daily-notes-right');
    const leftContainer = document.getElementById('left-container');
    
    const syncScroll = (source, target, prop) => {
        if (Math.abs(target[prop] - source[prop]) > 1) {
            target[prop] = source[prop];
        }
    };

    rightContainer.addEventListener('scroll', () => { 
        syncScroll(rightContainer, dailyNotesRight, 'scrollLeft');
        syncScroll(rightContainer, leftContainer, 'scrollTop');
    });
    dailyNotesRight.addEventListener('scroll', () => syncScroll(dailyNotesRight, rightContainer, 'scrollLeft'));
    leftContainer.addEventListener('scroll', () => syncScroll(leftContainer, rightContainer, 'scrollTop'));

    // レイアウト幅の同期
    const leftPane = document.getElementById('left-pane'); 
    const dailyNotesLeft = document.getElementById('daily-notes-left');
    if (window.ResizeObserver) { 
        new ResizeObserver(entries => { 
            for (let entry of entries) { 
                dailyNotesLeft.style.width = entry.contentRect.width + 'px'; 
                dailyNotesLeft.style.minWidth = entry.contentRect.width + 'px'; 
            } 
        }).observe(leftPane); 
    }

    // モーダル内の自動計算
    document.getElementById('modal-period-start').addEventListener('change', (e) => {
        const days = parseInt(document.getElementById('modal-period-days').value);
        if(!isNaN(days) && days > 0 && e.target.value) {
            document.getElementById('modal-period-end').value = calcEndDate(snapToWorkDay(e.target.value, 1), days);
        }
    });
    document.getElementById('modal-period-days').addEventListener('input', (e) => {
        const start = document.getElementById('modal-period-start').value;
        const days = parseInt(e.target.value);
        if(start && !isNaN(days) && days > 0) {
            document.getElementById('modal-period-end').value = calcEndDate(start, days);
        }
    });
    document.getElementById('modal-period-end').addEventListener('change', (e) => {
        const start = document.getElementById('modal-period-start').value;
        if(start && e.target.value) {
            document.getElementById('modal-period-days').value = calcDiffDays(start, e.target.value);
        }
    });
    
    // カレンダー上をクリックしてテキストを追加
    const chartAreaObj = document.getElementById('chart-area');
    chartAreaObj.addEventListener('mousedown', (e) => {
        if (currentTool !== 'text') return; 
        if (e.target.closest('.task-bar') || e.target.closest('.chart-text-box') || e.target.closest('path') || e.target.closest('circle')) return;
        if (e.button !== 0) return; 
        
        const rect = chartAreaObj.getBoundingClientRect();
        const clickX = e.clientX - rect.left + chartAreaObj.scrollLeft;
        const clickY = e.clientY - rect.top + chartAreaObj.scrollTop;

        const newId = generateId();
        state.texts.push({ id: newId, text: '', x: clickX, y: clickY, width: 100, height: 30 });
        window.setTool('pointer'); 
        
        selectedItem = { type: 'text', textId: newId };
        window.saveStateToHistory();
        renderAll();
        
        setTimeout(() => {
            const newDiv = document.querySelector(`.chart-text-box[data-id="${newId}"]`);
            if (newDiv) {
                newDiv.contentEditable = "true";
                newDiv.focus();
            }
        }, 50);
    });

    window.saveStateToHistory(); 
    renderAll();

    window.addEventListener('pywebviewready', async function() {
        if (window.pywebview && window.pywebview.api) {
            const initialContent = await window.pywebview.api.get_initial_data();
            if (initialContent) {
                try {
                    applyLoadedData(JSON.parse(initialContent));
                } catch (err) {
                    console.error('初期データの読み込みに失敗しました。');
                }
            }
        }
    });
});

// ---------------------------------------------------
// モーダル・コンテキストメニュー操作群
// ---------------------------------------------------
window.openHolidayModal = function() {
    document.getElementById('modal-sun').checked = state.holidays.sundays; 
    document.getElementById('modal-sat').checked = state.holidays.saturdays; 
    document.getElementById('modal-national').checked = state.holidays.nationalHolidays;
    window.renderCustomHolidays(); 
    document.getElementById('holiday-modal').style.display = 'flex';
};
window.closeHolidayModal = function() { document.getElementById('holiday-modal').style.display = 'none'; };

window.renderCustomHolidays = function() {
    const list = document.getElementById('custom-holiday-list'); list.innerHTML = '';
    state.holidays.custom.sort().forEach(date => { 
        list.innerHTML += `<li>${date} <button type="button" onclick="window.removeCustomHoliday('${date}')">×</button></li>`; 
    });
};
window.addCustomHoliday = function() {
    const val = document.getElementById('custom-holiday-input').value;
    if (val && !state.holidays.custom.includes(val)) { 
        state.holidays.custom.push(val); 
        window.renderCustomHolidays(); 
        document.getElementById('custom-holiday-input').value = ''; 
    }
};
window.removeCustomHoliday = function(date) { 
    state.holidays.custom = state.holidays.custom.filter(d => d !== date); 
    window.renderCustomHolidays(); 
};
window.saveHolidaySettings = function() {
    state.holidays.sundays = document.getElementById('modal-sun').checked; 
    state.holidays.saturdays = document.getElementById('modal-sat').checked; 
    state.holidays.nationalHolidays = document.getElementById('modal-national').checked;
    window.closeHolidayModal(); 
    window.saveStateToHistory(); 
    renderAll();
};

window.openPeriodModal = function(taskId, periodId) {
    const task = state.tasks.find(t => t.id === taskId);
    if(!task) return;
    const period = task.periods.find(p => p.pid === periodId);
    if(!period) return;

    editModalTaskId = taskId;
    editModalPeriodId = periodId;

    document.getElementById('modal-period-start').value = period.start || '';
    document.getElementById('modal-period-end').value = period.end || '';
    document.getElementById('modal-period-days').value = calcDiffDays(period.start, period.end) || '';
    document.getElementById('modal-period-progress').value = period.progress || 0;
    document.getElementById('modal-period-color').value = period.color || '#3b82f6';

    document.getElementById('period-modal').style.display = 'flex';
};

window.closePeriodModal = function() {
    document.getElementById('period-modal').style.display = 'none';
    editModalTaskId = null;
    editModalPeriodId = null;
};

window.savePeriodModal = function() {
    if(!editModalTaskId || !editModalPeriodId) return;

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

window.updateFormatToolbar = function() {
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
        document.getElementById('toolbar-text-color').value = t.color || '#212529';
        document.getElementById('toolbar-bg-color').value = t.backgroundColor || 'transparent';
        document.getElementById('toolbar-border-style').value = t.borderStyle || 'solid';
        document.getElementById('toolbar-border-width').value = t.borderWidth !== undefined ? t.borderWidth : 1;
        document.getElementById('toolbar-border-color').value = t.borderColor || '#6c757d';

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

window.handleFormatChange = function(prop, value) {
    const selection = window.getSelection();
    let activeEl = document.activeElement;
    
    if (savedSelectionRange && savedSelectionNode && savedSelectionNode.isContentEditable) {
        if (!activeEl || !activeEl.isContentEditable) {
            savedSelectionNode.focus();
            selection.removeAllRanges();
            selection.addRange(savedSelectionRange);
            activeEl = savedSelectionNode; 
        }
    }

    if (activeEl && activeEl.isContentEditable) {
        if (prop === 'fontWeight') {
            document.execCommand('bold', false, null);
        } else if (prop === 'color') {
            document.execCommand('foreColor', false, value);
        } else if (prop === 'backgroundColor') {
            if (!document.execCommand('hiliteColor', false, value)) {
                document.execCommand('backColor', false, value);
            }
        } else if (prop === 'fontFamily') {
            document.execCommand('fontName', false, value);
        } else if (prop === 'fontSize') {
            document.execCommand('fontSize', false, '7');
            const fonts = activeEl.getElementsByTagName('font');
            for (let i = fonts.length - 1; i >= 0; i--) {
                if (fonts[i].size === '7') {
                    fonts[i].removeAttribute('size');
                    fonts[i].style.fontSize = value + 'px';
                }
            }
            const spans = activeEl.getElementsByTagName('span');
            for (let i = spans.length - 1; i >= 0; i--) {
                if (spans[i].style.fontSize === 'xxx-large' || spans[i].style.fontSize === '7px') {
                    spans[i].style.fontSize = value + 'px';
                }
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

window.selectInput = function(type, taskId = null, field = null) {
    selectedItem = { type: type, taskId: taskId, field: field };
    window.updateFormatToolbar();
};

// ---------------------------------------------------
// 独自モーダル制御
// ---------------------------------------------------
let textInputCallback = null;
let confirmCallback = null;

window.openTextInputModal = function(title, defaultValue, callback) {
    document.getElementById('text-input-title').textContent = title;
    const inputField = document.getElementById('text-input-field');
    inputField.value = defaultValue;
    textInputCallback = callback;
    document.getElementById('text-input-modal').style.display = 'flex';
    inputField.focus();
};
window.closeTextInputModal = function() {
    document.getElementById('text-input-modal').style.display = 'none';
    textInputCallback = null;
};
window.saveTextInputModal = function() {
    const val = document.getElementById('text-input-field').value;
    if (textInputCallback) textInputCallback(val);
    window.closeTextInputModal();
};

window.openConfirmModal = function(title, message, callback) {
    document.getElementById('confirm-title').textContent = title;
    document.getElementById('confirm-message').textContent = message;
    confirmCallback = callback;
    document.getElementById('confirm-modal').style.display = 'flex';
};
window.closeConfirmModal = function() {
    document.getElementById('confirm-modal').style.display = 'none';
    confirmCallback = null;
};
window.executeConfirmModal = function() {
    if (confirmCallback) confirmCallback();
    window.closeConfirmModal();
};

// ---------------------------------------------------
// 印刷制御
// ---------------------------------------------------
let prePrintState = null;

window.openPrintModal = function() {
    document.getElementById('modal-print-start').value = state.displayStart;
    document.getElementById('modal-print-end').value = state.displayEnd;
    document.getElementById('modal-print-zoom').value = state.zoomRatio || 1.0;
    document.getElementById('print-modal').style.display = 'flex';
};

window.closePrintModal = function() {
    document.getElementById('print-modal').style.display = 'none';
};

window.executePrint = function() {
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
window.showContextMenu = function(e, title, type, data) { 
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

window.handleContextAction = function(action) {
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
        window.openConfirmModal('行の削除', 'この行（タスク）を完全に削除しますか？', function() {
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
            window.openTextInputModal('テキストの文字を編集', txt.text, function(val) {
                if (val) { txt.text = val; window.saveStateToHistory(); renderAll(); }
            });
        }
    } else if (action === 'delete_text') {
        const targetTextId = selectedItem.textId; 
        window.openConfirmModal('削除', 'このテキストを削除しますか？', function() {
            state.texts = state.texts.filter(t => t.id !== targetTextId);
            window.saveStateToHistory(); renderAll();
        });
    }

    selectedItem = null;
    document.getElementById('custom-context-menu').style.display = 'none';
};

window.handleRowContextMenu = function(e, taskId, taskNo, koshu) {
    if (e.target.tagName === 'INPUT' && e.target.type !== 'color') return;
    const title = `行アクション (No.${taskNo} ${koshu || '名称未設定'})`;
    window.showContextMenu(e, title, 'task', { taskId });
};


// ---------------------------------------------------
// 3. データ操作アクション
// ---------------------------------------------------
window.handleNewFile = function() {
    window.openConfirmModal('新規作成', '現在のデータは破棄されます。保存していない変更は失われますが、よろしいですか？', async function() {
        if (window.pywebview && window.pywebview.api) {
            await window.pywebview.api.clear_file_path();
        }

        const today = new Date(); 
        const nextMonth = new Date(today); 
        nextMonth.setMonth(nextMonth.getMonth() + 1);

        state = {
            projectName: '', companyName: '', viewRange: 'month', viewScale: 'day',
            projectStart: formatDate(today), projectEnd: formatDate(nextMonth),   
            displayStart: formatDate(today), displayEnd: formatDate(nextMonth),   
            notes: '', notesCollapsed: false, 
            dailyNoteTabs: [ { id: 'tab_general', name: '作業全般・天候' }, { id: 'tab_safety', name: '安全管理・行事' } ],
            activeDailyNoteTab: 'tab_general', 
            dailyNotesData: { 'tab_general': {}, 'tab_safety': {} },
            holidays: { sundays: true, saturdays: false, nationalHolidays: true, custom: [] },
            autoCreateBar: true, 
            tasks: [
                { 
                    id: generateId(), no: 1, koshu: "", shubetsu: "", saibetsu: "", collapsed: false, 
                    mergeAboveKoshu: false, mergeAboveShubetsu: false,
                    periods: [ { pid: generateId(), dep: "", start: "", end: "", progress: 0, color: "#3b82f6", displayRow: 0 } ] 
                }
            ],
            texts: [], shapes: [],
            globalFontFamily: "'Segoe UI', Tahoma, Geneva, Verdana, sans-serif",
            globalFontSize: 13
        };

        stateHistory = []; historyIndex = -1;
        window.saveStateToHistory(); renderAll();
    });
};

window.handleOverwriteSave = async function() {
    const dataStr = JSON.stringify(state, null, 2); 
    if (window.pywebview && window.pywebview.api) {
        const success = await window.pywebview.api.overwrite_file(dataStr);
        if (success) alert('上書き保存しました！');
        else window.handleSaveFile();
    } else {
        window.handleSaveFile();
    }
};

window.handleSaveFile = async function() {
    const dataStr = JSON.stringify(state, null, 2); 
    const fileName = state.projectName ? `${state.projectName}_工程表.csm` : '工程表.csm';
    
    if (window.pywebview && window.pywebview.api) {
        await window.pywebview.api.save_file(dataStr, fileName);
    } else {
        const blob = new Blob([dataStr], { type: "application/json" }); 
        const url = URL.createObjectURL(blob); 
        const a = document.createElement('a'); 
        a.href = url; a.download = fileName; 
        document.body.appendChild(a); a.click(); 
        document.body.removeChild(a); URL.revokeObjectURL(url);
    }
};

window.handleOpenFile = async function() {
    if (window.pywebview && window.pywebview.api) {
        const fileContent = await window.pywebview.api.open_file();
        if (fileContent) {
            try { applyLoadedData(JSON.parse(fileContent)); } 
            catch (err) { alert('読み込みに失敗しました。ファイルが壊れている可能性があります。'); }
        }
    } else {
        const input = document.createElement('input'); 
        input.type = 'file'; input.accept = '.csm';
        input.onchange = (e) => {
            const file = e.target.files[0]; 
            if (!file) return;
            const reader = new FileReader();
            reader.onload = (event) => {
                try { applyLoadedData(JSON.parse(event.target.result)); } 
                catch (err) { alert('読み込みに失敗しました。'); }
            };
            reader.readAsText(file);
        };
        input.click();
    }
};

window.handleExportExcel = async function() {
    const exportData = { state: state, nationalHolidays: window.nationalHolidays || {} };
    const dataStr = JSON.stringify(exportData); 
    const fileName = state.projectName ? `${state.projectName}_工程表.xlsx` : '工程表.xlsx'; 
    
    if (window.pywebview && window.pywebview.api) {
        const success = await window.pywebview.api.export_to_excel(dataStr, fileName);
        if (success) alert('Excelファイルとしてエクスポートが完了しました！');
    } else {
        alert('この機能はデスクトップアプリ版でのみ利用可能です。');
    }
};

function applyLoadedData(parsedData) {
    if (parsedData && Array.isArray(parsedData.tasks)) {
        if (!parsedData.holidays) parsedData.holidays = { sundays: true, saturdays: false, nationalHolidays: true, custom: [] };
        if (!parsedData.viewRange) parsedData.viewRange = 'custom';
        if (!parsedData.viewScale) parsedData.viewScale = 'day';
        if (!parsedData.dailyNoteTabs) { 
            parsedData.dailyNoteTabs = [{ id: 'tab_general', name: '作業全般・天候' }]; 
            parsedData.activeDailyNoteTab = 'tab_general'; 
            parsedData.dailyNotesData = { 'tab_general': {} }; 
        }
        if (parsedData.autoCreateBar === undefined) parsedData.autoCreateBar = true;
        
        if (parsedData.globalStart) { parsedData.projectStart = parsedData.globalStart; parsedData.displayStart = parsedData.globalStart; delete parsedData.globalStart; }
        if (parsedData.globalEnd) { parsedData.projectEnd = parsedData.globalEnd; parsedData.displayEnd = parsedData.globalEnd; delete parsedData.globalEnd; }

        parsedData.tasks.forEach(t => {
            if (t.name !== undefined) { t.koshu = t.name; delete t.name; }
            if (t.shubetsu === undefined) t.shubetsu = ""; if (t.saibetsu === undefined) t.saibetsu = "";
            if (t.mergeAboveKoshu === undefined) t.mergeAboveKoshu = false;
            if (t.mergeAboveShubetsu === undefined) t.mergeAboveShubetsu = false;
            if(t.periods) t.periods.forEach(p => { 
                if (p.progress === undefined) p.progress = 0; 
                if (p.color === undefined) p.color = '#3b82f6'; 
                if (p.displayRow === undefined) p.displayRow = 0; 
            });
        });
        state = parsedData; 
        window.saveStateToHistory(); renderAll(); 
    } else { 
        alert('無効なファイルです。'); 
    }
}

window.handleProjectInfoChange = function() { 
    state.projectName = document.getElementById('project-name').value; 
    state.companyName = document.getElementById('company-name').value; 
    window.saveStateToHistory(); 
};

window.handleProjectDateChange = function() {
    state.projectStart = document.getElementById('project-start').value;
    state.projectEnd = document.getElementById('project-end').value;
    if (state.viewRange === 'custom') {
        state.displayStart = state.projectStart;
        state.displayEnd = state.projectEnd;
    }
    window.saveStateToHistory(); renderAll();
};

window.handleViewScaleChange = function() {
    state.viewScale = document.getElementById('view-scale-selector').value;
    if (state.viewScale === 'month' && state.displayStart && state.displayEnd) {
        const s = new Date(state.displayStart); s.setDate(1); state.displayStart = formatDate(s); 
        const e = new Date(state.displayEnd); e.setMonth(e.getMonth() + 1); e.setDate(0); state.displayEnd = formatDate(e);
    }
    window.saveStateToHistory(); renderAll();
};

window.handleViewRangeChange = function(shouldRender = true) {
    state.viewRange = document.getElementById('view-range-selector').value;
    if (!state.projectStart) state.projectStart = formatDate(new Date());
    if (!state.displayStart) state.displayStart = state.projectStart;

    if (state.viewRange === 'custom') {
        state.displayStart = state.projectStart;
        state.displayEnd = state.projectEnd;
    } else {
        state.displayEnd = calcDisplayEndFromStr(state.displayStart, state.viewRange);
        if (state.viewRange === 'week' || state.viewRange === 'month') state.viewScale = 'day';
        else state.viewScale = 'month';
    }

    if (shouldRender) { window.saveStateToHistory(); renderAll(); }
};

window.handleDisplayStartChange = function() {
    const newStart = document.getElementById('display-start-date').value;
    if (newStart) {
        state.displayStart = newStart;
        if (state.viewRange !== 'custom') {
            state.displayEnd = calcDisplayEndFromStr(state.displayStart, state.viewRange);
        }
        window.saveStateToHistory(); renderAll();
    }
};

function calcDisplayEndFromStr(startStr, range) {
    if (!startStr) return '';
    let d = new Date(startStr);
    if (range === 'week') d.setDate(d.getDate() + 7);
    else if (range === 'month') d.setMonth(d.getMonth() + 1);
    else if (range === 'half-year') d.setMonth(d.getMonth() + 6);
    else if (range === 'year') d.setFullYear(d.getFullYear() + 1);
    return formatDate(d);
}

window.shiftDisplay = function(direction) {
    if (state.viewRange === 'custom') return; 
    let d = new Date(state.displayStart);
    if (state.viewRange === 'week') d.setDate(d.getDate() + (7 * direction));
    else if (state.viewRange === 'month') d.setMonth(d.getMonth() + (1 * direction));
    else if (state.viewRange === 'half-year') d.setMonth(d.getMonth() + (6 * direction));
    else if (state.viewRange === 'year') d.setFullYear(d.getFullYear() + (1 * direction));

    state.displayStart = formatDate(d);
    state.displayEnd = calcDisplayEndFromStr(state.displayStart, state.viewRange);
    window.saveStateToHistory(); renderAll();
};

window.handleAddTask = function() { 
    const newNo = state.tasks.length + 1; 
    state.tasks.push({ 
        id: generateId(), no: newNo, koshu: "", shubetsu: "", saibetsu: "", collapsed: false, 
        mergeAboveKoshu: false, mergeAboveShubetsu: false,
        periods: [ { pid: generateId(), dep: "", start: "", end: "", progress: 0, color: "#3b82f6", displayRow: 0 } ] 
    }); 
    window.saveStateToHistory(); 
    renderAll(); 

    // 追加後に一番下まで自動スクロールさせる処理
    setTimeout(() => {
        const leftContainer = document.getElementById('left-container');
        const rightContainer = document.getElementById('right-container');
        if (leftContainer) {
            leftContainer.scrollTop = leftContainer.scrollHeight;
        }
        if (rightContainer) {
            rightContainer.scrollTop = rightContainer.scrollHeight;
        }
    }, 50);
};

window.handleAddPeriod = function(taskId, actionType = 'normal') { 
    const task = state.tasks.find(t => t.id === taskId); 
    if (task) { 
        const maxRow = task.periods.reduce((max, p) => Math.max(max, p.displayRow || 0), 0);
        let targetRow = 0; let color = "#3b82f6";
        if (actionType === 'new_change') { targetRow = maxRow + 1; color = "#dc3545"; } 
        else if (actionType === 'current_change') { targetRow = maxRow === 0 ? 1 : maxRow; color = "#dc3545"; }

        task.periods.push({ pid: generateId(), dep: "", start: "", end: "", progress: 0, color: color, displayRow: targetRow }); 
        task.periods.sort((a, b) => (a.displayRow || 0) - (b.displayRow || 0)); 
        task.collapsed = false; 
        window.saveStateToHistory(); renderAll(); 
    } 
};

window.handleRemovePeriod = function(taskId, periodId) { 
    const task = state.tasks.find(t => t.id === taskId); 
    if (task && task.periods.length > 1) { 
        task.periods = task.periods.filter(p => p.pid !== periodId); 
        if (task.periods.length === 1) task.collapsed = false; 
        window.saveStateToHistory(); renderAll(); 
    } 
};

window.handleToggleCollapse = function(taskId) {
    const task = state.tasks.find(t => t.id === taskId);
    if (task) {
        task.collapsed = !task.collapsed; 
        window.saveStateToHistory(); renderAll();
    }
};

window.handleTaskDetailChange = function(taskId, field, value) { 
    const targetIndex = state.tasks.findIndex(t => t.id === taskId);
    if (targetIndex === -1) return;
    
    state.tasks[targetIndex][field] = value;
    
    if (field === 'koshu' || field === 'shubetsu') {
        for (let i = targetIndex + 1; i < state.tasks.length; i++) {
            if (field === 'koshu' && state.tasks[i].mergeAboveKoshu) state.tasks[i].koshu = value;
            else if (field === 'shubetsu' && state.tasks[i].mergeAboveShubetsu) state.tasks[i].shubetsu = value;
            else break; 
        }
    }
    window.saveStateToHistory(); renderAll(); 
};

window.handlePeriodChange = function(taskId, periodId, field, value) {
    const task = state.tasks.find(t => t.id === taskId); if (!task) return;
    const period = task.periods.find(p => p.pid === periodId); if (!period) return;
    
    let diffWorkDays = 0; 
    const oldStart = period.start; 
    const prevDuration = calcDiffDays(period.start, period.end);

    if (field === 'start') {
        const snappedVal = snapToWorkDay(value, 1); 
        period.start = snappedVal;
        if (oldStart && snappedVal) diffWorkDays = getWorkDayShift(oldStart, snappedVal);
        if (prevDuration > 0 && snappedVal) period.end = calcEndDate(snappedVal, prevDuration);
    } else if (field === 'end') { 
        period.end = snapToWorkDay(value, -1);
    } else if (field === 'days') { 
        period.end = calcEndDate(period.start, value);
    } else if (field === 'dep') { 
        period.dep = value;
    } else if (field === 'progress') {
        let p = parseInt(value, 10); 
        if (isNaN(p) || p < 0) p = 0; 
        if (p > 100) p = 100; 
        period.progress = p;
    } else if (field === 'color') { 
        period.color = value; 
    }

    if (field === 'start' || field === 'end' || field === 'days') {
        checkAndExtendProjectDates(period.start, period.end);
    }
    
    if (diffWorkDays !== 0) shiftDependentTasks(periodId, diffWorkDays);
    
    window.saveStateToHistory(); renderAll();
};

window.addDailyTab = function() { 
    window.openTextInputModal('新しい備考タブの名前', '新規タブ', function(name) {
        if (name) { 
            const id = 'tab_' + generateId(); 
            state.dailyNoteTabs.push({ id, name }); 
            state.dailyNotesData[id] = {}; 
            state.activeDailyNoteTab = id; 
            window.saveStateToHistory(); renderAll(); 
        } 
    });
};
window.editCurrentDailyTab = function() { 
    const currentTab = state.dailyNoteTabs.find(t => t.id === state.activeDailyNoteTab); 
    if (!currentTab) return; 
    window.openTextInputModal('タブの名前を変更', currentTab.name, function(newName) {
        if (newName) { currentTab.name = newName; window.saveStateToHistory(); renderAll(); } 
    });
};
window.deleteCurrentDailyTab = function() {
    if (state.dailyNoteTabs.length <= 1) { alert('最後の1つのタブは削除できません。'); return; }
    window.openConfirmModal('削除の確認', '現在の備考タブと入力データを完全に削除しますか？', function() {
        const id = state.activeDailyNoteTab; 
        state.dailyNoteTabs = state.dailyNoteTabs.filter(t => t.id !== id); 
        delete state.dailyNotesData[id];
        state.activeDailyNoteTab = state.dailyNoteTabs[0].id; 
        window.saveStateToHistory(); renderAll();
    });
};
window.handleDailyNoteChange = function(dateStr, value) {
    const tabId = state.activeDailyNoteTab;
    if (!state.dailyNotesData[tabId]) state.dailyNotesData[tabId] = {};
    state.dailyNotesData[tabId][dateStr] = value; 
    window.saveStateToHistory();
};

window.handleNotesChange = function() { 
    state.notes = document.getElementById('project-notes').innerHTML; 
    window.saveStateToHistory(); 
};

window.toggleNotes = function() { 
    state.notesCollapsed = !state.notesCollapsed; 
    window.saveStateToHistory(); renderAll(); 
};

window.handleZoomChange = function() {
    const slider = document.getElementById('zoom-slider');
    state.zoomRatio = parseFloat(slider.value); 
    renderAll(); 
};

window.handleAutoCreateBarChange = function() {
    const cb = document.getElementById('auto-create-bar-cb');
    state.autoCreateBar = cb.checked;
    window.saveStateToHistory();
};

// ---------------------------------------------------
// 4. 描画処理（View）
// ---------------------------------------------------
function renderAll() {
    document.documentElement.style.setProperty('--global-font-family', state.globalFontFamily || "'Segoe UI', Tahoma, Geneva, Verdana, sans-serif");
    document.documentElement.style.setProperty('--global-font-size', (state.globalFontSize || 13) + 'px');

    state.zoomRatio = state.zoomRatio || 1.0;
    CELL_WIDTH_DAY = 45 * state.zoomRatio;
    CELL_WIDTH_MONTH = 100 * state.zoomRatio;
    
    const zoomSlider = document.getElementById('zoom-slider');
    if (zoomSlider && zoomSlider.value != state.zoomRatio) zoomSlider.value = state.zoomRatio;

    const autoCreateCb = document.getElementById('auto-create-bar-cb');
    if (autoCreateCb) autoCreateCb.checked = state.autoCreateBar;

    document.getElementById('project-name').value = state.projectName; 
    document.getElementById('company-name').value = state.companyName; 
    document.getElementById('project-start').value = state.projectStart || ''; 
    document.getElementById('project-end').value = state.projectEnd || ''; 
    
    if (document.getElementById('display-start-date')) {
        document.getElementById('display-start-date').value = state.displayStart || '';
    }

    document.getElementById('view-range-selector').value = state.viewRange || 'custom'; 
    document.getElementById('view-scale-selector').value = state.viewScale || 'day';
    
    const shiftControls = document.getElementById('display-shift-controls');
    shiftControls.style.display = (state.viewRange === 'custom') ? 'none' : 'flex';

    // 備考欄の適用
    const notesArea = document.getElementById('project-notes');
    if (document.activeElement !== notesArea) notesArea.innerHTML = state.notes || '';
    
    const ns = state.notesStyle || {};
    notesArea.style.fontFamily = ns.fontFamily || '';
    notesArea.style.fontSize = ns.fontSize ? ns.fontSize + 'px' : '';
    notesArea.style.color = ns.color || '';
    notesArea.style.fontWeight = ns.fontWeight || '';
    notesArea.style.backgroundColor = ns.backgroundColor || '';
    notesArea.style.writingMode = ns.writingMode || 'horizontal-tb';
    notesArea.style.display = 'flex';
    notesArea.style.flexDirection = 'column';
    
    const hAlignMap = { 'left': 'flex-start', 'center': 'center', 'right': 'flex-end' };
    const vAlignMap = { 'flex-start': 'flex-start', 'center': 'center', 'flex-end': 'flex-end' };
    if (notesArea.style.writingMode === 'vertical-rl') {
        notesArea.style.justifyContent = hAlignMap[ns.textAlign] || 'flex-start';
        notesArea.style.alignItems = vAlignMap[ns.verticalAlign] || 'flex-start';
    } else {
        notesArea.style.justifyContent = vAlignMap[ns.verticalAlign] || 'flex-start';
        notesArea.style.alignItems = hAlignMap[ns.textAlign] || 'flex-start';
    }
    
    const notesPane = document.getElementById('notes-pane'); 
    const notesToggleBtn = document.getElementById('notes-toggle-btn');
    if (state.notesCollapsed) { 
        notesPane.classList.add('collapsed'); 
        notesToggleBtn.textContent = '◀'; notesToggleBtn.title = '備考欄を展開する'; 
    } else { 
        notesPane.classList.remove('collapsed'); 
        notesToggleBtn.textContent = '▶'; notesToggleBtn.title = '備考欄を折りたたむ'; 
    }
    
    renderCalendarHeader(); 
    renderTable(); 
    renderChart(); 
    renderDailyNotes(); 
}

function renderDailyNotes() {
    const tabContainer = document.getElementById('daily-tabs-container'); 
    tabContainer.innerHTML = '';
    
    state.dailyNoteTabs.forEach(tab => {
        const btn = document.createElement('button'); 
        btn.className = `daily-tab-btn ${tab.id === state.activeDailyNoteTab ? 'active' : ''}`; 
        btn.textContent = tab.name;
        btn.onclick = () => { state.activeDailyNoteTab = tab.id; renderAll(); }; 
        tabContainer.appendChild(btn);
    });
    
    const currentTab = state.dailyNoteTabs.find(t => t.id === state.activeDailyNoteTab); 
    document.getElementById('current-daily-tab-name').textContent = currentTab ? currentTab.name : '';
    
    const grid = document.getElementById('daily-notes-grid'); 
    grid.innerHTML = '';
    
    if (!state.displayStart || !state.displayEnd) return;
    const tabData = state.dailyNotesData[state.activeDailyNoteTab] || {};
    
    let currentDate = new Date(state.displayStart); 
    const endDate = new Date(state.displayEnd); 
    let totalWidth = 0;

    const hAlignMap = { 'left': 'flex-start', 'center': 'center', 'right': 'flex-end' };
    const vAlignMap = { 'flex-start': 'flex-start', 'center': 'center', 'flex-end': 'flex-end' };

    const drawCell = (dateStr, dateObj, widthPx) => {
        const cell = document.createElement('div'); 
        cell.className = 'daily-note-cell'; 
        cell.style.width = widthPx + 'px'; 
        cell.style.minWidth = widthPx + 'px';
        
        const textarea = document.createElement('div'); 
        textarea.contentEditable = "true";
        textarea.style.cssText = "width:100%; height:100%; box-sizing:border-box; padding:5px; outline:none; overflow-y:auto; display:flex; flex-direction:column;";
        
        const ds = state.dailyNotesStyle || {};
        textarea.style.fontFamily = ds.fontFamily || '';
        textarea.style.fontSize = ds.fontSize ? ds.fontSize + 'px' : '';
        textarea.style.color = ds.color || '';
        textarea.style.fontWeight = ds.fontWeight || '';
        
        const currentWM = ds.writingMode || 'vertical-rl';
        const currentAlign = ds.textAlign || 'center';
        const currentVAlign = ds.verticalAlign || 'flex-start';
        textarea.style.writingMode = currentWM;
        
        if (currentWM === 'vertical-rl') {
            textarea.style.justifyContent = hAlignMap[currentAlign] || 'center';
            textarea.style.alignItems = vAlignMap[currentVAlign] || 'flex-start';
        } else {
            textarea.style.justifyContent = vAlignMap[currentVAlign] || 'flex-start';
            textarea.style.alignItems = hAlignMap[currentAlign] || 'center';
        }
        
        if (ds.backgroundColor) {
            textarea.style.backgroundColor = ds.backgroundColor;
        } else if (state.viewScale === 'day') {
            const isNatHoliday = window.nationalHolidays && window.nationalHolidays[dateStr];
            if (isNatHoliday) textarea.style.backgroundColor = 'rgba(220,53,69,0.05)'; 
            else if (isHoliday(dateObj)) textarea.style.backgroundColor = 'rgba(0,0,0,0.03)';
            else textarea.style.backgroundColor = 'transparent';
        }
        
        textarea.innerHTML = tabData[dateStr] || ''; 
        textarea.onblur = (e) => window.handleDailyNoteChange(dateStr, e.target.innerHTML);
        textarea.onfocus = () => window.selectInput('daily_notes');
        
        cell.appendChild(textarea); 
        grid.appendChild(cell); 
        totalWidth += widthPx;
    };

    if (state.viewScale === 'day') {
        while (currentDate <= endDate) {
            drawCell(formatDate(currentDate), currentDate, CELL_WIDTH_DAY);
            currentDate.setDate(currentDate.getDate() + 1); 
        }
    } else if (state.viewScale === 'month') {
        let curMonth = new Date(currentDate.getFullYear(), currentDate.getMonth(), 1); 
        const endMonth = new Date(endDate.getFullYear(), endDate.getMonth(), 1); 
        while (curMonth <= endMonth) {
            const monthStr = curMonth.getFullYear() + '-' + String(curMonth.getMonth() + 1).padStart(2, '0'); 
            drawCell(monthStr, curMonth, CELL_WIDTH_MONTH);
            curMonth.setMonth(curMonth.getMonth() + 1); 
        }
    }
    grid.style.width = totalWidth + 'px';
}

function renderCalendarHeader() {
    const headerContainer = document.getElementById('calendar-header'); 
    const chartArea = document.getElementById('chart-area'); 
    const arrowLayer = document.getElementById('arrow-layer'); 
    headerContainer.innerHTML = '';
    
    if (!state.displayStart || !state.displayEnd) return;
    
    let currentDate = new Date(state.displayStart); 
    const endDate = new Date(state.displayEnd); 
    let totalWidth = 0;

    if (state.viewScale === 'day') {
        let daysCount = 0;
        while (currentDate <= endDate) {
            const dayDiv = document.createElement('div'); 
            dayDiv.className = 'day-cell'; 
            dayDiv.style.width = CELL_WIDTH_DAY + 'px';
            
            const dateStr = formatDate(currentDate); 
            const isNatHoliday = window.nationalHolidays && window.nationalHolidays[dateStr];
            
            if (isNatHoliday) { 
                dayDiv.classList.add('national-holiday-cell'); 
                dayDiv.title = window.nationalHolidays[dateStr]; 
            } else if (isHoliday(currentDate)) { 
                dayDiv.classList.add('holiday-cell'); 
            } else { 
                const dayOfWeek = currentDate.getDay(); 
                if (dayOfWeek === 6) dayDiv.classList.add('weekend-sat'); 
                else if (dayOfWeek === 0) dayDiv.classList.add('weekend-sun'); 
            }
            
            dayDiv.innerHTML = `${currentDate.getMonth() + 1}/${currentDate.getDate()}`; 
            headerContainer.appendChild(dayDiv); 
            currentDate.setDate(currentDate.getDate() + 1); 
            daysCount++;
        }
        totalWidth = daysCount * CELL_WIDTH_DAY; 
        chartArea.style.backgroundImage = `repeating-linear-gradient(to right, transparent, transparent ${CELL_WIDTH_DAY - 1}px, #dee2e6 ${CELL_WIDTH_DAY - 1}px, #dee2e6 ${CELL_WIDTH_DAY}px)`;
    
    } else if (state.viewScale === 'month') {
        let curMonth = new Date(currentDate.getFullYear(), currentDate.getMonth(), 1); 
        const endMonth = new Date(endDate.getFullYear(), endDate.getMonth(), 1); 
        let monthsCount = 0;
        
        while (curMonth <= endMonth) {
            const monthDiv = document.createElement('div'); 
            monthDiv.className = 'day-cell'; 
            monthDiv.style.width = CELL_WIDTH_MONTH + 'px'; 
            monthDiv.style.fontWeight = 'bold'; 
            monthDiv.innerHTML = `${curMonth.getFullYear()}年${curMonth.getMonth() + 1}月`;
            
            headerContainer.appendChild(monthDiv); 
            curMonth.setMonth(curMonth.getMonth() + 1); 
            monthsCount++;
        }
        totalWidth = monthsCount * CELL_WIDTH_MONTH; 
        chartArea.style.backgroundImage = `repeating-linear-gradient(to right, transparent, transparent ${CELL_WIDTH_MONTH - 1}px, #dee2e6 ${CELL_WIDTH_MONTH - 1}px, #dee2e6 ${CELL_WIDTH_MONTH}px)`;
    }
    
    chartArea.style.width = totalWidth + 'px'; 
    arrowLayer.setAttribute('width', totalWidth);
}

// ---------------------------------------------------
// 【リファクタリング】 テーブル描画（画面左側）
// ---------------------------------------------------
function renderTable() {
    const tbody = document.getElementById('task-tbody'); 
    tbody.innerHTML = ''; 

    // セル結合用の行数を計算
    let koshuRowspans = new Array(state.tasks.length).fill(1);
    let shubetsuRowspans = new Array(state.tasks.length).fill(1);
    for (let i = state.tasks.length - 1; i > 0; i--) {
        if (state.tasks[i].mergeAboveKoshu) {
            koshuRowspans[i-1] += koshuRowspans[i];
            koshuRowspans[i] = 0;
        }
        if (state.tasks[i].mergeAboveShubetsu) {
            shubetsuRowspans[i-1] += shubetsuRowspans[i];
            shubetsuRowspans[i] = 0;
        }
    }

    state.tasks.forEach((task, index) => {
        task.no = index + 1;
        const tr = document.createElement('tr'); 
        tr.className = `task-row ${task.collapsed ? 'collapsed' : ''}`; 
        tr.dataset.taskId = task.id; 
        tr.setAttribute('oncontextmenu', `window.handleRowContextMenu(event, '${task.id}', '${task.no}', '${task.koshu}')`);

        const maxRow = task.periods.reduce((max, p) => Math.max(max, p.displayRow || 0), 0);
        const numRows = maxRow + 1;
        const requiredHeight = Math.max(40, (18 * numRows) + (10 * maxRow) + 16);

        let html = `
            <td class="task-no-cell" style="text-align: center; vertical-align: middle; height: ${requiredHeight}px;">
                <span class="task-no">${task.no}</span>
            </td>
        `;

        // 書式スタイルの組み立て
        const st = task.styles || {};
        const createCss = (obj) => `color:${obj?.color||''}; font-weight:${obj?.fontWeight||''}; font-size:${obj?.fontSize?obj.fontSize+'px':''}; font-family:${obj?.fontFamily||''}; background-color:${obj?.backgroundColor||''}; text-align:${obj?.textAlign||''};`;
        const kCss = createCss(st.koshu);
        const shCss = createCss(st.shubetsu);
        const saCss = createCss(st.saibetsu);

        if (koshuRowspans[index] > 0) {
            html += `
            <td class="task-koshu-cell" rowspan="${koshuRowspans[index]}">
                <div class="task-name-header">
                    <input type="text" class="task-koshu" placeholder="工種" value="${task.koshu}" style="${kCss}" 
                           onfocus="window.selectInput('cell', '${task.id}', 'koshu')" 
                           onchange="window.handleTaskDetailChange('${task.id}', 'koshu', this.value)">
                </div>
            </td>`;
        }

        if (shubetsuRowspans[index] > 0) {
            html += `
            <td class="task-shubetsu-cell" rowspan="${shubetsuRowspans[index]}">
                <input type="text" class="task-shubetsu" placeholder="種別" value="${task.shubetsu}" style="${shCss}" 
                       onfocus="window.selectInput('cell', '${task.id}', 'shubetsu')" 
                       onchange="window.handleTaskDetailChange('${task.id}', 'shubetsu', this.value)">
            </td>`;
        }

        html += `
            <td class="task-saibetsu-cell">
                <div class="task-name-container">
                    <input type="text" class="task-saibetsu" placeholder="細別・規格" value="${task.saibetsu}" style="${saCss}" 
                           onfocus="window.selectInput('cell', '${task.id}', 'saibetsu')" 
                           onchange="window.handleTaskDetailChange('${task.id}', 'saibetsu', this.value)">
                </div>
            </td>
        `;

        ['start', 'end', 'days', 'progress', 'action'].forEach(col => {
            const isAction = col === 'action';
            html += `<td class="task-${col}-cell ${isAction ? 'print-hide' : ''}" style="vertical-align: middle; padding: 2px 4px; text-align: center;">
                        <div style="display: flex; flex-direction: column; gap: 10px; padding: 6px 0;">`;
            
            for (let r = 0; r <= maxRow; r++) {
                const periodsInRow = task.periods.filter(p => (p.displayRow || 0) === r);
                if (periodsInRow.length === 0) continue;
                
                const agg = aggregatePeriods(periodsInRow);
                if (agg) {
                    const isChange = r > 0;
                    const textColor = isChange ? '#dc3545' : 'inherit';
                    const emptyColor = isChange ? '#f8d7da' : '#adb5bd';
                    const titleText = isChange ? `クリックして第${r}回変更を編集` : `クリックして予定を編集`;
                    const defaultColor = isChange ? '#dc3545' : '#3b82f6';

                    html += `<div class="clickable-text" onclick="window.openPeriodModal('${task.id}', '${agg.pid}')" title="${titleText}" style="height: 18px; padding: 0; color: ${textColor}; display: flex; align-items: center; justify-content: center;">`;
                    
                    if (col === 'start') html += agg.minStart ? formatShortDate(agg.minStart) : `<span class="empty-text" style="color:${emptyColor};">未設定</span>`;
                    else if (col === 'end') html += agg.maxEnd ? formatShortDate(agg.maxEnd) : `<span class="empty-text" style="color:${emptyColor};">未設定</span>`;
                    else if (col === 'days') html += agg.overallDays ? agg.overallDays + '日' : '-';
                    else if (col === 'progress') html += agg.overallProgress + '%';
                    else if (col === 'action') {
                        const c = agg.color || defaultColor;
                        html += `<div style="width: 14px; height: 14px; background-color: ${c}; border: 1px solid #ced4da; border-radius: 3px;"></div>`;
                    }
                    html += `</div>`;
                }
            }
            html += `</div></td>`;
        });
        
        tr.innerHTML = html; 
        tbody.appendChild(tr);
    });
}

// 複数の期間データをまとめるための便利関数
function aggregatePeriods(periods) {
    if (periods.length === 0) return null;
    let minStart = null, maxEnd = null, totalWorkDays = 0, weightedProgress = 0;
    
    periods.forEach(p => {
        if (p.start && (!minStart || new Date(p.start) < new Date(minStart))) minStart = p.start;
        if (p.end && (!maxEnd || new Date(p.end) > new Date(maxEnd))) maxEnd = p.end;
        if (p.start && p.end) {
            const d = calcDiffDays(p.start, p.end) || 0;
            totalWorkDays += d;
            weightedProgress += d * (p.progress || 0);
        }
    });
    const overallProgress = totalWorkDays > 0 ? Math.round(weightedProgress / totalWorkDays) : (periods[0].progress || 0);
    const overallDays = calcDiffDays(minStart, maxEnd);
    return { minStart, maxEnd, overallDays, overallProgress, color: periods[0].color, pid: periods[0].pid };
}

// ---------------------------------------------------
// 【リファクタリング】 カレンダー描画（画面右側）
// ---------------------------------------------------
function renderChart() {
    const chartArea = document.getElementById('chart-area'); 
    const arrowLayer = document.getElementById('arrow-layer');
    const dStart = new Date(state.displayStart); 
    const dEnd = new Date(state.displayEnd);
    
    // 1. 描画エリアの初期化
    Array.from(chartArea.children).forEach(child => { 
        if (child.id !== 'arrow-layer') child.remove(); 
    }); 
    arrowLayer.innerHTML = '';

    // 2. 背景（休日）の描画
    drawChartBackground(chartArea, dStart, dEnd);

    // 3. 各タスクのバーの描画
    const calculatedBars = []; 
    let currentYOffset = 0; 
    const tableRows = document.querySelectorAll('#task-tbody .task-row');

    state.tasks.forEach((task, index) => {
        const trDOM = tableRows[index]; 
        if(!trDOM) return;
        const rowHeight = trDOM.offsetHeight; 
        
        // 行（背景）の作成
        const chartRow = createChartRow(task, rowHeight, chartArea);
        
        // バー本体の作成と配置
        drawBarsForTask(task, chartRow, rowHeight, currentYOffset, calculatedBars, dStart, dEnd);
        
        chartArea.appendChild(chartRow); 
        currentYOffset += rowHeight;
    });

    // 4. 今日の線の描画
    drawTodayLine(chartArea, dStart, dEnd);
    
    // 5. 矢印の描画
    arrowLayer.style.height = currentYOffset + 'px'; 
    drawArrows(calculatedBars, arrowLayer); 
    
    // 6. 画面下部の余白追加
    // ★修正: ボタンの正確な高さを取得し、枠線分の微調整（+1px）を加える
    const addRowBtn = document.querySelector('.add-row-btn');
    const btnHeight = addRowBtn ? addRowBtn.getBoundingClientRect().height : 34;
    
    const bottomPadding = document.createElement('div');
    bottomPadding.style.height = (btnHeight + 1) + 'px'; 
    bottomPadding.style.width = '1px';
    chartArea.appendChild(bottomPadding);
    
    // 7. テキストボックスの描画
    drawTextBoxes(chartArea);
}

// ---------------------------------------------------
// renderChart() を構成する小さなパーツ（関数）たち
// ---------------------------------------------------

// 休日背景を塗る
function drawChartBackground(chartArea, dStart, dEnd) {
    if (state.viewScale !== 'day') return;
    
    const bgLayer = document.createElement('div'); 
    bgLayer.style.cssText = 'position:absolute; top:0; left:0; width:100%; height:100%; z-index:0; pointer-events:none;';
    
    let bgHTML = ''; 
    let bgDate = new Date(dStart); 
    let offset = 0;
    
    while (bgDate <= dEnd) {
        const bgDateStr = formatDate(bgDate); 
        const isNatHoliday = window.nationalHolidays && window.nationalHolidays[bgDateStr];
        if (isNatHoliday) { 
            bgHTML += `<div style="position:absolute; left:${offset * CELL_WIDTH_DAY}px; top:0; width:${CELL_WIDTH_DAY}px; height:100%; background-color:rgba(220,53,69,0.08);"></div>`;
        } else if (isHoliday(bgDate)) { 
            bgHTML += `<div style="position:absolute; left:${offset * CELL_WIDTH_DAY}px; top:0; width:${CELL_WIDTH_DAY}px; height:100%; background-color:rgba(0,0,0,0.04);"></div>`; 
        }
        bgDate.setDate(bgDate.getDate() + 1); 
        offset++;
    }
    bgLayer.innerHTML = bgHTML; 
    chartArea.appendChild(bgLayer);
}

// カレンダーの「行」を作る（クリックでバー生成などのイベント付き）
function createChartRow(task, rowHeight, chartArea) {
    const chartRow = document.createElement('div'); 
    chartRow.className = 'chart-row'; 
    chartRow.style.height = rowHeight + 'px';
    
    // 右クリックで行メニュー表示
    chartRow.addEventListener('contextmenu', (e) => {
        if (e.target.closest('.task-bar') || e.target.closest('path')) return; 
        e.preventDefault();
        e.stopPropagation();
        const title = `行アクション (No.${task.no} ${task.koshu || '名称未設定'})`;
        window.showContextMenu(e, title, 'task', { taskId: task.id });
    });

    // 左クリックで自動バー生成
    chartRow.addEventListener('mousedown', (e) => {
        if (!state.autoCreateBar || currentTool !== 'pointer' || e.button !== 0) return;
        if (e.target.closest('.task-bar') || e.target.closest('path') || e.target.closest('.chart-text-box')) return;
        
        const rect = chartArea.getBoundingClientRect();
        const clickX = e.clientX - rect.left + chartArea.scrollLeft;
        const clickedDateStr = pxToDate(clickX);
        
        if (clickedDateStr) {
            const snappedStart = snapToWorkDay(clickedDateStr, 1);
            const snappedEnd = calcEndDate(snappedStart, 3); // デフォルト3日分
            
            const maxRow = task.periods.reduce((max, p) => Math.max(max, p.displayRow || 0), 0);
            let targetRow = 0; let color = "#3b82f6";

            if (barCreationMode === 'new_change') { targetRow = maxRow + 1; color = "#dc3545"; } 
            else if (barCreationMode === 'current_change') { targetRow = maxRow === 0 ? 1 : maxRow; color = "#dc3545"; }

            let targetPeriod = task.periods.find(p => !p.start && !p.end && (p.displayRow || 0) === targetRow);
            if (!targetPeriod) {
                targetPeriod = { pid: generateId(), dep: "", start: "", end: "", progress: 0, color: color, displayRow: targetRow };
                task.periods.push(targetPeriod);
            }
            targetPeriod.start = snappedStart;
            targetPeriod.end = snappedEnd;
            
            checkAndExtendProjectDates(snappedStart, snappedEnd);
            window.saveStateToHistory(); renderAll();
        }
    });
    return chartRow;
}

// 1つのタスク（行）の中にあるすべてのバーを描く
function drawBarsForTask(task, chartRow, rowHeight, currentYOffset, calculatedBars, dStart, dEnd) {
    const displayName = [task.koshu, task.shubetsu, task.saibetsu].filter(Boolean).join(' ') || '名称未設定';
    const maxRow = task.periods.reduce((max, p) => Math.max(max, p.displayRow || 0), 0);
    const numRows = maxRow + 1; 
    const totalBarHeight = (18 * numRows) + (10 * (numRows - 1));
    const startY = (rowHeight - totalBarHeight) / 2;

    task.periods.forEach((period, pIndex) => {
        if (!period.start || !period.end) return;
        
        const tStart = new Date(period.start); 
        const tEnd = new Date(period.end);
        const leftPx = dateToPx(period.start); 
        const tEndPlusOne = new Date(tEnd); 
        tEndPlusOne.setDate(tEndPlusOne.getDate() + 1); 
        const rightPx = dateToPx(formatDate(tEndPlusOne)); 
        const widthPx = rightPx - leftPx;
        
        const dRow = period.displayRow || 0; 
        const topOffset = startY + (dRow * 28); 
        const barCenterY = currentYOffset + topOffset + 9;

        // 矢印計算のためにデータを保存しておく
        calculatedBars.push({ 
            no: task.no.toString(), taskId: task.id, periodId: period.pid, dep: period.dep, 
            x1: leftPx, x2: leftPx + widthPx, y: barCenterY 
        });

        // 画面の表示範囲に被っている場合のみDOMを作る
        if (tStart <= dEnd && tEnd >= dStart && widthPx > 0) {
            const bar = document.createElement('div'); 
            bar.className = 'task-bar';
            if (selectedItem && selectedItem.type === 'bar' && selectedItem.periodId === period.pid) bar.classList.add('selected');
            
            bar.dataset.taskId = task.id; 
            bar.dataset.periodId = period.pid; 
            bar.style.cssText = `left:${leftPx}px; width:${widthPx}px; top:${topOffset}px; height:18px; line-height:18px; font-size:11px; background-color:${period.color || '#3b82f6'};`;
            
            const p = period.progress || 0;
            if (p > 0) { 
                const progressBar = document.createElement('div'); 
                progressBar.style.cssText = `position:absolute; left:0; top:0; height:100%; width:${p}%; background-color:rgba(0,0,0,0.25); border-radius:${p >= 100 ? '4px' : '4px 0 0 4px'}; pointer-events:none; z-index:1;`;
                bar.appendChild(progressBar); 
            }
            
            const textSpan = document.createElement('span'); 
            textSpan.className = 'task-bar-text'; 
            textSpan.textContent = task.periods.length > 1 ? `${displayName} (${pIndex+1})` : displayName; 
            bar.appendChild(textSpan);
            
            const resizeHandle = document.createElement('div'); 
            resizeHandle.className = 'resize-handle'; 
            bar.appendChild(resizeHandle);
            
            const linkHandle = document.createElement('div'); 
            linkHandle.className = 'link-handle'; 
            linkHandle.title = 'ドラッグして後続作業に繋げる'; 
            bar.appendChild(linkHandle);
            
            setupBarEvents(bar, resizeHandle, task, period, leftPx, widthPx); 
            setupLinkEvents(linkHandle, task, period, leftPx + widthPx - 18, barCenterY);
            
            chartRow.appendChild(bar); 
        }
    });
}

// 今日の赤い線を引く
function drawTodayLine(chartArea, dStart, dEnd) {
    const today = new Date(); 
    today.setHours(0,0,0,0);
    if (today >= dStart && today <= dEnd) {
        let todayX = dateToPx(formatDate(today));
        if (state.viewScale === 'day') { 
            todayX += CELL_WIDTH_DAY / 2; 
        } else { 
            const daysInMonth = new Date(today.getFullYear(), today.getMonth() + 1, 0).getDate(); 
            todayX += (CELL_WIDTH_MONTH / daysInMonth) / 2; 
        }
        const todayLine = document.createElement('div'); 
        todayLine.className = 'today-line'; 
        todayLine.style.left = todayX + 'px'; 
        chartArea.appendChild(todayLine);
    }
}

// テキストボックスの配置
function drawTextBoxes(chartArea) {
    state.texts.forEach(txt => {
        const div = document.createElement('div');
        div.className = 'chart-text-box';
        div.dataset.id = txt.id;
        div.innerHTML = txt.text || '';
        
        div.style.left = txt.x + 'px'; div.style.top = txt.y + 'px';
        div.style.fontFamily = txt.fontFamily || "'Segoe UI', Tahoma, Geneva, Verdana, sans-serif";
        div.style.fontSize = (txt.fontSize || 12) + 'px';
        div.style.color = txt.color || '#212529';
        div.style.backgroundColor = txt.backgroundColor || 'transparent';
        div.style.borderStyle = txt.borderStyle || 'solid';
        div.style.borderWidth = (txt.borderWidth !== undefined ? txt.borderWidth : 1) + 'px';
        div.style.borderColor = txt.borderColor || '#6c757d';
        div.style.fontWeight = txt.fontWeight || 'normal';
        if (txt.width) div.style.width = txt.width + 'px';
        if (txt.height) div.style.height = txt.height + 'px';

        div.style.display = 'flex';
        const alignMap = { 'left': 'flex-start', 'center': 'center', 'right': 'flex-end' };
        div.style.justifyContent = alignMap[txt.textAlign] || 'flex-start';
        div.style.alignItems = txt.verticalAlign || 'flex-start';

        if (selectedItem && selectedItem.type === 'text' && selectedItem.textId === txt.id) div.classList.add('selected');

        // イベント設定：ダブルクリックで編集
        div.addEventListener('dblclick', (e) => {
            if (currentTool !== 'pointer') return;
            e.stopPropagation();
            div.contentEditable = "true";
            div.style.cursor = "text";
            div.focus();
        });
        div.addEventListener('blur', () => {
            div.contentEditable = "false";
            div.style.cursor = "move";
            txt.text = div.innerHTML;
            window.saveStateToHistory();
        });

        // ドラッグ移動設定
        div.addEventListener('mousedown', (e) => {
            if (currentTool !== 'pointer') return; 
            if (div.contentEditable === "true") { e.stopPropagation(); return; }

            e.stopPropagation();
            selectedItem = { type: 'text', textId: txt.id };
            window.updateFormatToolbar();
            document.querySelectorAll('.chart-text-box').forEach(el => el.classList.remove('selected'));
            div.classList.add('selected');
            
            if (e.button === 2) { 
                window.showContextMenu(e, "テキスト操作", 'text', { textId: txt.id });
                return;
            }
            
            if (e.button === 0) { 
                const rect = div.getBoundingClientRect();
                const isResize = (e.clientX > rect.right - 15) && (e.clientY > rect.bottom - 15);
                
                if (isResize) {
                    const onMouseUpResize = () => {
                        txt.width = div.offsetWidth; txt.height = div.offsetHeight;
                        window.saveStateToHistory();
                        document.removeEventListener('mouseup', onMouseUpResize);
                    };
                    document.addEventListener('mouseup', onMouseUpResize);
                    return;
                }

                let isDragging = true;
                let startX = e.clientX, startY = e.clientY;
                let initialX = txt.x, initialY = txt.y;
                div.style.zIndex = '100';

                const onMouseMove = (ev) => {
                    if (!isDragging) return;
                    txt.x = initialX + (ev.clientX - startX); txt.y = initialY + (ev.clientY - startY);
                    div.style.left = txt.x + 'px'; div.style.top = txt.y + 'px';
                };
                const onMouseUp = () => {
                    isDragging = false;
                    document.removeEventListener('mousemove', onMouseMove);
                    document.removeEventListener('mouseup', onMouseUp);
                    div.style.zIndex = '';
                    window.saveStateToHistory(); 
                };
                document.addEventListener('mousemove', onMouseMove);
                document.addEventListener('mouseup', onMouseUp);
            }
        });
        chartArea.appendChild(div);
    });
}

function drawArrows(bars, svgLayer) {
    bars.forEach(bar => {
        if (!bar.dep) return; 
        const depList = bar.dep.toString().split(',');
        depList.forEach(depStr => {
            const depPid = depStr.trim(); 
            if (!depPid) return;
            
            const predBar = bars.find(b => b.periodId === depPid); 
            
            if (predBar) {
                const startX = predBar.x2, startY = predBar.y; 
                const endX = bar.x1, endY = bar.y; 
                let pathD = ''; 
                const curveOffset = 15;
                
                if (startX <= endX - curveOffset) { 
                    pathD = `M ${startX} ${startY} L ${startX + curveOffset} ${startY} L ${startX + curveOffset} ${endY} L ${endX} ${endY}`; 
                } else { 
                    const dropY = Math.max(startY, endY) + 20; 
                    pathD = `M ${startX} ${startY} L ${startX + curveOffset} ${startY} L ${startX + curveOffset} ${dropY} L ${endX - curveOffset} ${dropY} L ${endX - curveOffset} ${endY} L ${endX} ${endY}`; 
                }
                
                const hitPath = document.createElementNS("http://www.w3.org/2000/svg", "path"); 
                hitPath.setAttribute("d", pathD); 
                hitPath.setAttribute("stroke", "rgba(255, 255, 255, 0.01)"); 
                hitPath.setAttribute("stroke-width", "15"); 
                hitPath.setAttribute("fill", "none");
                hitPath.style.cursor = "pointer";
                hitPath.setAttribute("pointer-events", "stroke");

                const path = document.createElementNS("http://www.w3.org/2000/svg", "path"); 
                path.setAttribute("d", pathD); 
                path.setAttribute("stroke", "#ff6b6b"); 
                path.setAttribute("stroke-width", "2"); 
                path.setAttribute("fill", "none");
                path.setAttribute("pointer-events", "none");
                
                if (selectedItem && selectedItem.type === 'arrow' && selectedItem.taskId === bar.taskId && selectedItem.periodId === bar.periodId && selectedItem.predPid === depPid) {
                    path.classList.add('selected');
                    path.setAttribute("stroke", "#dc3545"); 
                    path.setAttribute("stroke-width", "4");
                }
                
                hitPath.addEventListener('mousedown', (e) => { 
                    e.stopPropagation(); 
                    selectedItem = { type: 'arrow', taskId: bar.taskId, periodId: bar.periodId, predPid: depPid }; 
                    renderChart(); 
                    if (e.button === 2) window.showContextMenu(e, "依存関係の操作", 'arrow', { taskId: bar.taskId, periodId: bar.periodId, predPid: depPid }); 
                });
                hitPath.addEventListener('contextmenu', (e) => e.preventDefault());
                
                const arrowHead = document.createElementNS("http://www.w3.org/2000/svg", "polygon"); 
                arrowHead.setAttribute("points", `${endX},${endY} ${endX-6},${endY-4} ${endX-6},${endY+4}`); 
                arrowHead.setAttribute("fill", "#ff6b6b");
                arrowHead.setAttribute("pointer-events", "none");

                const reconnectHandle = document.createElementNS("http://www.w3.org/2000/svg", "circle");
                reconnectHandle.setAttribute("cx", endX);
                reconnectHandle.setAttribute("cy", endY);
                reconnectHandle.setAttribute("r", "12"); 
                reconnectHandle.setAttribute("fill", "rgba(255, 255, 255, 0.01)"); 
                reconnectHandle.style.cursor = "crosshair"; 
                reconnectHandle.setAttribute("pointer-events", "all");
                
                reconnectHandle.addEventListener('mousedown', (e) => {
                    if (e.button !== 0) return; 
                    e.stopPropagation();
                    
                    const chartArea = document.getElementById('chart-area'); 
                    const rect = chartArea.getBoundingClientRect(); 
                    
                    let tempPath = document.createElementNS("http://www.w3.org/2000/svg", "path"); 
                    tempPath.setAttribute("stroke", "#ff6b6b"); 
                    tempPath.setAttribute("stroke-width", "2"); 
                    tempPath.setAttribute("stroke-dasharray", "4,4"); 
                    tempPath.setAttribute("fill", "none"); 
                    tempPath.setAttribute("pointer-events", "none"); 
                    svgLayer.appendChild(tempPath);
                    
                    const onMouseMove = (ev) => {
                        const currentX = ev.clientX - rect.left + chartArea.scrollLeft; 
                        const currentY = ev.clientY - rect.top + chartArea.scrollTop; 
                        const cpX = (startX + currentX) / 2; 
                        tempPath.setAttribute("d", `M ${startX} ${startY} C ${cpX} ${startY}, ${cpX} ${currentY}, ${currentX} ${currentY}`);
                    };
                    
                    const onMouseUp = (ev) => {
                        document.removeEventListener('mousemove', onMouseMove); 
                        document.removeEventListener('mouseup', onMouseUp); 
                        if (tempPath) tempPath.remove();
                        
                        const targetEl = document.elementFromPoint(ev.clientX, ev.clientY); 
                        let targetBar = targetEl ? targetEl.closest('.task-bar') : null;
                        
                        let oldDeps = bar.dep.toString().split(',').map(s => s.trim()).filter(d => d !== depPid);
                        const sourceTask = state.tasks.find(t => t.id === bar.taskId);
                        const sourcePeriod = sourceTask.periods.find(p => p.pid === bar.periodId);
                        sourcePeriod.dep = oldDeps.join(', ');

                        if (targetBar) {
                            const targetTaskId = targetBar.dataset.taskId; 
                            const targetPeriodId = targetBar.dataset.periodId;
                            
                            if (targetTaskId !== predBar.taskId || targetPeriodId !== predBar.periodId) {
                                const targetTask = state.tasks.find(t => t.id === targetTaskId); 
                                const targetPeriod = targetTask.periods.find(p => p.pid === targetPeriodId);
                                
                                if (targetTask && targetPeriod) {
                                    let currentDep = targetPeriod.dep ? targetPeriod.dep.toString().split(',').map(s=>s.trim()) : [];
                                    if (!currentDep.includes(depPid)) {
                                        currentDep.push(depPid);
                                        targetPeriod.dep = currentDep.join(', ');
                                    }
                                }
                            }
                        }
                        window.saveStateToHistory();
                        renderAll();
                    };
                    document.addEventListener('mousemove', onMouseMove); 
                    document.addEventListener('mouseup', onMouseUp);
                });

                svgLayer.appendChild(hitPath);
                svgLayer.appendChild(path); 
                svgLayer.appendChild(arrowHead);
                svgLayer.appendChild(reconnectHandle);
            }
        });
    });
}

// ---------------------------------------------------
// 5. ドラッグ＆リサイズイベント等
// ---------------------------------------------------
function setupBarEvents(bar, handle, task, period, initialLeft, initialWidth) {
    let dStartX = 0, currentLeft = initialLeft; 
    let moved = false;
    
    const onDragMove = (e) => { 
        moved = true; 
        bar.style.left = (currentLeft + (e.clientX - dStartX)) + 'px'; 
    };
    const onDragUp = (e) => {
        document.removeEventListener('mousemove', onDragMove); 
        document.removeEventListener('mouseup', onDragUp); 
        bar.style.zIndex = '2';
        if (moved) { 
            let newLeft = parseFloat(bar.style.left); 
            window.handlePeriodChange(task.id, period.pid, 'start', pxToDate(newLeft)); 
        } else { 
            renderChart(); 
        }
    };
    
    bar.addEventListener('mousedown', (e) => {
        e.stopPropagation(); 
        selectedItem = { type: 'bar', taskId: task.id, periodId: period.pid }; 
        document.querySelectorAll('.task-bar').forEach(b => b.classList.remove('selected')); 
        document.querySelectorAll('#arrow-layer path').forEach(p => p.classList.remove('selected')); 
        bar.classList.add('selected');
        
        if (e.button === 2) { 
            const name = [task.koshu, task.shubetsu, task.saibetsu].filter(Boolean).join(' ') || "名称未設定"; 
            window.showContextMenu(e, name, 'bar', { taskId: task.id, periodId: period.pid }); 
            return; 
        }
        
        if (e.button !== 0) return; 
        if (e.target.closest('.resize-handle') || e.target.closest('.link-handle')) return;

        dStartX = e.clientX; 
        currentLeft = parseFloat(bar.style.left) || 0; 
        bar.style.zIndex = '10'; 
        moved = false; 
        document.addEventListener('mousemove', onDragMove); 
        document.addEventListener('mouseup', onDragUp); 
        e.preventDefault(); 
    });

    if (handle) {
        let rStartX = 0, currentWidth = initialWidth;
        let resizing = false;

        const onResizeMove = (e) => {
            resizing = true;
            let newWidth = currentWidth + (e.clientX - rStartX);
            if (newWidth < 10) newWidth = 10;
            bar.style.width = newWidth + 'px';
        };

        const onResizeUp = (e) => {
            document.removeEventListener('mousemove', onResizeMove);
            document.removeEventListener('mouseup', onResizeUp);
            bar.style.zIndex = '2';
            
            if (resizing) {
                let finalWidth = parseFloat(bar.style.width);
                let newRightPx = initialLeft + finalWidth;
                let newEndPlusOneStr = pxToDate(newRightPx);
                let d = new Date(newEndPlusOneStr);
                d.setDate(d.getDate() - 1);
                window.handlePeriodChange(task.id, period.pid, 'end', formatDate(d));
            } else {
                renderChart();
            }
        };

        handle.addEventListener('mousedown', (e) => {
            if (e.button !== 0) return; 
            e.stopPropagation(); 
            rStartX = e.clientX;
            currentWidth = parseFloat(bar.style.width) || initialWidth;
            bar.style.zIndex = '10';
            resizing = false;
            document.addEventListener('mousemove', onResizeMove);
            document.addEventListener('mouseup', onResizeUp);
            e.preventDefault();
        });
    }
}

let linkingState = null;
function setupLinkEvents(handle, task, period, startX, startY) {
    handle.addEventListener('mousedown', (e) => {
        if (e.button !== 0) return; 
        e.stopPropagation();
        
        const chartArea = document.getElementById('chart-area'); 
        const svgLayer = document.getElementById('arrow-layer'); 
        const rect = chartArea.getBoundingClientRect(); 
        
        linkingState = { sourceTask: task, sourcePeriod: period, startX: startX, startY: startY };
        
        let tempPath = document.createElementNS("http://www.w3.org/2000/svg", "path"); 
        tempPath.setAttribute("id", "temp-link-line"); 
        tempPath.setAttribute("stroke", "#ff6b6b"); 
        tempPath.setAttribute("stroke-width", "2"); 
        tempPath.setAttribute("stroke-dasharray", "4,4"); 
        tempPath.setAttribute("fill", "none"); 
        tempPath.setAttribute("pointer-events", "none"); 
        svgLayer.appendChild(tempPath);
        
        const onMouseMove = (ev) => {
            if (!linkingState) return;
            const currentX = ev.clientX - rect.left + chartArea.scrollLeft; 
            const currentY = ev.clientY - rect.top + chartArea.scrollTop; 
            const cpX = (startX + currentX) / 2; 
            tempPath.setAttribute("d", `M ${startX} ${startY} C ${cpX} ${startY}, ${cpX} ${currentY}, ${currentX} ${currentY}`);
        };
        
        const onMouseUp = (ev) => {
            document.removeEventListener('mousemove', onMouseMove); 
            document.removeEventListener('mouseup', onMouseUp); 
            if (tempPath) tempPath.remove();
            
            const targetEl = document.elementFromPoint(ev.clientX, ev.clientY); 
            let targetBar = targetEl ? targetEl.closest('.task-bar') : null;
            
            if (targetBar) {
                const targetTaskId = targetBar.dataset.taskId; 
                const targetPeriodId = targetBar.dataset.periodId;
                if (targetTaskId && (targetTaskId !== linkingState.sourceTask.id || targetPeriodId !== linkingState.sourcePeriod.pid)) {
                    const targetTask = state.tasks.find(t => t.id === targetTaskId); 
                    const targetPeriod = targetTask.periods.find(p => p.pid === targetPeriodId);
                    if (targetTask && targetPeriod) {
                        let currentDep = targetPeriod.dep ? targetPeriod.dep.toString().trim() : ""; 
                        const sourcePid = linkingState.sourcePeriod.pid;
                        let deps = currentDep ? currentDep.split(',').map(s => s.trim()) : [];
                        if (!deps.includes(sourcePid)) { 
                            if (currentDep) currentDep += ", " + sourcePid;
                            else currentDep = sourcePid;
                            window.handlePeriodChange(targetTaskId, targetPeriodId, 'dep', currentDep); 
                        }
                    }
                }
            }
            linkingState = null;
        };
        document.addEventListener('mousemove', onMouseMove); 
        document.addEventListener('mouseup', onMouseUp);
    });
}

function setupTableResizing() {
    const table = document.getElementById('task-table');
    if (!table) return;
    const cols = table.querySelectorAll('thead th');
    
    cols.forEach((col, index) => {
        if (index === cols.length - 1) return;

        const resizer = document.createElement('div');
        resizer.className = 'col-resizer';
        col.appendChild(resizer);
        
        let startX, startWidth;

        resizer.addEventListener('mousedown', function(e) {
            startX = e.clientX;
            startWidth = col.offsetWidth;
            
            cols.forEach(c => c.style.width = c.offsetWidth + 'px');
            resizer.classList.add('resizing');

            const mouseMoveHandler = function(e) {
                const newWidth = startWidth + (e.clientX - startX);
                if (newWidth > 30) col.style.width = newWidth + 'px';
            };

            const mouseUpHandler = function() {
                resizer.classList.remove('resizing');
                document.removeEventListener('mousemove', mouseMoveHandler);
                document.removeEventListener('mouseup', mouseUpHandler);
            };

            document.addEventListener('mousemove', mouseMoveHandler);
            document.addEventListener('mouseup', mouseUpHandler);
            e.stopPropagation();
            e.preventDefault();
        });
    });
}

document.addEventListener('DOMContentLoaded', () => {
    setTimeout(setupTableResizing, 100); 
});