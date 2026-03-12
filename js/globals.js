/**
 * Civil Schedule Master (CivilGanttApp)
 * Copyright (c) 2026 [Your Company Name]
 * All rights reserved.
 * This software is proprietary and confidential.
 */

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
            window.renderAll();
        } else {
            console.warn('祝日データの取得に失敗しました（サーバーエラー）。祝日は稼働日として扱われます。');
            showHolidayWarning();
        }
    } catch (e) {
        console.warn('祝日データの取得に失敗しました（ネットワークエラー）。', e);
        showHolidayWarning();
    }
}

function showHolidayWarning() {
    // 画面上部に警告バナーを表示（邪魔にならないよう数秒で消える）
    const banner = document.createElement('div');
    banner.className = 'holiday-warning-banner';
    banner.textContent = '⚠ インターネットに接続できないため、祝日データを取得できませんでした。日本の祝日は稼働日として扱われます。';
    document.body.appendChild(banner);
    setTimeout(() => {
        banner.style.opacity = '0';
        setTimeout(() => banner.remove(), 500);
    }, 6000);
}

/**
 * alert() の代わりに使う非ブロッキングのトースト通知。
 * 保存完了などの軽い通知に使う。
 */
function showToastNotification(message) {
    const toast = document.createElement('div');
    toast.className = 'toast-notification';
    toast.textContent = message;
    document.body.appendChild(toast);
    // アニメーションのために少し遅らせてクラスを付加
    setTimeout(() => toast.classList.add('show'), 10);
    setTimeout(() => {
        toast.classList.remove('show');
        setTimeout(() => toast.remove(), 400);
    }, 2500);
}

// ---------------------------------------------------
// 1. データモデルと履歴管理 (State)
// ---------------------------------------------------

/**
 * デフォルトのstateオブジェクトを生成する関数。
 * globals.js の初期化と handleNewFile() の両方で使用することで
 * state の初期定義を一元管理する。
 */
window.createDefaultState = function () {
    const today = new Date();
    const nextMonth = new Date(today);
    nextMonth.setMonth(nextMonth.getMonth() + 1);

    return {
        projectName: '', companyName: '', viewRange: 'month', viewScale: 'day',
        projectStart: formatDate(today), projectEnd: formatDate(nextMonth),
        displayStart: formatDate(today), displayEnd: formatDate(nextMonth),
        notes: '', notesCollapsed: false,
        dailyNoteTabs: [
            { id: 'tab_general', name: '作業全般・天候' },
            { id: 'tab_safety', name: '安全管理・行事' }
        ],
        activeDailyNoteTab: 'tab_general',
        dailyNotesData: { 'tab_general': {}, 'tab_safety': {} },
        dailyNotesMerges: { 'tab_general': {}, 'tab_safety': {} },
        holidays: { sundays: true, saturdays: false, nationalHolidays: true, custom: [] },
        tasks: [
            {
                id: generateId(), no: 1, koshu: "", shubetsu: "", saibetsu: "", collapsed: false,
                mergeAboveKoshu: false, mergeAboveShubetsu: false,
                periods: [{ pid: generateId(), dep: "", start: "", end: "", progress: 0, color: "#3b82f6", displayRow: 0 }]
            }
        ],
        texts: [], shapes: [],
        autoCreateBar: true,
        globalFontFamily: "'Segoe UI', Tahoma, Geneva, Verdana, sans-serif",
        globalFontSize: 13
    };
};

let state = window.createDefaultState();

// 描画ツールの状態管理
let currentTool = 'pointer'; // 初期状態は「選択・操作」ツール

window.setTool = function (toolName) {
    currentTool = toolName;

    // 全てのメニューから「✓」を外し、選ばれたものだけに「✓」をつける
    const tools = ['pointer', 'text'];
    tools.forEach(t => {
        const el = document.getElementById('menu-tool-' + t);
        if (el) {
            if (!el.hasAttribute('data-label')) {
                const text = el.textContent.replace(/^([✓] |　 )/, '').trim();
                el.setAttribute('data-label', text);
            }
            const label = el.getAttribute('data-label') || '';
            el.textContent = (t === toolName ? '✓ ' : '　 ') + label;
        }
    });

    // 描画ツールの時は、カレンダー上のカーソルを十字にする
    const chartArea = document.getElementById('chart-area');
    if (chartArea) {
        chartArea.style.cursor = (toolName === 'pointer') ? 'default' : 'crosshair';
    }
};

// バー作成モードの管理
let barCreationMode = 'normal'; // 初期状態は「予定」

window.setBarCreationMode = function (mode) {
    barCreationMode = mode;

    const btnNormal = document.getElementById('mode-normal');
    const btnNew = document.getElementById('mode-new-change');
    const btnCurrent = document.getElementById('mode-current-change');

    if (btnNormal && btnNew && btnCurrent) {
        [btnNormal, btnNew, btnCurrent].forEach(btn => btn.classList.remove('active'));

        if (mode === 'normal') {
            btnNormal.classList.add('active');
        } else if (mode === 'new_change') {
            btnNew.classList.add('active');
        } else if (mode === 'current_change') {
            btnCurrent.classList.add('active');
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

window.saveStateToHistory = function () {
    if (historyIndex < stateHistory.length - 1) {
        stateHistory = stateHistory.slice(0, historyIndex + 1);
    }
    stateHistory.push(JSON.parse(JSON.stringify(state)));
    if (stateHistory.length > MAX_HISTORY) {
        stateHistory.shift();
        // shift()で先頭を削除した分、indexは変わらない（常に末尾を指すよう維持）
        historyIndex = stateHistory.length - 1;
    } else {
        historyIndex++;
    }
    updateUndoRedoUI();
}

window.undo = function () {
    if (historyIndex > 0) {
        historyIndex--;
        state = JSON.parse(JSON.stringify(stateHistory[historyIndex]));
        window.renderAll();
        updateUndoRedoUI();
    }
}

window.redo = function () {
    if (historyIndex < stateHistory.length - 1) {
        historyIndex++;
        state = JSON.parse(JSON.stringify(stateHistory[historyIndex]));
        window.renderAll();
        updateUndoRedoUI();
    }
}

function updateUndoRedoUI() {
    const undoBtn = document.getElementById('menu-undo');
    const redoBtn = document.getElementById('menu-redo');
    if (undoBtn) undoBtn.classList.toggle('disabled', historyIndex <= 0);
    if (redoBtn) redoBtn.classList.toggle('disabled', historyIndex >= stateHistory.length - 1);
}

// 矢印接続中の状態管理（events.jsと共有）
let linkingState = null;

// 依存関係の循環（ループ）チェック関数
window.isCircularDependency = function (sourcePid, targetPid) {
    if (sourcePid === targetPid) return true;

    const visited = new Set();
    const queue = [sourcePid];

    while (queue.length > 0) {
        const currentPid = queue.shift();
        if (currentPid === targetPid) return true;

        if (!visited.has(currentPid)) {
            visited.add(currentPid);
            for (const task of state.tasks) {
                const period = task.periods.find(p => p.pid === currentPid);
                if (period && period.dep) {
                    const deps = period.dep.toString().split(',').map(s => s.trim()).filter(Boolean);
                    for (const dep of deps) {
                        if (!visited.has(dep)) queue.push(dep);
                    }
                }
            }
        }
    }
    return false;
};

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

// ヘルパー関数: UI要素でのクリックかどうかを判定
function isUIElementClick(e) {
    if (e.target.closest('#color-palette-popup')) return false;
    return e.target.closest('.context-menu') ||
        e.target.closest('button') ||
        e.target.closest('.menubar') ||
        e.target.closest('.ql-toolbar') ||    // Quillエディタのツールバー
        e.target.closest('.ql-container') ||  // Quillエディタのコンテナ
        e.target.closest('.daily-note-cell') || 
        e.target.closest('.ql-editor') ||     // Quillのエディタエリア本体
        e.target.tagName === 'INPUT' ||
        e.target.tagName === 'TEXTAREA' ||
        e.target.closest('[contenteditable="true"]');
}

// ヘルパー関数: チャート要素でのクリックかどうかを判定
function isChartElementClick(e) {
    return e.target.closest('.task-bar') ||
        e.target.closest('path') ||
        e.target.closest('.chart-text-box');
}

document.addEventListener('mousedown', (e) => {
    const menu = document.getElementById('custom-context-menu');

    // UI要素のクリックならコンテキストメニューを閉じて何もしない
    if (isUIElementClick(e)) {
        if (menu && !e.target.closest('.context-menu')) {
            menu.style.display = 'none';
        }
        return;
    }

    // チャート要素以外の場所をクリックした場合は選択状態を解除して再描画
    if (!isChartElementClick(e)) {
        if (selectedItem) {
            selectedItem = null;
            if (window.updateFormatToolbar) window.updateFormatToolbar();
            window.renderChart();
        }
    }

    // それ以外の通常のクリックでもコンテキストメニューは非表示にする
    if (menu) menu.style.display = 'none';
});
