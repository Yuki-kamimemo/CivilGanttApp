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
    window.openConfirmModal(
        'ファイルを開く',
        '現在のデータは破棄されます。保存していない変更は失われますが、よろしいですか？',
        async function() {
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
        }
    );
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
    // 内容が変わっていない場合は履歴を保存しない（Undoの無駄遣いを防ぐ）
    const oldValue = state.dailyNotesData[tabId][dateStr] || '';
    if (oldValue === value) return;
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
