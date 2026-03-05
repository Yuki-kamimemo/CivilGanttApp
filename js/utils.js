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
    window.renderChart();
};
