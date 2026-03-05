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

    setTimeout(setupTableResizing, 100);
});
