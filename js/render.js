// ---------------------------------------------------
// 4. 描画処理（View）
// ---------------------------------------------------
function renderAll() {
    // ★追加：再描画の前に現在のスクロール位置を記憶する
    const leftContainer = document.getElementById('left-container');
    const rightContainer = document.getElementById('right-container');
    const dailyNotesRight = document.getElementById('daily-notes-right');
    const savedScrollTop = rightContainer ? rightContainer.scrollTop : 0;
    const savedScrollLeft = rightContainer ? rightContainer.scrollLeft : 0;
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
    const notesArea = document.getElementById('project-notes-editor-container');
    if (window.applyNotesToEditor) {
        window.applyNotesToEditor();
    }

    const ns = state.notesStyle || {};
    if (notesArea) {
        notesArea.style.fontFamily = ns.fontFamily || '';
        notesArea.style.fontSize = ns.fontSize ? ns.fontSize + 'px' : '';
        notesArea.style.color = ns.color || '';
        notesArea.style.fontWeight = ns.fontWeight || '';
        notesArea.style.backgroundColor = ns.backgroundColor || '';
        notesArea.style.writingMode = ns.writingMode || 'horizontal-tb';
        notesArea.style.display = 'flex';
    }
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
    renderChart(true); // renderAll 側でスクロール復元を行うため、renderChart 内での復元はスキップする
    renderDailyNotes();

    // ★追加：再描画が終わった直後にスクロール位置を復元する
    if (rightContainer) {
        rightContainer.scrollTop = savedScrollTop;
        rightContainer.scrollLeft = savedScrollLeft;
    }
    if (leftContainer) {
        leftContainer.scrollTop = savedScrollTop;
    }
    if (dailyNotesRight) {
        dailyNotesRight.scrollLeft = savedScrollLeft;
    }
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
    const merges = (state.dailyNotesMerges && state.dailyNotesMerges[state.activeDailyNoteTab]) || {};

    // 全セルキーのリストを事前に作成（結合スキップ処理のため）
    const allCellKeys = [];
    if (state.viewScale === 'day') {
        let d = new Date(state.displayStart);
        const end = new Date(state.displayEnd);
        while (d <= end) {
            allCellKeys.push({ key: formatDate(d), dateObj: new Date(d) });
            d.setDate(d.getDate() + 1);
        }
    } else if (state.viewScale === 'month') {
        let d = new Date(new Date(state.displayStart).getFullYear(), new Date(state.displayStart).getMonth(), 1);
        const end = new Date(new Date(state.displayEnd).getFullYear(), new Date(state.displayEnd).getMonth(), 1);
        while (d <= end) {
            const key = d.getFullYear() + '-' + String(d.getMonth() + 1).padStart(2, '0');
            allCellKeys.push({ key, dateObj: new Date(d) });
            d.setMonth(d.getMonth() + 1);
        }
    }

    let totalWidth = 0;

    const hAlignMap = { 'left': 'flex-start', 'center': 'center', 'right': 'flex-end' };
    const vAlignMap = { 'flex-start': 'flex-start', 'center': 'center', 'flex-end': 'flex-end' };

    const drawCell = (dateStr, dateObj, widthPx, isMerged = false) => {
        const cell = document.createElement('div');
        cell.className = 'daily-note-cell' + (isMerged ? ' daily-note-cell--merged' : '');
        cell.style.width = widthPx + 'px';
        cell.style.minWidth = widthPx + 'px';
        cell.dataset.dateStr = dateStr;

        const textarea = document.createElement('div');
        // Quillの仕様にあわせたセルスタイル
        textarea.style.cssText = "width:100%; height:100%; box-sizing:border-box; padding:5px; outline:none; overflow-y:auto; display:flex; flex-direction:column; cursor: pointer;";
        textarea.classList.add('ql-editor', 'ql-editor-content');
        textarea.title = "クリックで日報を編集";

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

        // Quillエディタから保存されたHTMLをそのまま流し込む
        textarea.innerHTML = tabData[dateStr] || '';

        // インライン編集のフック
        textarea.addEventListener('mousedown', (e) => {
            // すでに自分が編集対象（アクティブ）なら、何もしないでブラウザの標準動作（選択など）に任せる
            if (window.editorManager && window.editorManager.activeContainer === textarea) {
                return;
            }

            // バブリングを止めて「画面外クリック判定（閉じる処理）」を阻止
            e.stopPropagation();

            window.selectInput('daily_notes');

            if (window.editorManager) {
                const isVertical = currentWM === 'vertical-rl';
                // 前のエディタが残っている可能性を考慮して一旦閉じる
                if (window.editorManager.activeContainer && window.editorManager.activeContainer !== window.editorManager.defaultContainer) {
                    window.editorManager.closeEditor();
                }

                window.editorManager.openEditor(textarea, (html, delta) => {
                    const cleanHtml = (html === '<p><br></p>' || !html) ? '' : html;
                    window.handleDailyNoteChange(dateStr, cleanHtml);
                }, isVertical);
            }
        });

        // 右クリックで結合メニューを表示
        cell.addEventListener('contextmenu', (e) => {
            e.preventDefault();
            e.stopPropagation();
            const tabId = state.activeDailyNoteTab;
            const currentMerges = (state.dailyNotesMerges && state.dailyNotesMerges[tabId]) || {};
            const currentColspan = currentMerges[dateStr] || 1;
            window.showContextMenu(e, '備考セル', 'daily_note_cell', {
                dateStr,
                isMerged: currentColspan > 1
            });
        });

        cell.appendChild(textarea);
        grid.appendChild(cell);
        totalWidth += widthPx;
    };

    const baseWidth = state.viewScale === 'day' ? CELL_WIDTH_DAY : CELL_WIDTH_MONTH;
    let skipCount = 0;
    allCellKeys.forEach(({ key, dateObj }, idx) => {
        if (skipCount > 0) { skipCount--; return; }
        const colspan = merges[key] || 1;
        const actualColspan = Math.min(colspan, allCellKeys.length - idx);
        drawCell(key, dateObj, baseWidth * actualColspan, actualColspan > 1);
        if (actualColspan > 1) skipCount = actualColspan - 1;
    });

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

// HTMLタグを除いたプレーンテキストを返す（バーラベル・コンテキストメニュー用）
function stripHtml(html) {
    if (!html) return '';
    return html.replace(/<[^>]+>/g, '').replace(/&amp;/g, '&').replace(/&lt;/g, '<').replace(/&gt;/g, '>');
}

function renderTable() {
    const tbody = document.getElementById('task-tbody');
    tbody.innerHTML = '';

    // セル結合用の行数を計算
    let koshuRowspans = new Array(state.tasks.length).fill(1);
    let shubetsuRowspans = new Array(state.tasks.length).fill(1);
    for (let i = state.tasks.length - 1; i > 0; i--) {
        if (state.tasks[i].mergeAboveKoshu) {
            koshuRowspans[i - 1] += koshuRowspans[i];
            koshuRowspans[i] = 0;
        }
        if (state.tasks[i].mergeAboveShubetsu) {
            shubetsuRowspans[i - 1] += shubetsuRowspans[i];
            shubetsuRowspans[i] = 0;
        }
    }

    state.tasks.forEach((task, index) => {
        task.no = index + 1;
        const tr = document.createElement('tr');
        tr.className = `task-row ${task.collapsed ? 'collapsed' : ''}`;
        tr.dataset.taskId = task.id;
        // data属性にIDを保持してaddEventListenerで呼ぶ（工種名の特殊文字対策）
        tr.addEventListener('contextmenu', (e) => {
            if (e.target.closest('.task-cell-editor-container')) return;
            const title = `行アクション (No.${task.no} ${stripHtml(task.koshu) || '名称未設定'})`;
            window.showContextMenu(e, title, 'task', { taskId: task.id });
        });

        const maxRow = task.periods.reduce((max, p) => Math.max(max, p.displayRow || 0), 0);
        const numRows = maxRow + 1;
        const requiredHeight = Math.max(40, (18 * numRows) + (10 * maxRow) + 16);

        let html = `
            <td class="task-no-cell" style="text-align: center; vertical-align: middle; height: ${requiredHeight}px;">
                <span class="task-no">${task.no}</span>
            </td>
        `;

        if (koshuRowspans[index] > 0) {
            html += `
            <td class="task-koshu-cell" rowspan="${koshuRowspans[index]}">
                <div class="task-name-header">
                    <div class="task-cell-editor-container"
                         data-task-id="${task.id}"
                         data-field="koshu"
                         data-placeholder="工種"></div>
                </div>
            </td>`;
        }

        if (shubetsuRowspans[index] > 0) {
            html += `
            <td class="task-shubetsu-cell" rowspan="${shubetsuRowspans[index]}">
                <div class="task-cell-editor-container"
                     data-task-id="${task.id}"
                     data-field="shubetsu"
                     data-placeholder="種別"></div>
            </td>`;
        }

        html += `
            <td class="task-saibetsu-cell">
                <div class="task-name-container">
                    <div class="task-cell-editor-container"
                         data-task-id="${task.id}"
                         data-field="saibetsu"
                         data-placeholder="細別・規格"></div>
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
        // divコンテナにHTMLを流し込みイベントリスナーを設定する
        ['koshu', 'shubetsu', 'saibetsu'].forEach(field => {
            const container = tr.querySelector(`.task-cell-editor-container[data-field="${field}"]`);
            if (!container) return;
            const htmlContent = task[field] || '';
            container.innerHTML = htmlContent;
            container.classList.toggle('empty-cell', !htmlContent);
            setupCellEditorListener(container, task.id, field);
        });
        tbody.appendChild(tr);
    });
}

// 工種/種別/細別セルのQuillインライン編集イベントを設定（日報セルと同パターン）
function setupCellEditorListener(container, taskId, field) {
    container.addEventListener('mousedown', (e) => {
        if (window.editorManager && window.editorManager.activeContainer === container) return;
        e.stopPropagation();

        if (window.editorManager) {
            if (window.editorManager.activeContainer &&
                window.editorManager.activeContainer !== window.editorManager.defaultContainer) {
                window.editorManager.closeEditor();
            }
            window.editorManager.openEditor(container, (html) => {
                const cleanHtml = (html === '<p><br></p>' || !html) ? '' : html;
                container.classList.toggle('empty-cell', !cleanHtml);
                window.handleTaskDetailChange(taskId, field, cleanHtml);
            });
        }
    });
}

// 複数の期間データをまとめるための便利関数
function aggregatePeriods(periods) {
    if (periods.length === 0) return null;
    let minStart = null, maxEnd = null, totalWorkDays = 0, weightedProgress = 0;

    periods.forEach(p => {
        // ISO-8601形式の日付文字列はそのまま文字列比較できる（new Date()生成不要）
        if (p.start && (!minStart || p.start < minStart)) minStart = p.start;
        if (p.end && (!maxEnd || p.end > maxEnd)) maxEnd = p.end;
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
function renderChart(skipScrollRestore = false) {
    // ★追加：再描画の前にスクロール位置を記憶する（スキップされない場合のみ）
    const rightContainer = document.getElementById('right-container');
    let savedScrollTop = 0, savedScrollLeft = 0;

    if (!skipScrollRestore && rightContainer) {
        savedScrollTop = rightContainer.scrollTop;
        savedScrollLeft = rightContainer.scrollLeft;
    }

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
        if (!trDOM) return;
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

    // ★追加：スクロール位置を復元する（スキップされない場合のみ）
    if (!skipScrollRestore && rightContainer) {
        rightContainer.scrollTop = savedScrollTop;
        rightContainer.scrollLeft = savedScrollLeft;
    }
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
        const title = `行アクション (No.${task.no} ${stripHtml(task.koshu) || '名称未設定'})`;
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
    const displayName = [task.koshu, task.shubetsu, task.saibetsu].map(stripHtml).filter(Boolean).join(' ') || '名称未設定';
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
            textSpan.textContent = task.periods.length > 1 ? `${displayName} (${pIndex + 1})` : displayName;
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
    today.setHours(0, 0, 0, 0);
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
        // すでに編集中のコンテナがある場合、そのコンテナは再作成せずに維持したい
        const isEditingThis = window.editorManager && window.editorManager.activeContainer && window.editorManager.activeContainer.dataset.id === txt.id;

        let div;
        if (isEditingThis) {
            div = window.editorManager.activeContainer;
            // 座標やサイズのみ更新（内容はQuillが管理中なので触らない）
            div.style.left = txt.x + 'px'; div.style.top = txt.y + 'px';
            if (txt.width) div.style.width = txt.width + 'px';
            if (txt.height) div.style.height = txt.height + 'px';
            chartArea.appendChild(div);
            return; // 編集中の場合は以下の新規作成・イベント設定をスキップ
        }

        div = document.createElement('div');
        // ql-editor クラスを付与することで、Quillの書式（ql-size等）を有効にする
        div.className = 'chart-text-box ql-editor ql-editor-content';
        div.dataset.id = txt.id;
        div.innerHTML = txt.text || '';

        // 基本的な配置スタイル（padding:0 は ql-editor のデフォルトを打ち消すため）
        div.style.position = 'absolute';
        div.style.padding = '5px';
        div.style.left = txt.x + 'px'; div.style.top = txt.y + 'px';
        div.style.fontFamily = txt.fontFamily || "'Segoe UI', Tahoma, Geneva, Verdana, sans-serif";

        // 内部にリッチテキスト（HTML）がある場合は、コンテナ側でのfontSize指定を避ける
        // (Quill内部の span.ql-size-... 等に任せる)
        if (!txt.text || !txt.text.includes('<')) {
            div.style.fontSize = (txt.fontSize || 12) + 'px';
        }

        div.style.color = txt.color || '#212529';
        div.style.backgroundColor = txt.backgroundColor || 'transparent';
        div.style.borderStyle = txt.borderStyle || 'solid';
        div.style.borderWidth = (txt.borderWidth !== undefined ? txt.borderWidth : 1) + 'px';
        div.style.borderColor = txt.borderColor || '#6c757d';
        div.style.fontWeight = txt.fontWeight || 'normal';
        if (txt.width) div.style.width = txt.width + 'px';
        if (txt.height) div.style.height = txt.height + 'px';

        div.style.display = 'flex';
        div.style.flexDirection = 'column';
        div.style.overflow = 'hidden';

        const alignMap = { 'left': 'flex-start', 'center': 'center', 'right': 'flex-end' };

        // 内部にリッチテキストがある場合は、Quillの <p class="ql-align-..."> に任せるため、
        // コンテナ側の flex 揃えは補助的なものに留める
        if (txt.text && txt.text.includes('ql-align-')) {
            div.style.justifyContent = 'flex-start';
            div.style.alignItems = 'stretch';
        } else {
            div.style.justifyContent = txt.verticalAlign || 'flex-start';
            div.style.alignItems = alignMap[txt.textAlign] || 'flex-start';
        }

        if (selectedItem && selectedItem.type === 'text' && selectedItem.textId === txt.id) div.classList.add('selected');

        // イベント設定：ダブルクリックで編集（シングルクリックは選択と移動に使う）
        // ただし、すでに選択されている状態でクリックした場合は編集を開始する
        div.addEventListener('click', (e) => {
            if (currentTool !== 'pointer') return;
            // エディタマネージャーがこの要素を編集中の場合は何もしない（Quillに任せる）
            if (window.editorManager && window.editorManager.activeContainer === div) {
                return;
            }

            // すでに選択されているテキストボックスを再度クリックした場合は編集開始
            if (selectedItem && selectedItem.type === 'text' && selectedItem.textId === txt.id) {
                startEditingTextBox(div, txt);
            }
        });

        div.addEventListener('dblclick', (e) => {
            if (currentTool !== 'pointer') return;
            e.stopPropagation();
            startEditingTextBox(div, txt);
        });

        // 共通の編集開始処理
        function startEditingTextBox(targetDiv, targetTxt) {
            if (!window.editorManager) return;
            const isVertical = targetDiv.style.writingMode === 'vertical-rl';

            window.editorManager.openEditor(targetDiv, (html, delta, boxStyles) => {
                const cleanHtml = (html === '<p><br></p>' || !html) ? '' : html;
                targetTxt.text = cleanHtml;

                if (boxStyles) {
                    targetTxt.borderStyle = boxStyles.borderStyle;
                    targetTxt.borderWidth = boxStyles.borderWidth;
                    targetTxt.borderColor = boxStyles.borderColor;
                    targetTxt.backgroundColor = boxStyles.backgroundColor;
                }

                window.saveStateToHistory();
                // 編集終了後の再描画。編集中（activeContainerがある間）は全体再描画を抑制するのが望ましい
                window.renderChart();
            }, isVertical);
        }

        // ドラッグ移動設定
        div.addEventListener('mousedown', (e) => {
            if (currentTool !== 'pointer') return;
            // エディタマネージャーがこの要素を編集中の場合はドラッグを無効化
            if (window.editorManager && window.editorManager.activeContainer === div) {
                return;
            }

            e.stopPropagation();
            selectedItem = { type: 'text', textId: txt.id };
            // 他の要素の選択を解除
            document.querySelectorAll('.chart-text-box, .task-bar').forEach(el => el.classList.remove('selected'));
            div.classList.add('selected');

            if (window.updateFormatToolbar) window.updateFormatToolbar();

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
    // O(n²)のfind()をO(1)のMap参照に変換
    const barMap = new Map(bars.map(b => [b.periodId, b]));

    bars.forEach(bar => {
        if (!bar.dep) return;
        const depList = bar.dep.toString().split(',');
        depList.forEach(depStr => {
            const depPid = depStr.trim();
            if (!depPid) return;

            const predBar = barMap.get(depPid);  // O(1)での参照

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
                    window.renderChart();
                    if (e.button === 2) window.showContextMenu(e, "依存関係の操作", 'arrow', { taskId: bar.taskId, periodId: bar.periodId, predPid: depPid });
                });
                hitPath.addEventListener('contextmenu', (e) => e.preventDefault());

                const arrowHead = document.createElementNS("http://www.w3.org/2000/svg", "polygon");
                arrowHead.setAttribute("points", `${endX},${endY} ${endX - 6},${endY - 4} ${endX - 6},${endY + 4}`);
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
                                    let currentDep = targetPeriod.dep ? targetPeriod.dep.toString().split(',').map(s => s.trim()) : [];
                                    if (!currentDep.includes(depPid)) {
                                        if (window.isCircularDependency && window.isCircularDependency(depPid, targetPeriodId)) {
                                            alert('循環依存になるため接続できません。');
                                            // エラー時は元の接続に戻す
                                            oldDeps.push(depPid);
                                            sourcePeriod.dep = oldDeps.join(', ');
                                        } else {
                                            currentDep.push(depPid);
                                            targetPeriod.dep = currentDep.join(', ');
                                        }
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

// 他ファイルから window 経由で呼べるよう公開
window.renderAll = renderAll;
window.renderChart = renderChart;
