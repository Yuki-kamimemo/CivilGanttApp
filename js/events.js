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
            window.renderChart();
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
                window.renderChart();
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

// linkingState は globals.js で宣言済み
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
                            if (window.isCircularDependency && window.isCircularDependency(sourcePid, targetPeriodId)) {
                                alert('循環依存になるため接続できません。（A→B→Aのようなループ等）');
                            } else {
                                if (currentDep) currentDep += ", " + sourcePid;
                                else currentDep = sourcePid;
                                window.handlePeriodChange(targetTaskId, targetPeriodId, 'dep', currentDep);
                            }
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

        resizer.addEventListener('mousedown', function (e) {
            startX = e.clientX;
            startWidth = col.offsetWidth;

            cols.forEach(c => c.style.width = c.offsetWidth + 'px');
            resizer.classList.add('resizing');

            const mouseMoveHandler = function (e) {
                const newWidth = startWidth + (e.clientX - startX);
                if (newWidth > 30) col.style.width = newWidth + 'px';
            };

            const mouseUpHandler = function () {
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
