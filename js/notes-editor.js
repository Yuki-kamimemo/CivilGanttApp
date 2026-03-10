let notesSaveTimer = null;

window.initNotesEditor = function () {
    const container = document.getElementById('project-notes-editor-container');
    if (!container) return;

    if (window.editorManager) {
        window.editorManager.setDefaultContainer(container, (html, delta) => {
            clearTimeout(notesSaveTimer);
            notesSaveTimer = setTimeout(() => {
                if (typeof state !== 'undefined') {
                    state.notesDelta = delta;
                    state.notes = html;
                    window.saveStateToHistory();
                    window.renderChart();
                }
            }, 500);
        });
    }
};

window.applyNotesToEditor = function () {
    const container = document.getElementById('project-notes-editor-container');
    if (!container) return;

    if (window.editorManager && window.editorManager.defaultContainer === container) {
        if (window.editorManager.activeContainer === container) {
            if (state.notesDelta) {
                window.editorManager.quill.setContents(state.notesDelta, 'api');
            } else if (state.notes) {
                window.editorManager.quill.clipboard.dangerouslyPasteHTML(state.notes);
                window.editorManager.quill.history.clear();
            } else {
                window.editorManager.quill.setText('');
            }
        } else {
            container.innerHTML = state.notes || '';
        }
    }
};

document.addEventListener('DOMContentLoaded', () => {
    // global-editor-manager のあとにロードされるため初期化
    setTimeout(() => {
        window.initNotesEditor();
    }, 100);
});
