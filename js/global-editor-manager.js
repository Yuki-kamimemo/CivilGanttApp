// ---------------------------------------------------
// グローバル Quill ツールバー管理および シングルトンエディタ
// （備考欄をデフォルトドックとし、日報クリック時だけ一時的に移動する）
// ---------------------------------------------------

class GlobalEditorManager {
    constructor() {
        this.activeContainer = null;
        this.onSaveCallback = null;
        
        this.defaultContainer = null;
        this.defaultOnSave = null;
        
        this.globalToolbarContainer = document.getElementById('global-quill-toolbar-container');
        if (this.globalToolbarContainer) {
            this.globalToolbarContainer.style.display = 'block'; // 常に表示する
        }
        
        this.singletonEditorWrap = document.createElement('div');
        this.singletonEditorWrap.id = 'singleton-quill-editor';
        this.singletonEditorWrap.style.width = '100%';
        this.singletonEditorWrap.style.height = '100%';
        this.singletonEditorWrap.style.boxSizing = 'border-box';
        
        this.setupPristineToolbar();

        this.quill = new Quill(this.singletonEditorWrap, {
            theme: 'snow',
            modules: {
                toolbar: '#global-quill-toolbar',
                history: { delay: 1000, maxStack: 100, userOnly: true }
            }
        });

        // カスタムのキーボードバインディング (Ctrl+S 無効化など)
        this.quill.keyboard.addBinding({ key: 's', shortKey: true }, (range, context) => {
            if (this.activeContainer !== this.defaultContainer) {
                this.closeEditor(); // 直感的UI: インライン編集時は Ctrl+S で保存して閉じる
            } else {
                // デフォルト（備考欄）の場合、入力状態を保存する
                if (this.defaultOnSave) {
                    this.defaultOnSave(this.quill.root.innerHTML, this.quill.getContents());
                }
            }
            return false;
        });

        // エディタ外をクリックしたときにインラインエディタを閉じてドックに戻る処理
        document.addEventListener('mousedown', (e) => {
            if (this.activeContainer && this.activeContainer !== this.defaultContainer) {
                const isToolbar = e.target.closest('#global-quill-toolbar-container');
                const isEditor = this.singletonEditorWrap.contains(e.target);
                const isPickerOption = e.target.closest('.ql-picker-options') || e.target.closest('.ql-tooltip');
                
                // mousedown時にDOMが書き換わってe.targetがロストする現象（クリック吸収）を防ぐため
                // 呼び出し元の mousedown イベントで e.stopPropagation() されている前提だが、
                // 念のため包含チェックを行う
                if (!isToolbar && !isEditor && !isPickerOption) {
                    this.closeEditor();
                }
            }
        });
        
        // テキストが変更されるたび、自動保存を発火させる
        this.quill.on('text-change', (delta, oldDelta, source) => {
            if (source === 'user') {
                if (this.activeContainer === this.defaultContainer) {
                    if (this.defaultOnSave) { // 備考欄のリアルタイム保存
                        this.defaultOnSave(this.quill.root.innerHTML, this.quill.getContents());
                    }
                }
            }
        });
        
        // セレクションが変更されたとき、現在の対象を明確にするためのフックを呼び出す
        this.quill.on('selection-change', range => {
            if (range && this.activeContainer === this.defaultContainer) {
                if (window.selectInput) window.selectInput('notes');
            }
        });
    }

    setupPristineToolbar() {
        const toolbarDOM = document.getElementById('global-quill-toolbar');
        if (!toolbarDOM) return;
        toolbarDOM.innerHTML = `
            <span class="ql-formats">
                <button class="ql-bold" title="太字"></button>
                <button class="ql-italic" title="斜体"></button>
                <button class="ql-underline" title="下線"></button>
                <button class="ql-strike" title="取り消し線"></button>
            </span>
            <span class="ql-formats">
                <button class="ql-list" value="ordered" title="番号付きリスト"></button>
                <button class="ql-list" value="bullet" title="箇条書きリスト"></button>
            </span>
            <span class="ql-formats">
                <select class="ql-size" title="文字サイズ">
                    <option value="small"></option>
                    <option selected></option>
                    <option value="large"></option>
                    <option value="huge"></option>
                </select>
            </span>
            <span class="ql-formats">
                <select class="ql-color" title="文字色"></select>
                <select class="ql-background" title="背景色"></select>
            </span>
            <span class="ql-formats">
                <select class="ql-align" title="揃え"></select>
            </span>
            <span class="ql-formats">
                <button class="ql-clean" title="書式クリア"></button>
            </span>
        `;
    }

    /**
     * ベースとなるデフォルトコンテナ（全体備考欄）を指定
     */
    setDefaultContainer(container, onSave) {
        this.defaultContainer = container;
        this.defaultOnSave = onSave;
        
        // 初期状態として、ドックに戻る
        this.dockToDefault();
    }
    
    /**
     * エディタをデフォルトコンテナにドック（収納）する
     */
    dockToDefault() {
        if (!this.defaultContainer) return;

        this.activeContainer = this.defaultContainer;
        this.onSaveCallback = this.defaultOnSave;

        this.singletonEditorWrap.classList.remove('vertical-editor');
        this.defaultContainer.innerHTML = '';
        this.defaultContainer.appendChild(this.singletonEditorWrap);
        
        // 外部（Stateなど）の初期HTMLを読み込む
        if (typeof state !== 'undefined') {
            if (state.notesDelta) {
                this.quill.setContents(state.notesDelta, 'api');
            } else if (state.notes) {
                this.quill.clipboard.dangerouslyPasteHTML(state.notes);
                this.quill.history.clear();
            } else {
                this.quill.setText('');
            }
        }
    }

    /**
     * 指定されたコンテナでインラインQuillエディタを開始する（日報用）
     */
    openEditor(container, onSave, isVertical = false) {
        if (this.activeContainer === container) return;

        // もし直前まで別のインラインセルを開いていたなら、閉じる
        if (this.activeContainer && this.activeContainer !== this.defaultContainer) {
            this.closeEditor();
        }

        // デフォルト（備考欄）に居た場合は、今の備考欄データが安全にStateに反映されているか保存
        if (this.activeContainer === this.defaultContainer && this.defaultOnSave) {
            const currentHtml = this.quill.root.innerHTML;
            this.defaultOnSave(currentHtml, this.quill.getContents());
            
            // ★追加: ドック（右側備考欄）から一時的にエディタが出張する際、
            // 空のコンテナに現在のHTMLを流し込んでおくことで、文字が消えたように見える現象を防ぎます。
            this.defaultContainer.innerHTML = `<div class="ql-editor ql-editor-content" style="padding:0;">${currentHtml}</div>`;
        }

        this.activeContainer = container;
        this.onSaveCallback = onSave;

        if (isVertical) {
            this.singletonEditorWrap.classList.add('vertical-editor');
        } else {
            this.singletonEditorWrap.classList.remove('vertical-editor');
        }

        const existingHtml = container.innerHTML;
        if (existingHtml.trim() === '' || existingHtml === '<p><br></p>' || existingHtml.includes('<br>')) {
             // 単なる改行のみ等なら空にする
            if (!existingHtml.trim().replace(/<br>/g, '') && !existingHtml.includes('<p>')) {
                this.quill.setText('');
            } else {
                this.quill.clipboard.dangerouslyPasteHTML(existingHtml);
                this.quill.history.clear(); 
            }
        } else {
            this.quill.clipboard.dangerouslyPasteHTML(existingHtml);
            this.quill.history.clear(); 
        }

        // 移動
        container.innerHTML = '';
        container.appendChild(this.singletonEditorWrap);

        // フォーカス
        setTimeout(() => {
            this.quill.focus();
            const length = this.quill.getLength();
            this.quill.setSelection(length, length);
        }, 50);
    }

    /**
     * インライン編集を終了し、結果を保存して、再びデフォルトドックに帰還する
     */
    closeEditor() {
        if (!this.activeContainer || this.activeContainer === this.defaultContainer) return;

        const html = this.quill.root.innerHTML;
        const isEmpty = this.quill.getText().trim() === '';
        const finalHtml = isEmpty ? '' : html;

        const container = this.activeContainer;
        const callback = this.onSaveCallback;

        // コールバック関数を実行して状態保存
        if (callback) {
            callback(finalHtml, this.quill.getContents());
        }

        // 抜け殻となったコンテナに結果を取り残す
        container.innerHTML = finalHtml;

        // デフォルトのドックに戻還する
        this.dockToDefault();
    }
}

document.addEventListener('DOMContentLoaded', () => {
    window.editorManager = new GlobalEditorManager();
});
