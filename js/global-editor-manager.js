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
            this.globalToolbarContainer.style.display = 'block'; 
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

        // カスタムのキーボードバインディング
        this.quill.keyboard.addBinding({ key: 's', shortKey: true }, (range, context) => {
            if (this.activeContainer !== this.defaultContainer) {
                this.closeEditor();
            } else {
                if (this.defaultOnSave) {
                    this.defaultOnSave(this.quill.root.innerHTML, this.quill.getContents());
                }
            }
            return false;
        });

        // エディタ外をクリックしたときにインラインエディタを閉じてドックに戻る処理
        document.addEventListener('mousedown', (e) => {
            if (!this.activeContainer || this.activeContainer === this.defaultContainer) return;

            // パレットやツールバー、エディタ本体、またはアクティブなコンテナそのものの中なら閉じない
            const isToolbar = e.target.closest('#global-quill-toolbar-container');
            const isEditor = this.singletonEditorWrap.contains(e.target);
            const isPickerOption = e.target.closest('.ql-picker-options') || e.target.closest('.ql-tooltip');
            const isColorPopup = e.target.closest('#color-palette-popup');
            const isInsideActiveContainer = this.activeContainer.contains(e.target);
            
            if (!isToolbar && !isEditor && !isPickerOption && !isColorPopup && !isInsideActiveContainer) {
                this.closeEditor();
            }
        });
        
        this.quill.on('text-change', (delta, oldDelta, source) => {
            if (source === 'user') {
                if (this.activeContainer === this.defaultContainer) {
                    if (this.defaultOnSave) {
                        this.defaultOnSave(this.quill.root.innerHTML, this.quill.getContents());
                    }
                }
            }
        });
        
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
                <select class="ql-size" title="文字サイズ">
                    <option value="small"></option>
                    <option selected></option>
                    <option value="large"></option>
                    <option value="huge"></option>
                </select>
                <button class="ql-bold" title="太字"></button>
                <select class="ql-color" title="文字色"></select>
                <select class="ql-background" title="文字背景色"></select>
                <button class="ql-align" value="" title="左揃え"></button>
                <button class="ql-align" value="center" title="中央揃え"></button>
                <button class="ql-align" value="right" title="右揃え"></button>
            </span>
            <span class="ql-formats" id="quill-box-styles" style="display: none; border-left: 1px solid #ccc; padding-left: 10px; margin-left: 10px; align-items: center;">
                <span style="font-size: 11px; color: #666; margin-right: 5px;">枠線:</span>
                <select id="ql-border-style" title="枠線の種類" style="width: 65px; height: 24px; font-size: 11px; padding: 0 2px;">
                    <option value="none">なし</option>
                    <option value="solid">実線</option>
                    <option value="dashed">破線</option>
                    <option value="dotted">点線</option>
                </select>
                <input type="number" id="ql-border-width" title="枠線の太さ" min="0" max="10" value="1" style="width: 35px; height: 24px; font-size: 11px; margin-left: 2px;">
                <input type="color" id="ql-border-color" title="枠線の色" style="width: 24px; height: 24px; padding: 0; border: 1px solid #ccc; margin-left: 2px; cursor: pointer; background: none;">
                <span style="font-size: 11px; color: #666; margin-left: 8px; margin-right: 5px;">背景:</span>
                <input type="color" id="ql-box-bg-color" title="ボックス背景色" style="width: 24px; height: 24px; padding: 0; border: 1px solid #ccc; cursor: pointer; background: none;">
                <button type="button" id="ql-box-bg-clear" title="背景色をクリア" style="width: 20px; height: 24px; padding: 0; border: none; background: none; cursor: pointer; font-size: 10px; color: #dc3545; display: flex; align-items: center; justify-content: center;">✖</button>
            </span>
            <span class="ql-formats">
                <button class="ql-clean" title="書式クリア"></button>
            </span>
        `;

        setTimeout(() => {
            const borderStyle = document.getElementById('ql-border-style');
            const borderWidth = document.getElementById('ql-border-width');
            const borderColor = document.getElementById('ql-border-color');
            const boxBgColor = document.getElementById('ql-box-bg-color');
            const boxBgClear = document.getElementById('ql-box-bg-clear');

            if (!borderStyle) return;

            const updateBoxStyles = () => {
                if (this.activeContainer && this.activeContainer.classList.contains('chart-text-box')) {
                    this.activeContainer.style.borderStyle = borderStyle.value;
                    this.activeContainer.style.borderWidth = borderWidth.value + 'px';
                    this.activeContainer.style.borderColor = borderColor.value;
                    this.activeContainer.style.backgroundColor = boxBgColor.dataset.isTransparent === 'true' ? 'transparent' : boxBgColor.value;
                }
            };

            [borderStyle, borderWidth, borderColor, boxBgColor].forEach(el => {
                el.addEventListener('change', () => {
                    if (el === boxBgColor) boxBgColor.dataset.isTransparent = 'false';
                    updateBoxStyles();
                });
            });

            boxBgClear.addEventListener('click', (e) => {
                e.stopPropagation();
                boxBgColor.dataset.isTransparent = 'true';
                if (this.activeContainer) {
                    this.activeContainer.style.backgroundColor = 'transparent';
                    updateBoxStyles();
                }
            });
        }, 0);
    }

    rgbToHex(rgb) {
        if (!rgb || rgb === 'transparent' || rgb === 'rgba(0, 0, 0, 0)') return '#ffffff';
        const match = rgb.match(/^rgb\((\d+),\s*(\d+),\s*(\d+)\)$/);
        if (!match) {
            if (rgb.startsWith('#')) return rgb;
            return '#ffffff';
        }
        return "#" + ("0" + parseInt(match[1]).toString(16)).slice(-2) +
                     ("0" + parseInt(match[2]).toString(16)).slice(-2) +
                     ("0" + parseInt(match[3]).toString(16)).slice(-2);
    }

    setDefaultContainer(container, onSave) {
        this.defaultContainer = container;
        this.defaultOnSave = onSave;
        this.dockToDefault();
    }
    
    /**
     * エディタを完全にDOMから切り離す（PDFエクスポート用）
     */
    detachEditor() {
        if (this.singletonEditorWrap.parentNode) {
            this.singletonEditorWrap.parentNode.removeChild(this.singletonEditorWrap);
        }
    }

    dockToDefault() {
        if (!this.defaultContainer) return;

        this.activeContainer = this.defaultContainer;
        this.onSaveCallback = this.defaultOnSave;

        this.singletonEditorWrap.classList.remove('vertical-editor');
        this.defaultContainer.innerHTML = '';
        this.defaultContainer.appendChild(this.singletonEditorWrap);

        const boxStyles = document.getElementById('quill-box-styles');
        if (boxStyles) boxStyles.style.display = 'none';
        
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

    openEditor(container, onSave, isVertical = false) {
        if (this.activeContainer === container) return;

        // もし直前まで別のインラインセルを開いていたなら、閉じる
        if (this.activeContainer && this.activeContainer !== this.defaultContainer) {
            this.closeEditor();
        }

        // デフォルト（備考欄）に居た場合は、今の備考欄データを保存
        if (this.activeContainer === this.defaultContainer && this.defaultOnSave) {
            const currentHtml = this.quill.root.innerHTML;
            this.defaultOnSave(currentHtml, this.quill.getContents());
            this.defaultContainer.innerHTML = `<div class="ql-editor ql-editor-content" style="padding:0;">${currentHtml}</div>`;
        }

        // 既存のHTMLを先に取得してからコンテナを空にする（非常に重要）
        const existingHtml = container.innerHTML;

        this.activeContainer = container;
        this.onSaveCallback = onSave;

        // ボックススタイルの表示切り替え
        const boxStyles = document.getElementById('quill-box-styles');
        if (boxStyles) {
            const isTextBox = container.classList.contains('chart-text-box');
            boxStyles.style.display = isTextBox ? 'inline-flex' : 'none';
            if (isTextBox) {
                document.getElementById('ql-border-style').value = container.style.borderStyle || 'solid';
                document.getElementById('ql-border-width').value = parseInt(container.style.borderWidth) || 1;
                document.getElementById('ql-border-color').value = this.rgbToHex(container.style.borderColor);
                const bg = container.style.backgroundColor;
                const boxBgColor = document.getElementById('ql-box-bg-color');
                boxBgColor.value = this.rgbToHex(bg);
                boxBgColor.dataset.isTransparent = (!bg || bg === 'transparent' || bg === 'rgba(0, 0, 0, 0)') ? 'true' : 'false';
            }
        }

        if (isVertical) {
            this.singletonEditorWrap.classList.add('vertical-editor');
        } else {
            this.singletonEditorWrap.classList.remove('vertical-editor');
        }

        // コンテンツの流し込み
        if (existingHtml.trim() === '' || existingHtml === '<p><br></p>' || existingHtml === '<br>') {
            this.quill.setText('');
        } else {
            this.quill.clipboard.dangerouslyPasteHTML(existingHtml);
        }
        this.quill.history.clear(); 

        // 移動と配置
        container.innerHTML = '';
        container.appendChild(this.singletonEditorWrap);

        // フォーカス
        setTimeout(() => {
            this.quill.focus();
            const length = this.quill.getLength();
            this.quill.setSelection(length, length);
        }, 50);
    }

    closeEditor() {
        if (!this.activeContainer || this.activeContainer === this.defaultContainer) return;

        const html = this.quill.root.innerHTML;
        const isEmpty = this.quill.getText().trim() === '';
        const finalHtml = isEmpty ? '' : html;

        const container = this.activeContainer;
        const callback = this.onSaveCallback;

        if (callback) {
            const boxStyles = container.classList.contains('chart-text-box') ? {
                borderStyle: container.style.borderStyle,
                borderWidth: parseInt(container.style.borderWidth) || 0,
                borderColor: container.style.borderColor,
                backgroundColor: container.style.backgroundColor
            } : null;
            callback(finalHtml, this.quill.getContents(), boxStyles);
        }

        container.innerHTML = finalHtml;
        this.dockToDefault();
    }
}

document.addEventListener('DOMContentLoaded', () => {
    window.editorManager = new GlobalEditorManager();
});
