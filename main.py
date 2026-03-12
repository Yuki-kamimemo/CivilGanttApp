# Civil Schedule Master (CivilGanttApp)
# Copyright (c) 2026 [Your Company Name]
# All rights reserved.
# This software is proprietary and confidential.

import webview
import os
import sys
import json

class Api:
    def __init__(self):
        self._window = None
        self._current_file_path = None
        self._initial_file_content = None # ★追加: 起動時に読み込むデータ
        self._is_dirty = False            # 未保存変更フラグ（Python側で管理）

    # ★追加: JavaScriptから初期データを受け取るための関数
    def get_initial_data(self):
        return self._initial_file_content

    # JSから変更を通知される（clean→dirty に切り替わる瞬間のみ呼ばれる）
    def notify_change(self):
        self._is_dirty = True
        return True

    def save_file(self, data_str, default_filename):
        if not self._window:
            return False

        file_types = ('Civil Schedule Master Files (*.csm)', 'All files (*.*)')
        result = self._window.create_file_dialog(
            webview.FileDialog.SAVE,
            directory='',
            save_filename=default_filename,
            file_types=file_types
        )

        if result:
            if isinstance(result, str):
                result = (result,)
            if len(result) > 0:
                file_path = result[0]
                try:
                    with open(file_path, 'w', encoding='utf-8') as f:
                        f.write(data_str)
                    self._current_file_path = file_path
                    self._is_dirty = False  # 保存成功でクリーン状態に
                    return True
                except Exception as e:
                    print(f"ファイルの保存に失敗しました: {e}")
                    if self._window:
                        error_msg = str(e).replace("'", "\\'").replace("\n", "\\n")
                        self._window.evaluate_js(f"alert('ファイルの保存に失敗しました:\\n{error_msg}')")
                    return False
        return False

    def open_file(self):
        if not self._window:
            return None

        file_types = ('Civil Schedule Master Files (*.csm)', 'All files (*.*)')
        result = self._window.create_file_dialog(
            webview.FileDialog.OPEN,
            allow_multiple=False,
            file_types=file_types
        )

        if result:
            if isinstance(result, str):
                result = (result,)
            if len(result) > 0:
                file_path = result[0]
                try:
                    with open(file_path, 'r', encoding='utf-8') as f:
                        file_content = f.read()
                    self._current_file_path = file_path
                    self._is_dirty = False  # ファイル読み込み後はクリーン状態に
                    return file_content
                except Exception as e:
                    print(f"ファイルの読み込みに失敗しました: {e}")
                    if self._window:
                        error_msg = str(e).replace("'", "\\'").replace("\n", "\\n")
                        self._window.evaluate_js(f"alert('ファイルの読み込みに失敗しました:\\n{error_msg}')")
                    return None
        return None

    def overwrite_file(self, data_str):
        if not self._current_file_path:
            return False

        try:
            with open(self._current_file_path, 'w', encoding='utf-8') as f:
                f.write(data_str)
            self._is_dirty = False  # 上書き保存成功でクリーン状態に
            return True
        except Exception as e:
            print(f"上書き保存に失敗しました: {e}")
            if self._window:
                error_msg = str(e).replace("'", "\\'").replace("\n", "\\n")
                self._window.evaluate_js(f"alert('上書き保存に失敗しました:\\n{error_msg}')")
            return False

    def clear_file_path(self):
        self._current_file_path = None
        self._is_dirty = False  # 新規作成はクリーン状態に
        return True

    def generate_pdf_from_html(self, html_content, settings, file_path):
        try:
            from playwright.sync_api import sync_playwright
            with sync_playwright() as p:
                browser = p.chromium.launch(headless=True)
                page = browser.new_page()
                page.set_content(html_content, wait_until='load')
                
                # Default to A4 Landscape if not specified
                paper_size = settings.get('paperSize', 'a4-landscape')
                format_name = 'A3' if 'a3' in paper_size.lower() else 'A4'
                
                # Default margins (can be overridden by settings if passed from JS)
                margin = {
                    "top": "10mm",
                    "bottom": "10mm",
                    "left": "10mm",
                    "right": "10mm"
                }

                page.pdf(
                    path=file_path,
                    format=format_name,
                    landscape=True, # Always landscape for Gantt
                    print_background=True,
                    margin=margin
                )
                browser.close()
            return True
        except Exception as e:
            print(f"Playwright PDF generation failed: {e}")
            raise e

    def save_pdf_file(self, html_content, settings, default_filename):
        """JS からプレビュー用のHTML文字列と設定を受け取り、PlaywrightでPDF化して保存"""
        if not self._window:
            return False

        file_types = ('PDF Files (*.pdf)', 'All files (*.*)')
        result = self._window.create_file_dialog(
            webview.FileDialog.SAVE,
            directory='',
            save_filename=default_filename,
            file_types=file_types
        )

        if result:
            if isinstance(result, str):
                result = (result,)
            if len(result) > 0:
                file_path = result[0]
                try:
                    # HTML文字列をPlaywrightに渡してPDF生成
                    self.generate_pdf_from_html(html_content, settings, file_path)
                    return True
                except Exception as e:
                    print(f"PDF保存に失敗しました: {e}")
                    if self._window:
                        error_msg = str(e).replace("'", "\\'").replace("\n", "\\n")
                        self._window.evaluate_js(f"alert('PDFの保存に失敗しました:\\n{error_msg}')")
                    return False
        return False

    def save_png_file(self, html_content, default_filename):
        """JS からHTML文字列を受け取り、PlaywrightでPNG化して保存"""
        if not self._window:
            return False

        file_types = ('PNG Files (*.png)', 'All files (*.*)')
        result = self._window.create_file_dialog(
            webview.FileDialog.SAVE,
            directory='',
            save_filename=default_filename,
            file_types=file_types
        )

        if result:
            if isinstance(result, str):
                result = (result,)
            if len(result) > 0:
                file_path = result[0]
                try:
                    from playwright.sync_api import sync_playwright
                    import tempfile, os
                    with tempfile.NamedTemporaryFile(
                        suffix='.html', delete=False, mode='w', encoding='utf-8'
                    ) as f:
                        f.write(html_content)
                        tmp_path = f.name
                    try:
                        with sync_playwright() as p:
                            browser = p.chromium.launch(headless=True)
                            page = browser.new_page()
                            page.set_viewport_size({'width': 2480, 'height': 3508})
                            page.goto(f'file:///{tmp_path}', wait_until='load')
                            page.screenshot(path=file_path, full_page=True)
                            browser.close()
                    finally:
                        os.unlink(tmp_path)
                    return True
                except Exception as e:
                    print(f"PNG保存に失敗しました: {e}")
                    if self._window:
                        error_msg = str(e).replace("'", "\\'").replace("\n", "\\n")
                        self._window.evaluate_js(f"alert('PNGの保存に失敗しました:\\n{error_msg}')")
                    return False
        return False

    def save_image_file(self, data_uri, default_filename):
        """JS から base64 PNG data URI を受け取りファイル保存"""
        if not self._window:
            return False
        import base64

        file_types = ('PNG Files (*.png)', 'All files (*.*)')
        result = self._window.create_file_dialog(
            webview.FileDialog.SAVE,
            directory='',
            save_filename=default_filename,
            file_types=file_types
        )

        if result:
            if isinstance(result, str):
                result = (result,)
            if len(result) > 0:
                file_path = result[0]
                try:
                    if ',' in data_uri:
                        b64_data = data_uri.split(',', 1)[1]
                    else:
                        b64_data = data_uri
                    img_bytes = base64.b64decode(b64_data)
                    with open(file_path, 'wb') as f:
                        f.write(img_bytes)
                    return True
                except Exception as e:
                    print(f"PNG保存に失敗しました: {e}")
                    if self._window:
                        error_msg = str(e).replace("'", "\\'").replace("\n", "\\n")
                        self._window.evaluate_js(f"alert('PNGの保存に失敗しました:\\n{error_msg}')")
                    return False
        return False

    # ★改修：実務レベルのExcel出力機能（すべての矢印描画・クリティカルパス赤線化・シート分割）
    def export_to_excel(self, data_str, default_filename):
        if not self._window:
            return False
            
        file_types = ('Excel Files (*.xlsx)', 'All files (*.*)')
        result = self._window.create_file_dialog(
            webview.FileDialog.SAVE, 
            directory='', 
            save_filename=default_filename, 
            file_types=file_types
        )
        
        if result:
            if isinstance(result, str):
                result = (result,)
            if len(result) > 0:
                file_path = result[0]
                try:
                    import excel_exporter
                    excel_exporter.export_data_to_excel(data_str, file_path)
                    return True
                except Exception as e:
                    print(f"Excel出力プロセスでエラーが発生しました: {e}")
                    if self._window:
                        error_msg = str(e).replace("'", "\\'").replace("\n", "\\n")
                        self._window.evaluate_js(f"alert('Excel出力に失敗しました:\\n{error_msg}')")
                    return False
        return False

    def open_manual(self):
        """別ウィンドウで操作マニュアルを表示する"""
        manual_path = get_resource_path('manual.html')
        if os.path.exists(manual_path):
            webview.create_window(
                'Civil Schedule Master - 操作マニュアル',
                url=manual_path,
                width=1000,
                height=800
            )
        else:
            if self._window:
                self._window.evaluate_js("alert('マニュアルファイル(manual.html)が見つかりません。')")

def get_resource_path(relative_path):
    try:
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.dirname(os.path.abspath(__file__))
    return os.path.join(base_path, relative_path)

def start_app():
    api = Api()
    # ★追加: ダブルクリックなどでファイルが渡された場合の処理
    if len(sys.argv) > 1:
        initial_file = sys.argv[1]
        if os.path.exists(initial_file):
            try:
                with open(initial_file, 'r', encoding='utf-8') as f:
                    api._initial_file_content = f.read()
                    api._current_file_path = initial_file
            except Exception as e:
                print(f"初期ファイルの読み込みに失敗しました: {e}")
    html_path = get_resource_path('index.html')

    window = webview.create_window(
        'Civil Schedule Master',
        url=html_path,
        width=1200,
        height=800,
        min_size=(1000, 700),
        js_api=api
    )
    api._window = window

    def on_closing():
        # evaluate_js はこのスレッドから呼ぶとデッドロックするため
        # Python側の _is_dirty フラグを直接参照する
        if api._is_dirty:
            import ctypes
            MB_YESNO       = 0x00000004
            MB_ICONWARNING = 0x00000030
            MB_TOPMOST     = 0x00040000
            IDYES = 6
            user32 = ctypes.WinDLL('user32', use_last_error=True)  # type: ignore[attr-defined]
            result = user32.MessageBoxW(
                0,
                '保存されていない変更があります。\n終了してもよろしいですか？',
                '終了の確認',
                MB_YESNO | MB_ICONWARNING | MB_TOPMOST
            )
            return result == IDYES
        return True

    window.events.closing += on_closing
    webview.start(debug=False)

if __name__ == '__main__':
    start_app()