import webview
import os
import sys
import json

class Api:
    def __init__(self):
        self._window = None
        self._current_file_path = None 
        self._initial_file_content = None # ★追加: 起動時に読み込むデータ
        
    # ★追加: JavaScriptから初期データを受け取るための関数
    def get_initial_data(self):
        return self._initial_file_content

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
            return True
        except Exception as e:
            print(f"上書き保存に失敗しました: {e}")
            if self._window:
                error_msg = str(e).replace("'", "\\'").replace("\n", "\\n")
                self._window.evaluate_js(f"alert('上書き保存に失敗しました:\\n{error_msg}')")
            return False

    def clear_file_path(self):
        self._current_file_path = None
        return True

    def save_pdf_file(self, data_uri, default_filename):
        """JS から base64 PDF data URI を受け取りファイル保存"""
        if not self._window:
            return False
        import base64

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
                    # "data:application/pdf;base64," プレフィックスを除去
                    if ',' in data_uri:
                        b64_data = data_uri.split(',', 1)[1]
                    else:
                        b64_data = data_uri
                    pdf_bytes = base64.b64decode(b64_data)
                    with open(file_path, 'wb') as f:
                        f.write(pdf_bytes)
                    return True
                except Exception as e:
                    print(f"PDF保存に失敗しました: {e}")
                    if self._window:
                        error_msg = str(e).replace("'", "\\'").replace("\n", "\\n")
                        self._window.evaluate_js(f"alert('PDFの保存に失敗しました:\\n{error_msg}')")
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
    webview.start(debug=True)

if __name__ == '__main__':
    start_app()