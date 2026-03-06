# ヘッドレスブラウザによるPDFエクスポート機能 実装計画書

## 目的
現在のJavaScriptベースのPDFエクスポート機能（`html2canvas` + `jsPDF`）を全面廃止し、ヘッドレスブラウザ（Playwright）を使用した強固なPythonバックエンド方式へと刷新します。この変更は、複雑なCSSレイアウト、特に「縦書き（`writing-mode: vertical-rl`）」を完璧にレンダリングするという目的のために必要不可欠な措置です。

新しいアーキテクチャでは以下のことを行います：
1. フロントエンド（JS）が、ガントチャートと日次備考欄を表現する「印刷に最適化された純粋なHTML文字列」を生成します。
2. その生成されたHTML文字列を、モーダル内の `iframe` を使用して「100%正確なプレビュー」としてリアルタイムに表示します。
3. HTML文字列をPythonバックエンドに送信します。
4. Playwright（Chromiumエンジン）を使用してそのHTMLをレンダリングし、ピクセル単位で完璧なPDFとして保存します。

## ユーザーの確認が必要な事項（重要）
> [!IMPORTANT]
> **依存ファイルの増加とサイズ肥大化**: このアプローチでは、プロジェクトに新しく `playwright` を追加する必要があります。`PyInstaller` を使用して実行可能ファイル（`.exe`）をビルドする際、内部にChromiumブラウザエンジンをバンドル（またはダウンロード処理）しなければならないため、**出来上がるexeファイルの容量が大幅（100MB〜200MBほど）に増加する**可能性があります。
> 
> **「完璧なPDFレンダリング」と引き換えに、実行可能ファイルのサイズが増加してしまうことは許容可能でしょうか？**

## 変更予定のファイルと内容

### バックエンド（Python）
#### [MODIFY] [main.py](file:///c:/Users/Y-Katsuta/Documents/CivilGanttApp/main.py)
- `playwright.sync_api` をインポートします。
- `Api` クラスに新しいメソッド `generate_pdf_from_html(self, html_content, settings, output_path)` を追加します。
- このメソッドの処理内容：
  1. ヘッドレス状態のChromiumインスタンスを起動。
  2. 新しいページを作成し、`page.set_content(html_content)` でフロントからのHTMLを読み込ませる。
  3. `page.pdf()` を呼び出し、A3横などの用紙サイズや余白設定を適用して保存する。
  4. 成功ステータス、またはエラーメッセージを返す。
- 既存の `save_pdf_file` （または新規エンドポイント）を改修し、Data URIではなく「HTML文字列と設定値」を受け取るように変更します。

#### [MODIFY] [requirements.txt](file:///c:/Users/Y-Katsuta/Documents/CivilGanttApp/requirements.txt)
- `playwright` を追加します（`pip freeze`にて実行済み）。

### フロントエンド（JavaScript & UI）
#### [MODIFY] [js/pdf-export.js](file:///c:/Users/Y-Katsuta/Documents/CivilGanttApp/js/pdf-export.js)
- **削除**: これまでの `html2canvas` および `jsPDF` のすべてのロジックを削除します。
- **追加**: `buildPrintableHtml(settings)` 関数。現在のガントチャートのDOMを複製・抽出・成形し、必要なCSSスタイルを直接 `<style>` タグとして注入した上で、独立したHTMLドキュメント文字列にする処理を実装します。
- **追加**: `updateHtmlPreview()` 関数。`buildPrintableHtml` の出力結果を、プレビューモーダル内の `iframe` の `srcdoc` 属性にセットします。これにより、完全に正確なプレビューが実現します。
- **変更**: `executePdfSave()` 関数。ブラウザ内でPDFデータを生成するのではなく、生成したHTML文字列とレイアウト設定（用紙サイズ、余白など）を `window.pywebview.api.generate_pdf_from_html` に送信するように変更します。

#### [MODIFY] [index.html](file:///c:/Users/Y-Katsuta/Documents/CivilGanttApp/index.html)
- PDFプレビューモーダルのエリア（`#pdf-preview-area`）を更新します。
- `<canvas>` ベースのプレビュー要素を削除し、代わりに `<iframe id="pdf-preview-iframe" style="width:100%; height:100%; border:1px solid #ccc;"></iframe>` を配置します。

## 検証計画

### 手動による確認テスト
1. アプリケーションを起動します。
2. サンプルプロジェクト（`sample_道路改良工事.csm`）をロードします。
3. 「印刷 / エクスポート」ボタンをクリックします。
4. 新しいiframeベースのプレビューが正しく読み込まれ、日次備考欄の「縦書き」を含め、メイン画面のUIと完璧に一致しているか確認します。
5. 余白や縮尺などの印刷設定を変更し、プレビュー画面がリアルタイムで連動して更新されるか確認します。
6. 「PDFとして保存」をクリックし、PCに保存します。
7. 生成されたPDFファイルを標準的なビューア（Adobe Acrobat, Chromeブラウザ, Edgeなど）で開きます。
8. **重要チェック項目:**
   - フォント名や色、線の太さなどは正確に反映されているか？
   - 縦書きテキスト（`writing-mode: vertical-rl`）が、文字の回転がおかしくなることなく完璧にレンダリングされているか？
   - 矢印（依存関係の線）がタスクバーと正確につながっているか？
   - チャートが横に長い場合、改ページが綺麗に行われているか？
