# Yomitoku OCR & EPUB Workflow 統合ツール v5.9（単体 .py）

PDFを画像化して **Yomitoku** でOCRし、Markdown（+ページ画像）として出力します。  
さらに、出力したMarkdownの **分割 /（安全分割）/ 結合 / EPUB化** をGUIでまとめて行える、Windows向けのワークフローツールです。

参考（使い方の補足・背景）：
- note記事：<https://note.com/miya_bee_note/n/n6c4710d3a313>

---

## 動作環境
- Windows 10/11 推奨
- Python 3.10+ 推奨（3.9でも動く可能性はあります）
- GPU利用時：CUDA対応のPyTorch環境（任意）

### GPUで使う場合（任意）
CUDA対応のPyTorchを先に用意した上で、次を実行します。

~~~bash
pip install -U "yomitoku[gpu]" onnxruntime-gpu
~~~

---

## できること（タブ構成）

### Tab1：OCR（PDF → Markdown）
- PDFをページ単位で画像化 → YomitokuでOCR → `output.md` を生成
- 各ページ画像は `assets/` に出力（Markdownから参照リンク）
- 上下トリミング（％指定）＋「ビジュアル範囲指定」で実ページを見ながら調整可能
- CUDAメモリ不足などの場合、状況によりCPUへ切り替えて続行することがあります
- 暗号化（パスワード保護）PDFの疑いがある場合はエラー表示します

**出力例**
- `output.md`（本文 + ページ画像リンク）
- `assets/`（各ページ画像）

### Tab2：MD分割（通常分割）
- 大きいMarkdownを校正/編集しやすいサイズに分割
- 見出しレベル（1〜6）を境界として扱えます（例：レベル1なら `# ` のみで分割）
- 「テスト（サイズ確認）」で、書き出しなしに分割結果（サイズ/行数/文字数/既存ファイル有無）を一覧表示
- JSON書き出しでテスト結果一覧を保存可能
- 「分割実行」で `{name_root}_01.md ...` のように書き出し（上書き確認あり）

### Tab2-2：安全分割（目次誤分割対策 + “選択を結合”）
- 「本文開始位置を推定」し、本文開始以降の `# 1<br>` などを章境界として扱う **目次誤分割対策版**
- 実行時は入力MDと同じフォルダに `split_output_safe/実行時刻/` を作って出力
- 分割結果から **選択した項目を結合**して、狙った単位にまとめ直すことができます

### 想定ワークフロー：Tab2で出力 → 生成AIで校正 → Tab3以降で続き
本ツール自体に「校正」タブはありません。  
**Tab2/Tab2-2で分割したMDを外部の生成AI等で校正**し、校正後のMDを **Tab3〜Tab4** で続けて処理する想定です。

### Tab3：MD結合（クリップボード監視で集約）
- クリップボードを監視し、コピーした本文をスタックに積む
- 順序調整・選択編集・プレビューが可能
- 「結合して保存」で1つのMarkdownとして保存

> ※クリップボード監視を使うため、機密情報のコピーには注意してください。

### Tab4：EPUB化（Markdown → EPUB）
- MarkdownをHTML化してEPUBを書き出し
- 入力MD / 出力EPUB / タイトル / 著者を指定して作成

### Tab5：その他
- Send to Kindle を開く
- ログクリア

---

## 事前準備（重要）：Poppler（PDF画像化に必須）
本ツールは `pdf2image` を使用してPDFを画像化します。Windowsでは **Poppler** が必要です。

### 方式A：Popplerを `C:\poppler\Library\bin` に配置（推奨）
このツールは `C:\poppler\Library\bin` が存在する場合、Popplerパスに自動で初期値として入ります。  
（違う場所でもOK。その場合はTab1で「Popplerパス」を指定してください）

### 方式B：任意の場所に配置して、Tab1でパス指定
Tab1の「Popplerパス」に、`pdftoppm.exe` 等が入っている `bin` フォルダを指定してください。

---

## インストール（Python依存ライブラリ）

### 1) 仮想環境（推奨）
~~~bash
py -m venv .venv
.venv\Scripts\activate
python -m pip install -U pip
~~~

### 2) 依存ライブラリのインストール
~~~bash
pip install -r requirements.txt
~~~

補足：
- 起動時に依存ライブラリ不足を検出した場合、ダイアログで `pip install ...` の案内が出ます。

---

## 起動方法
~~~bash
python app.py
~~~

---

## クイックスタート（最短）
1. Poppler を用意（`C:\poppler\Library\bin` が推奨）
2. `pip install -r requirements.txt`
3. `python app.py`
4. Tab1でPDFを選択 → OCR実行 → `output.md` を生成

---

## よくあるトラブルシュート

### Tab1でPDFが画像化できない / OCRが進まない
- Popplerパスが正しいか確認してください（`bin` 配下に `pdftoppm.exe` 等があること）
- 暗号化PDFの可能性がある場合は、解除するか別PDFで試してください

### GPUを使いたい
- CUDA対応のPyTorchを先に整備してください
- 環境によっては `yomitoku[gpu]` / `onnxruntime-gpu` が必要です
- GPUメモリ不足時はCPUに切り替えて続行することがあります

---

## セキュリティ/プライバシー注意
- Tab3のクリップボード監視は、コピーしたテキストをスタックに取り込みます。  
  機密情報（パスワード、個人情報など）をコピーする作業と併用しないことを推奨します。

---

## ライセンス
- MIT License
