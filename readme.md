
# Yomitoku OCR & EPUB Workflow 統合ツール v5.9（単体 .py）

PDFを画像化して **Yomitoku** でOCRし、Markdown（+ページ画像）として出力します。  
さらに、出力したMarkdownの **分割 /（安全分割）/ 結合 / EPUB化** をGUIでまとめて行える、Windows向けのワークフローツールです。
以下のブログ(note)も参照してください
https://note.com/miya_bee_note/n/n6c4710d3a313

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

### Tab3：MD結合（クリップボード監視で集約）
- クリップボードを監視し、コピーした本文をスタックに積む
- 順序調整・選択編集・プレビューが可能
- 「結合して保存」で1つのMarkdownとして保存

※クリップボード監視を使うため、機密情報のコピーには注意してください。

### Tab4：EPUB化（Markdown → EPUB）
- MarkdownをHTML化してEPUBを書き出し
- 入力MD / 出力EPUB / タイトル / 著者を指定して作成

### Tab5：その他
- Send to Kindle を開く
- ログクリア

---

## 動作環境
- Windows 10/11 推奨
- Python 3.10+ 推奨（3.9でも動く可能性はあります）
- GPU利用時：CUDA対応のPyTorch環境（任意）

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
```bash
py -m venv .venv
.venv\Scripts\activate
python -m pip install -U pip
