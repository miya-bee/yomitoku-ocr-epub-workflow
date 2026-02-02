import sys
import tkinter as tk
from tkinter import filedialog, messagebox, scrolledtext, ttk
import threading
import os
import time
import traceback
import gc
import re
import math
import subprocess
import webbrowser
import hashlib
import unicodedata  # 正規化用
import importlib
import json  # JSON書き出し用

# ==========================================
# 0. ライブラリ読み込み (堅牢な環境判定版)
# ==========================================
MISSING_LIBS = []
import_error_detail = ""


def get_yomitoku_requirements():
    """環境に合わせて必要なyomitoku関連パッケージを返す"""
    try:
        import torch
        if torch.cuda.is_available():
            return ["yomitoku[gpu]", "onnxruntime-gpu"]
    except Exception:
        pass
    return ["yomitoku", "onnxruntime"]


def safe_import(module_name, pip_packages):
    """
    pip_packages: 文字列またはリストを受け取り、不足があればMISSING_LIBSに追加
    """
    try:
        return importlib.import_module(module_name)
    except ImportError:
        pkgs = [pip_packages] if isinstance(pip_packages, str) else pip_packages
        for pkg in pkgs:
            if pkg not in MISSING_LIBS:
                MISSING_LIBS.append(pkg)
        return None
    except Exception as e:
        global import_error_detail
        import_error_detail += f"\n[{module_name}]: {str(e)}"
        return None


cv2 = safe_import("cv2", "opencv-python")
np = safe_import("numpy", "numpy")
pdf2image = safe_import("pdf2image", "pdf2image")
ebooklib = safe_import("ebooklib", "EbookLib")
markdown_lib = safe_import("markdown", "markdown")
pyperclip = safe_import("pyperclip", "pyperclip")
torch = safe_import("torch", "torch")
PIL = safe_import("PIL", "Pillow")

yomitoku_pkgs = get_yomitoku_requirements()
yomitoku = safe_import("yomitoku", yomitoku_pkgs)

if not MISSING_LIBS:
    try:
        from pdf2image import convert_from_path, pdfinfo_from_path
        from yomitoku import DocumentAnalyzer
        from ebooklib import epub
        from PIL import Image, ImageTk
    except Exception as e:
        import_error_detail += f"\n[Sub-module Import]: {str(e)}"
        for lib in ["pdf2image", "EbookLib", "Pillow"] + yomitoku_pkgs:
            if lib not in MISSING_LIBS:
                MISSING_LIBS.append(lib)

if MISSING_LIBS:
    root = tk.Tk()
    root.withdraw()

    libs_unique = sorted(list(set(MISSING_LIBS)))
    install_cmd = f"pip install " + " ".join(libs_unique)

    error_message = (
        "【環境エラー】必要なライブラリが不足しています。\n\n"
        "■ 以下のコマンドをコピーして実行してください:\n"
        f"{install_cmd}\n\n"
        "※ GPUを使用する場合はCUDAの設定が必要です。\n"
        f"▼ 詳細:\n{import_error_detail}"
    )
    messagebox.showerror("起動失敗", error_message)
    sys.exit(1)

MAX_LOG_LINES = 1000


# Popplerのデフォルトパス（Windows想定）
DEFAULT_POPPLER_PATH = r"C:\poppler\Library\bin"

# ==========================================
# ビジュアルクロップダイアログ
# ==========================================
class VisualCropDialog(tk.Toplevel):
    def __init__(self, parent, pdf_path, poppler_path, current_top_pct, current_bottom_pct, callback):
        super().__init__(parent)
        self.title("上下カット範囲の指定")
        self.geometry("600x950")

        self.lift()
        self.focus_force()
        self.grab_set()

        self.callback = callback
        self.pdf_path = pdf_path
        self.poppler_path = poppler_path

        self.current_page = 1
        self.total_pages = 1

        self.top_pct = current_top_pct
        self.bottom_pct = current_bottom_pct

        self.img_h = 0
        self.img_w = 0
        self.tk_img = None
        self.offset_x = 0

        # 判定用（画像の実描画サイズ）
        self.preview_w = 0
        self.preview_h = 0
        self.offset_y = 0

        frame_top = tk.Frame(self, bg="#eee", pady=5)
        frame_top.pack(side="top", fill="x")

        tk.Label(
            frame_top,
            text="画像上をクリックすると、近い方（赤=上 / 青=下）の線がその位置へ移動します。ドラッグでも調整できます。",
            bg="#eee",
        ).pack()
        self.lbl_info = tk.Label(
            frame_top,
            text=f"上: {self.top_pct:.1f}%  下: {self.bottom_pct:.1f}%",
            font=("Meiryo", 12, "bold"),
            bg="#eee",
            fg="#333",
        )
        self.lbl_info.pack(pady=2)

        frame_bottom = tk.Frame(self, bg="#ddd", pady=10)
        frame_bottom.pack(side="bottom", fill="x")

        frame_page_ctrl = tk.Frame(frame_bottom, bg="#ddd")
        frame_page_ctrl.pack(pady=2)

        self.btn_prev = tk.Button(frame_page_ctrl, text="< 前", command=self.prev_page, width=8, state="disabled")
        self.btn_prev.pack(side="left", padx=5)

        self.lbl_page = tk.Label(frame_page_ctrl, text="読み込み中...", font=("Meiryo", 10, "bold"), bg="#ddd", width=16)
        self.lbl_page.pack(side="left", padx=5)

        self.btn_next = tk.Button(frame_page_ctrl, text="次 >", command=self.next_page, width=8, state="disabled")
        self.btn_next.pack(side="left", padx=5)

        frame_jump = tk.Frame(frame_bottom, bg="#ddd")
        frame_jump.pack(pady=5)

        tk.Label(frame_jump, text="指定ページへ:", bg="#ddd", font=("Meiryo", 9)).pack(side="left")
        self.entry_jump = tk.Entry(frame_jump, width=6, justify="center")
        self.entry_jump.pack(side="left", padx=2)
        self.entry_jump.bind("<Return>", lambda event: self.jump_to_page())

        btn_jump = tk.Button(frame_jump, text="移動", command=self.jump_to_page, bg="#fff", width=6, font=("Meiryo", 9))
        btn_jump.pack(side="left", padx=2)

        tk.Button(
            frame_bottom,
            text="これで決定",
            command=self.on_ok,
            bg="#C8E6C9",
            width=20,
            height=2,
            font=("Meiryo", 10, "bold"),
        ).pack(pady=10)

        self.canvas = tk.Canvas(self, bg="gray", cursor="crosshair")
        self.canvas.pack(fill="both", expand=True)

        self.canvas.bind("<Button-1>", self.on_click)
        self.canvas.bind("<B1-Motion>", self.on_drag)
        self.canvas.bind("<ButtonRelease-1>", self.on_release)

        self.drag_target = None

        self.after(200, self.init_pdf_info)

    def init_pdf_info(self):
        try:
            self.config(cursor="watch")
            self.update_idletasks()

            info = pdfinfo_from_path(self.pdf_path, poppler_path=self.poppler_path)
            self.total_pages = info["Pages"]
            self.update_page_label()
            self.load_preview()

        except Exception as e:
            messagebox.showerror("読み込みエラー", f"PDFを開けませんでした:\n{e}\n\nPopplerのパスやファイルを確認してください。")
            self.destroy()
        finally:
            self.config(cursor="")

    def update_page_label(self):
        self.lbl_page.config(text=f"Page {self.current_page} / {self.total_pages}")
        self.btn_prev.config(state="normal" if self.current_page > 1 else "disabled")
        self.btn_next.config(state="normal" if self.current_page < self.total_pages else "disabled")

    def prev_page(self):
        if self.current_page > 1:
            self.current_page -= 1
            self.update_page_label()
            self.load_preview()

    def next_page(self):
        if self.current_page < self.total_pages:
            self.current_page += 1
            self.update_page_label()
            self.load_preview()

    def jump_to_page(self):
        val = self.entry_jump.get().strip()
        if not val.isdigit():
            return
        p = int(val)
        if 1 <= p <= self.total_pages:
            self.current_page = p
            self.update_page_label()
            self.load_preview()
            self.entry_jump.delete(0, tk.END)
        else:
            messagebox.showwarning("範囲外", f"1 〜 {self.total_pages} の範囲で指定してください")

    def load_preview(self):
        try:
            self.canvas.delete("all")
            images = convert_from_path(
                self.pdf_path,
                first_page=self.current_page,
                last_page=self.current_page,
                poppler_path=self.poppler_path,
            )
            img = images[0]

            self.img_w, self.img_h = img.size

            canvas_w = self.canvas.winfo_width()
            canvas_h = self.canvas.winfo_height()
            if canvas_w <= 0 or canvas_h <= 0:
                self.after(100, self.load_preview)
                return

            scale = min(canvas_w / self.img_w, canvas_h / self.img_h)
            new_w = int(self.img_w * scale)
            new_h = int(self.img_h * scale)

            self.preview_w = new_w
            self.preview_h = new_h
            self.offset_y = 0  # 上詰め表示

            img_resized = img.resize((new_w, new_h), Image.LANCZOS)
            self.tk_img = ImageTk.PhotoImage(img_resized)

            self.offset_x = (canvas_w - new_w) // 2

            self.canvas.create_image(self.offset_x, self.offset_y, anchor="nw", image=self.tk_img)

            top_y = self.offset_y + int(new_h * (self.top_pct / 100.0))
            bottom_y = self.offset_y + int(new_h * (self.bottom_pct / 100.0))

            self.canvas.create_line(self.offset_x, top_y, self.offset_x + new_w, top_y, fill="red", width=2, tags="top")
            self.canvas.create_line(self.offset_x, bottom_y, self.offset_x + new_w, bottom_y, fill="blue", width=2, tags="bottom")

            self.lbl_info.config(text=f"上: {self.top_pct:.1f}%  下: {self.bottom_pct:.1f}%")

        except Exception as e:
            messagebox.showerror("プレビューエラー", f"ページ画像の表示に失敗しました:\n{e}")

    # ★改善：クリックは「近い方の線」を選び、その位置へ移動（=確実にマウス指定できる）
    def on_click(self, event):
        if self.preview_h <= 0 or self.preview_w <= 0:
            self.drag_target = None
            return

        # 画像の縦範囲内のみ反応（横は多少外れてもOKにするため、yだけで判定）
        if not (self.offset_y <= event.y <= self.offset_y + self.preview_h):
            self.drag_target = None
            return

        y = event.y
        rel_y = y - self.offset_y
        rel_y = max(0, min(self.preview_h, rel_y))
        pct = (rel_y / self.preview_h) * 100.0

        top_y = self.offset_y + int(self.preview_h * (self.top_pct / 100.0))
        bottom_y = self.offset_y + int(self.preview_h * (self.bottom_pct / 100.0))

        # 近い方を対象にする（従来の「15px以内」条件を撤廃）
        if abs(y - top_y) <= abs(y - bottom_y):
            self.drag_target = "top"
            self.top_pct = min(pct, self.bottom_pct - 1)
        else:
            self.drag_target = "bottom"
            self.bottom_pct = max(pct, self.top_pct + 1)

        self.load_preview()

    def on_drag(self, event):
        if not self.drag_target:
            return
        if self.preview_h <= 0:
            return

        rel_y = event.y - self.offset_y
        rel_y = max(0, min(self.preview_h, rel_y))
        pct = (rel_y / self.preview_h) * 100.0

        if self.drag_target == "top":
            self.top_pct = min(pct, self.bottom_pct - 1)
        elif self.drag_target == "bottom":
            self.bottom_pct = max(pct, self.top_pct + 1)

        self.load_preview()

    def on_release(self, event):
        self.drag_target = None

    def on_ok(self):
        if self.callback:
            self.callback(self.top_pct, self.bottom_pct)
        self.destroy()


# ==========================================
# メインアプリケーション
# ==========================================
class UnifiedYomitokuApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Yomitoku OCR & EPUB Workflow 統合ツール v5.9")
        self.root.geometry("980x820")

        self.device = "cuda" if (torch and torch.cuda.is_available()) else "cpu"

        self.tab_control = ttk.Notebook(root)
        self.tab_control.pack(expand=1, fill="both")

        self.tab_ocr = ttk.Frame(self.tab_control)
        self.tab_split = ttk.Frame(self.tab_control)
        self.tab_merge = ttk.Frame(self.tab_control)
        self.tab_epub = ttk.Frame(self.tab_control)
        self.tab_tools = ttk.Frame(self.tab_control)

        self.tab_control.add(self.tab_ocr, text="1. OCR処理")
        self.tab_control.add(self.tab_split, text="2. MD分割")
        self.tab_control.add(self.tab_merge, text="3. MD結合")
        self.tab_control.add(self.tab_epub, text="4. EPUB化")
        self.tab_control.add(self.tab_tools, text="5. その他")

        self.create_menu()
        self.init_tab_ocr()
        self.init_tab_split()
        self.init_tab_merge()
        self.init_tab_epub()
        self.init_tab_tools()
        self.init_log_area()

        self.log("起動しました。")

        self._split_preview_params = None
        self._last_split_preview_plan = None

    def safe_showinfo(self, title, message):
        if threading.current_thread() is threading.main_thread():
            messagebox.showinfo(title, message)
        else:
            self.root.after(0, lambda: messagebox.showinfo(title, message))

    def safe_showerror(self, title, message):
        if threading.current_thread() is threading.main_thread():
            messagebox.showerror(title, message)
        else:
            self.root.after(0, lambda: messagebox.showerror(title, message))

    def create_menu(self):
        menubar = tk.Menu(self.root)
        help_menu = tk.Menu(menubar, tearoff=0)
        help_menu.add_command(label="使い方", command=self.show_global_usage)
        menubar.add_cascade(label="ヘルプ", menu=help_menu)
        self.root.config(menu=menubar)

    # ★改善：使い方をより詳細に
    def show_global_usage(self):
        msg = (
            "【Yomitoku OCR & EPUB Workflow 統合ツール v5.9 使い方】\n\n"
            "------------------------------------------------------------\n"
            "1) OCR処理（Tab1）\n"
            "------------------------------------------------------------\n"
            "■ 目的\n"
            "  PDFをページ単位で画像化し、YomitokuでOCRしてMarkdownに出力します。\n\n"
            "■ 手順\n"
            "  (1) 入力PDF：『参照』でPDFを選択\n"
            "  (2) Popplerパス：popplerのbinが入ったフォルダを指定（pdf2imageに必要）\n"
            "  (3) 出力先フォルダ：結果（output.md / assets/）の保存先\n"
            "  (4) ページ範囲：開始～終了を指定（終了は自動補正されます）\n"
            "  (5) 上下カット(%)：上/下のトリミング範囲（0～100）\n"
            "      - 『ビジュアル範囲指定』で実ページ画像を見ながら調整できます\n"
            "  (6) 『OCR実行』で処理開始\n\n"
            "■ 出力\n"
            "  output.md（本文 + ページ画像リンク）\n"
            "  assets/（各ページ画像）\n\n"
            "■ 補足\n"
            "  ・CUDAメモリ不足などの場合、CPUへ切替して続行する場合があります。\n"
            "  ・暗号化（パスワード保護）PDFの疑いがある場合はエラー表示します。\n\n"
            "------------------------------------------------------------\n"
            "2) MD分割（Tab2）\n"
            "------------------------------------------------------------\n"
            "■ 目的\n"
            "  大きいMarkdownを校正/編集しやすいサイズに分割します。\n\n"
            "■ 主な項目\n"
            "  ・分割数：例）16\n"
            "  ・見出しレベル(1-6)：^#{1,N}\\s の見出しを「分割の境界」として扱います。\n"
            "    - 例）1なら『# 』のみを境界にするため、OCRで###が大量に混ざっても影響を受けにくいです。\n\n"
            "■ テスト（サイズ確認）\n"
            "  ・『テスト（サイズ確認）』はファイル書き出しをせず、\n"
            "    分割後の各ファイルの想定サイズ/行数/文字数/既存ファイル有無を一覧表示します。\n"
            "  ・一覧の行をダブルクリックすると、その分割予定テキストをプレビュー表示します（書き出しなし）。\n\n"
            "■ JSON書き出し\n"
            "  ・『JSON書き出し』で、テスト結果一覧（No/ファイル名/サイズ/既存など）をJSONに保存できます。\n"
            "    - 後で外部処理（ログ、レポート、検証）に使えます。\n\n"
            "■ 分割実行\n"
            "  ・『分割実行』で実際に {name_root}_01.md ... のように書き出します。\n"
            "  ・同名ファイルが既にある場合は上書き確認が出ます。\n\n"
            "------------------------------------------------------------\n"
            "3) MD結合（Tab3）\n"
            "------------------------------------------------------------\n"
            "■ 目的\n"
            "  クリップボードを監視してコピーした本文をスタックに積み、順序調整・編集して結合保存します。\n\n"
            "■ 手順\n"
            "  (1) 『監視開始』でクリップボード監視\n"
            "  (2) テキストをコピーすると自動でスタックに追加\n"
            "  (3) ダブルクリックでプレビュー\n"
            "  (4) 『選択編集』『↑↓』で整形\n"
            "  (5) 『結合して保存』で1つのMDとして保存\n\n"
            "------------------------------------------------------------\n"
            "4) EPUB化（Tab4）\n"
            "------------------------------------------------------------\n"
            "■ 目的\n"
            "  MarkdownをHTML化してEPUBを書き出します。\n\n"
            "■ 手順\n"
            "  (1) 入力MD/出力EPUB/タイトル/著者を指定\n"
            "  (2) 『EPUB作成』\n\n"
            "------------------------------------------------------------\n"
            "5) その他（Tab5）\n"
            "------------------------------------------------------------\n"
            "  ・Send to Kindle を開く\n"
            "  ・ログクリア\n"
        )
        messagebox.showinfo("使い方", msg)

    def init_log_area(self):
        frame = ttk.LabelFrame(self.root, text="ログ")
        frame.pack(side="bottom", fill="both", padx=10, pady=5)
        self.log_text = scrolledtext.ScrolledText(frame, height=10, wrap=tk.WORD)
        self.log_text.pack(fill="both", expand=True)

    def log(self, message):
        if not self.log_text.winfo_exists():
            return
        self.log_text.insert(tk.END, message + "\n")
        self.log_text.see(tk.END)
        lines = int(self.log_text.index("end-1c").split(".")[0])
        if lines > MAX_LOG_LINES:
            self.log_text.delete("1.0", "2.0")

    # ==========================================
    # Tab 1: OCR
    # ==========================================
    def init_tab_ocr(self):
        frm = ttk.Frame(self.tab_ocr)
        frm.pack(fill="both", expand=True, padx=10, pady=10)

        row1 = ttk.Frame(frm)
        row1.pack(fill="x", pady=5)
        ttk.Label(row1, text="入力PDF:").pack(side="left")
        self.pdf_path_var = tk.StringVar()
        ttk.Entry(row1, textvariable=self.pdf_path_var).pack(side="left", fill="x", expand=True, padx=5)
        ttk.Button(row1, text="参照", command=self.select_pdf).pack(side="left")

        row2 = ttk.Frame(frm)
        row2.pack(fill="x", pady=5)
        ttk.Label(row2, text="Popplerパス:").pack(side="left")
        default_poppler = DEFAULT_POPPLER_PATH if os.path.isdir(DEFAULT_POPPLER_PATH) else ""

        self.poppler_path_var = tk.StringVar(value=default_poppler)
        ttk.Entry(row2, textvariable=self.poppler_path_var).pack(side="left", fill="x", expand=True, padx=5)
        ttk.Button(row2, text="参照", command=self.select_poppler_dir).pack(side="left")

        row3 = ttk.Frame(frm)
        row3.pack(fill="x", pady=5)
        ttk.Label(row3, text="出力先フォルダ:").pack(side="left")
        self.output_dir_var = tk.StringVar()
        ttk.Entry(row3, textvariable=self.output_dir_var).pack(side="left", fill="x", expand=True, padx=5)
        ttk.Button(row3, text="参照", command=self.select_output_dir).pack(side="left")

        row4 = ttk.Frame(frm)
        row4.pack(fill="x", pady=5)
        ttk.Label(row4, text="ページ範囲:").pack(side="left")
        self.page_start_var = tk.IntVar(value=1)
        self.page_end_var = tk.IntVar(value=9999)
        ttk.Label(row4, text="開始").pack(side="left", padx=(10, 0))
        ttk.Entry(row4, textvariable=self.page_start_var, width=6).pack(side="left", padx=5)
        ttk.Label(row4, text="終了").pack(side="left")
        ttk.Entry(row4, textvariable=self.page_end_var, width=6).pack(side="left", padx=5)

        row5 = ttk.Frame(frm)
        row5.pack(fill="x", pady=5)
        ttk.Label(row5, text="上下カット(%):").pack(side="left")
        self.top_crop_var = tk.DoubleVar(value=0.0)
        self.bottom_crop_var = tk.DoubleVar(value=100.0)
        ttk.Label(row5, text="上").pack(side="left", padx=(10, 0))
        ttk.Entry(row5, textvariable=self.top_crop_var, width=6).pack(side="left", padx=5)
        ttk.Label(row5, text="下").pack(side="left")
        ttk.Entry(row5, textvariable=self.bottom_crop_var, width=6).pack(side="left", padx=5)
        ttk.Button(row5, text="ビジュアル範囲指定", command=self.open_visual_crop_dialog).pack(side="left", padx=10)

        row6 = ttk.Frame(frm)
        row6.pack(fill="x", pady=10)
        self.btn_run_ocr = ttk.Button(row6, text="OCR実行", command=self.run_ocr_thread)
        self.btn_run_ocr.pack(side="left", padx=5)

        self.progress_var = tk.DoubleVar(value=0.0)
        self.progress_bar = ttk.Progressbar(row6, variable=self.progress_var, maximum=100)
        self.progress_bar.pack(side="left", fill="x", expand=True, padx=10)
        self.lbl_status = ttk.Label(row6, text="待機中")
        self.lbl_status.pack(side="left")

        self.analyzer = None
        self.ocr_thread = None
        self.stop_event = threading.Event()

    def select_pdf(self):
        p = filedialog.askopenfilename(filetypes=[("PDF files", "*.pdf"), ("All files", "*.*")])
        if p:
            self.pdf_path_var.set(p)
            base = os.path.splitext(os.path.basename(p))[0]
            out_dir = os.path.join(os.path.dirname(p), base + "_out")
            self.output_dir_var.set(out_dir)

    def select_poppler_dir(self):
        d = filedialog.askdirectory()
        if d:
            self.poppler_path_var.set(d)

    def select_output_dir(self):
        d = filedialog.askdirectory()
        if d:
            self.output_dir_var.set(d)

    def open_visual_crop_dialog(self):
        pdf_path = self.pdf_path_var.get()
        poppler_path = self.poppler_path_var.get()
        if not pdf_path or not os.path.exists(pdf_path):
            self.safe_showerror("エラー", "PDFファイルを指定してください")
            return
        if not poppler_path or not os.path.exists(poppler_path):
            self.safe_showerror("エラー", "Popplerパスを指定してください")
            return

        def cb(top, bottom):
            self.top_crop_var.set(top)
            self.bottom_crop_var.set(bottom)

        VisualCropDialog(self.root, pdf_path, poppler_path, self.top_crop_var.get(), self.bottom_crop_var.get(), cb)

    def run_ocr_thread(self):
        if self.ocr_thread and self.ocr_thread.is_alive():
            self.safe_showerror("実行中", "OCR処理が実行中です。")
            return

        self.stop_event.clear()
        self.btn_run_ocr.config(state="disabled")
        self.lbl_status.config(text="準備中...")

        self.ocr_thread = threading.Thread(target=self.process_ocr)
        self.ocr_thread.daemon = True
        self.ocr_thread.start()

    def update_ocr_progress(self, pct, text):
        self.progress_var.set(pct)
        self.lbl_status.config(text=text)

    def process_ocr(self):
        pdf_path = self.pdf_path_var.get()
        poppler_path = self.poppler_path_var.get()
        out_dir = self.output_dir_var.get()
        start_page = self.page_start_var.get()
        end_page = self.page_end_var.get()
        top_pct = self.top_crop_var.get()
        bottom_pct = self.bottom_crop_var.get()

        if not pdf_path or not os.path.exists(pdf_path):
            self.root.after(0, lambda: self.safe_showerror("エラー", "PDFファイルを指定してください"))
            self.root.after(0, lambda: self.btn_run_ocr.config(state="normal"))
            return
        if not poppler_path or not os.path.exists(poppler_path):
            self.root.after(0, lambda: self.safe_showerror("エラー", "Popplerパスを指定してください"))
            self.root.after(0, lambda: self.btn_run_ocr.config(state="normal"))
            return
        if not out_dir:
            self.root.after(0, lambda: self.safe_showerror("エラー", "出力先フォルダを指定してください"))
            self.root.after(0, lambda: self.btn_run_ocr.config(state="normal"))
            return

        os.makedirs(out_dir, exist_ok=True)

        md_out_path = os.path.join(out_dir, "output.md")

        # 旧仕様との互換のため assets フォルダは残すが、クロップ画像（作業用）の残骸は必ず消す
        assets_dir = os.path.join(out_dir, "assets")
        os.makedirs(assets_dir, exist_ok=True)

        def _clear_dir_files(dir_path: str):
            if not dir_path or not os.path.isdir(dir_path):
                return
            # 下位も含めてファイルを削除（フォルダは残す）
            for root, dirs, files in os.walk(dir_path, topdown=False):
                for fn in files:
                    fp = os.path.join(root, fn)
                    try:
                        os.remove(fp)
                    except Exception:
                        pass
                for dn in dirs:
                    dp = os.path.join(root, dn)
                    try:
                        os.rmdir(dp)
                    except Exception:
                        pass

        _clear_dir_files(assets_dir)

        def _is_pdf_password_error(err: Exception) -> bool:
            msg = str(err).lower()
            return ("password" in msg) or ("incorrect password" in msg) or ("encrypted" in msg) or ("requires a password" in msg)

        def _is_cuda_oom(err: Exception) -> bool:
            msg = str(err).lower()
            return ("cuda" in msg and "out of memory" in msg) or ("cublas" in msg and "alloc" in msg)

        def _safe_after(fn):
            try:
                self.root.after(0, fn)
            except Exception:
                pass

        def _analyze_image(pil_img):
            """YomiTokuのDocumentAnalyzerはcv2画像(ndarray)が前提。PIL->ndarrayへ変換して推論する。
            戻り値: (results, cv2_img_bgr)
            """
            if cv2 is None or np is None:
                raise RuntimeError("opencv-python / numpy が読み込めません")
            rgb = np.array(pil_img)
            if rgb.ndim == 2:
                bgr = cv2.cvtColor(rgb, cv2.COLOR_GRAY2BGR)
            else:
                bgr = cv2.cvtColor(rgb, cv2.COLOR_RGB2BGR)

            out = self.analyzer(bgr)
            # バージョン差異： (results, ocr_vis, layout_vis) / (results, layout_vis) / results
            results = out[0] if isinstance(out, tuple) else out
            return results, bgr

        def _results_to_markdown(results, tmp_md_path, img_cv=None):
            """推論結果をMarkdown文字列へ（to_markdownがあれば利用）。

            NOTE:
            - YomiToku の to_markdown は図表(figure)の書き出し時に `img` が必須になることがあり、
              `img is required for saving figures` が出る場合は `img` を渡す必要があります。
            """
            if hasattr(results, "to_markdown"):
                last_err = None
                # 互換のため、複数シグネチャを順に試す
                call_variants = []
                if img_cv is not None:
                    call_variants += [
                        lambda: results.to_markdown(tmp_md_path, img=img_cv),
                        lambda: results.to_markdown(output_path=tmp_md_path, img=img_cv),
                        lambda: results.to_markdown(tmp_md_path, image=img_cv),
                        lambda: results.to_markdown(output_path=tmp_md_path, image=img_cv),
                    ]
                call_variants += [
                    lambda: results.to_markdown(tmp_md_path),
                    lambda: results.to_markdown(output_path=tmp_md_path),
                ]

                for fn in call_variants:
                    try:
                        fn()
                        last_err = None
                        break
                    except TypeError as e:
                        last_err = e
                        continue
                    except Exception as e:
                        last_err = e
                        # img必須系エラーなら、img付きの別シグネチャを試すため継続
                        continue

                if last_err is not None:
                    # ここで失敗しても、後段のフォールバック（テキスト抽出）は継続したいので例外は投げない
                    pass
                else:
                    try:
                        with open(tmp_md_path, "r", encoding="utf-8") as f:
                            return f.read()
                    finally:
                        try:
                            os.remove(tmp_md_path)
                        except Exception:
                            pass

            # フォールバック：paragraphs等からテキスト化
            chunks = []
            paras = getattr(results, "paragraphs", None)
            if paras:
                for para in paras:
                    if isinstance(para, dict):
                        t = para.get("contents") or para.get("text") or ""
                    else:
                        t = getattr(para, "contents", "") or getattr(para, "text", "")
                    t = (t or "").strip()
                    if t:
                        chunks.append(t)
            txt = "\n\n".join(chunks).strip()
            if not txt:
                txt = getattr(results, "text", "") or str(results)
            return txt

        # ---------------------------------------------------------
        # Tab1 post-process:
        #   (1) remove unwanted HTML tags without adding newlines (e.g. <BR> -> join)
        #   (2) keep ruby, but move it to separate lines with blank lines above/below (Method B)
        # NOTE: Other tabs / flows are intentionally untouched.
        # ---------------------------------------------------------
        _KANA_ONLY_RE = re.compile(r'^[ぁ-ゖァ-ヶー]+$')

        def _median(values):
            values = [v for v in values if v is not None]
            if not values:
                return None
            values = sorted(values)
            mid = len(values) // 2
            if len(values) % 2 == 1:
                return values[mid]
            return (values[mid - 1] + values[mid]) / 2.0

        def _get_words_from_results(res):
            # Try to extract OCR word list from various likely shapes:
            # - res.words
            # - res.ocr.words
            # - dict-like structures
            cand = None
            if isinstance(res, dict):
                cand = res.get("words") or (res.get("ocr") or {}).get("words")
            else:
                cand = getattr(res, "words", None)
                if cand is None:
                    ocr = getattr(res, "ocr", None) or getattr(res, "ocr_result", None) or getattr(res, "ocr_results", None)
                    cand = getattr(ocr, "words", None) if ocr is not None else None
                    if cand is None and isinstance(ocr, dict):
                        cand = ocr.get("words")
            if cand is None:
                return []
            return list(cand) if isinstance(cand, (list, tuple)) else []

        def _word_text_and_points(w):
            # Return (text, points) where points is list[[x,y],...]
            if isinstance(w, dict):
                txt = w.get("content") or w.get("text") or w.get("contents") or ""
                pts = w.get("points") or w.get("polygon")
                bbox = w.get("bbox") or w.get("box")
            else:
                txt = getattr(w, "content", None) or getattr(w, "text", None) or getattr(w, "contents", None) or ""
                pts = getattr(w, "points", None) or getattr(w, "polygon", None)
                bbox = getattr(w, "bbox", None) or getattr(w, "box", None)
            if pts is None and bbox is not None and isinstance(bbox, (list, tuple)) and len(bbox) == 4:
                x1, y1, x2, y2 = bbox
                pts = [[x1, y1], [x2, y1], [x2, y2], [x1, y2]]
            if not isinstance(pts, (list, tuple)):
                pts = None
            return (str(txt) if txt is not None else "", pts)

        def _box_w_h(points):
            if not points:
                return None, None
            try:
                xs = [p[0] for p in points]
                ys = [p[1] for p in points]
                w = max(xs) - min(xs)
                h = max(ys) - min(ys)
                return w, h
            except Exception:
                return None, None

        def _build_ruby_token_set(res):
            # Use bbox size to detect likely ruby tokens (smaller text vs body).
            words = _get_words_from_results(res)
            sizes = []
            items = []
            for w in words:
                txt, pts = _word_text_and_points(w)
                if not txt or not pts:
                    continue
                bw, bh = _box_w_h(pts)
                if bw is None or bh is None:
                    continue
                size = min(abs(bw), abs(bh))
                if size <= 0:
                    continue
                sizes.append(size)
                items.append((txt, size))

            med = _median(sizes)
            if med is None:
                return set()

            thr = med * 0.70  # ruby tends to be clearly smaller than body text
            ruby = set()
            for txt, size in items:
                t = (txt or "").strip()
                if not t:
                    continue
                # Ruby is usually kana; keep short tokens to reduce false positives
                if _KANA_ONLY_RE.match(t) and len(t) <= 4 and size <= thr:
                    ruby.add(t)
            return ruby

        def _cleanup_unwanted_html(md_text):
            """HTMLタグ由来の混入を消しつつ、rubyは『ルビ→本文』の順に独立行化する"""
            if not md_text:
                return md_text
            md = md_text

            # (A) <br> / <BR> は改行にせず除去して「行をくっつける」
            md = re.sub(r"<br\s*/?>", "", md, flags=re.IGNORECASE)

            # (B) HTML ruby tags が混ざる場合：ルビ（rt）→空行→本文（base）
            def _ruby_to_lines(m):
                inner = m.group(1) or ""
                rts = re.findall(r"<rt\b[^>]*>(.*?)</rt>", inner, flags=re.IGNORECASE | re.DOTALL)
                rt_txt = "".join([re.sub(r"<[^>]+>", "", x).strip() for x in rts]).strip()
                rt_txt = re.sub(r"\s+", "", rt_txt)  # 半角空白等があれば連結

                base = re.sub(r"<rt\b[^>]*>.*?</rt>", "", inner, flags=re.IGNORECASE | re.DOTALL)
                base = re.sub(r"<rp\b[^>]*>.*?</rp>", "", base, flags=re.IGNORECASE | re.DOTALL)
                base = re.sub(r"<[^>]+>", "", base).strip()

                if base and rt_txt:
                    return "\n\n" + rt_txt + "\n\n" + base
                return base or rt_txt or ""

            md = re.sub(r"<ruby\b[^>]*>(.*?)</ruby>", _ruby_to_lines, md, flags=re.IGNORECASE | re.DOTALL)

            # (C) ラッパータグは削除
            md = re.sub(r"</?(?:span|div|p|font|center|small|big)\b[^>]*>", "", md, flags=re.IGNORECASE)

            md = re.sub(r"[ \t]{2,}", " ", md)
            return md
        def _separate_ruby_lines(md_text, ruby_tokens):
            """'み き たけ' など、本文に混入したルビを独立行へ。
            ルビは本文行の『前』に出し、上下に空行を入れる。
            ルビに半角空白があっても改行せず連結して1行にする。
            """
            if not md_text:
                return md_text

            out_lines = []
            in_code = False

            cand_re = re.compile(r"(?:[ぁ-ゖァ-ヶー]{1,3}\s+){1,16}[ぁ-ゖァ-ヶー]{1,3}")

            for line in md_text.splitlines():
                s = line.rstrip("\n")
                stripped = s.strip()

                if stripped.startswith("```"):
                    in_code = not in_code
                    out_lines.append(s)
                    continue

                if in_code or stripped == "":
                    out_lines.append(s)
                    continue

                ruby_hits = []

                def _repl(m):
                    seq = m.group(0)
                    toks = seq.split()
                    if len(toks) < 2:
                        return seq

                    if ruby_tokens:
                        hit = sum(1 for t in toks if t in ruby_tokens)
                        if hit / float(len(toks)) < 0.70:
                            return seq

                    ruby_hits.append("".join(toks))
                    return ""

                new_line = cand_re.sub(_repl, s)
                new_line = re.sub(r"\s{2,}", " ", new_line).strip()

                # ★ルビがある場合は本文の前へ（上下空行つき）
                if ruby_hits:
                    out_lines.append("")
                    joined_ruby = re.sub(r"\s+", "", "".join(ruby_hits))  # 空白は連結
                    if joined_ruby:
                        out_lines.append(joined_ruby)
                    out_lines.append("")

                if new_line:
                    out_lines.append(new_line)

            joined = "\n".join(out_lines)
            joined = re.sub(r"\n{3,}", "\n\n", joined)
            return joined
        def _postprocess_page_md(page_md, res):
            md = _cleanup_unwanted_html(page_md)
            ruby_tokens = _build_ruby_token_set(res)
            md = _separate_ruby_lines(md, ruby_tokens)
            return md

        try:
            try:
                info = pdfinfo_from_path(pdf_path, poppler_path=poppler_path)
                total_pages = int(info.get("Pages", 0)) or 0
            except Exception as e:
                if _is_pdf_password_error(e):
                    _safe_after(lambda: self.safe_showerror("PDFエラー", "PDFがパスワード保護されている可能性があります。解除してから再試行してください。"))
                else:
                    _safe_after(lambda: self.safe_showerror("PDFエラー", f"PDF情報取得に失敗しました:\n{e}"))
                return

            if start_page < 1:
                start_page = 1
            if end_page > total_pages:
                end_page = total_pages
            if end_page < start_page:
                end_page = start_page

            self.log(f"【入力】PDF: {os.path.basename(pdf_path)}")
            self.log(f"【出力】フォルダ: {out_dir}")
            self.log(f"【出力】Markdown: {os.path.basename(md_out_path)}")
            self.log("【注意】クロップ画像（assetsフォルダ）は処理後に自動削除します")

            if self.analyzer is None:
                self.log(f"AIモデルロード中 ({self.device})...")
                try:
                    self.analyzer = DocumentAnalyzer(device=self.device)
                except RuntimeError as e:
                    if self.device == "cuda" and _is_cuda_oom(e):
                        self.log("CUDAメモリ不足のためCPUで再試行します。")
                        self.device = "cpu"
                        self.analyzer = DocumentAnalyzer(device=self.device)
                    else:
                        raise

            pages_to_process = end_page - start_page + 1
            md_lines = []
            md_lines.append(f"<!-- source_pdf: {os.path.basename(pdf_path)} -->\n\n")

            # ---- 表紙（無加工）を保存：クロップ前画像を表紙に使う（方針2） ----
            cover_path = os.path.join(out_dir, "cover.png")
            cover_saved = False

            # もし開始ページが1より後でも、表紙はPDFの1ページ目を無加工で保存しておく
            if start_page > 1:
                try:
                    imgs_cover = convert_from_path(
                        pdf_path,
                        first_page=1,
                        last_page=1,
                        poppler_path=poppler_path,
                    )
                    if imgs_cover:
                        imgs_cover[0].convert("RGB").save(cover_path, "PNG")
                        cover_saved = True
                        self.log(f"[INFO] 表紙画像を保存しました（無加工）: {cover_path}")
                    del imgs_cover
                    gc.collect()
                except Exception as e:
                    self.log(f"[WARN] 表紙画像の事前抽出に失敗しました: {e}")

            for idx, p in enumerate(range(start_page, end_page + 1), start=1):
                if self.stop_event.is_set():
                    self.log("停止要求を受け取りました。")
                    break

                pct = (idx / max(1, pages_to_process)) * 100.0
                _safe_after(lambda pct=pct, idx=idx, total=pages_to_process: self.update_ocr_progress(pct, f"OCR中... {idx}/{total}"))

                try:
                    imgs = convert_from_path(
                        pdf_path,
                        first_page=p,
                        last_page=p,
                        poppler_path=poppler_path,
                    )
                except Exception as e:
                    if _is_pdf_password_error(e):
                        self.log(f"[ERROR] パスワード保護PDFの疑い: {e}")
                        _safe_after(lambda: self.safe_showerror("PDFエラー", "PDFがパスワード保護されている可能性があります。解除してから再試行してください。"))
                        break
                    self.log(f"[ERROR] PDF->画像変換失敗 (page {p}): {e}")
                    continue

                if not imgs:
                    self.log(f"[WARN] 画像が生成されませんでした (page {p})")
                    continue

                img = imgs[0]

                # ---- 表紙用：クロップ前の画像を保持して保存（方針2） ----
                img_raw = img.copy()
                if p == 1 and not cover_saved:
                    try:
                        img_raw.convert("RGB").save(cover_path, "PNG")
                        cover_saved = True
                        self.log(f"[INFO] 表紙画像を保存しました（無加工）: {cover_path}")
                    except Exception as e:
                        self.log(f"[WARN] 表紙画像の保存に失敗しました: {e}")

                # ---- クロップ（上下カット） ----
                try:
                    w, h = img.size
                    top_px = int(h * (top_pct / 100.0))
                    bottom_px = int(h * (bottom_pct / 100.0))
                    if bottom_px <= top_px:
                        bottom_px = h
                        top_px = 0
                    img = img.crop((0, top_px, w, bottom_px))
                except Exception as e:
                    self.log(f"[WARN] crop失敗 (page {p}): {e}")

                # ---- OCR（YomiToku） ----
                try:
                    results, img_cv = _analyze_image(img)
                    tmp_md = os.path.join(out_dir, f"__tmp_page_{p:04}.md")
                    page_md = _results_to_markdown(results, tmp_md, img_cv=img_cv)
                    page_md = _postprocess_page_md(page_md, results)
                except RuntimeError as e:
                    if self.device == "cuda" and _is_cuda_oom(e):
                        self.log("CUDAメモリ不足。CPUへ切替して続行します。")
                        self.device = "cpu"
                        self.analyzer = DocumentAnalyzer(device=self.device)
                        results, img_cv = _analyze_image(img)
                        tmp_md = os.path.join(out_dir, f"__tmp_page_{p:04}.md")
                        page_md = _results_to_markdown(results, tmp_md, img_cv=img_cv)
                        page_md = _postprocess_page_md(page_md, results)
                    else:
                        self.log(f"[ERROR] OCR失敗 (page {p}): {e}")
                        continue
                except Exception as e:
                    self.log(f"[ERROR] OCR失敗 (page {p}): {e}")
                    continue

                # ---- まとめ出力 ----
                md_lines.append("\n\n")  # ページ区切り（ページ番号は出力しない）
                md_lines.append(page_md.rstrip() + "\n")

                del imgs
                del img
                try:
                    del img_raw
                except Exception:
                    pass
                gc.collect()

            try:
                with open(md_out_path, "w", encoding="utf-8") as f:
                    f.writelines(md_lines)
                self.log(f"Markdownを書き出しました: {md_out_path}")
                _safe_after(lambda: self.safe_showinfo("完了", f"OCRが完了しました。\n{md_out_path}"))
            except Exception as e:
                self.log(f"[ERROR] Markdown書き出し失敗: {e}")
                _safe_after(lambda: self.safe_showerror("エラー", f"Markdownの書き出しに失敗しました:\n{e}"))

        except Exception as e:
            self.log("【致命的エラー】\n" + traceback.format_exc())
            _safe_after(lambda: self.safe_showerror("致命的エラー", f"OCR処理中にエラーが発生しました:\n{e}"))
        finally:
            # クロップ画像（assetsフォルダ）を必ず消す
            try:
                _clear_dir_files(assets_dir)
            except Exception:
                pass
            _safe_after(lambda: self.btn_run_ocr.config(state="normal"))
            _safe_after(lambda: self.update_ocr_progress(0.0, "待機中"))

    # ==========================================
    # Tab 2: MD分割（テスト + ダブルクリックプレビュー + JSON書き出し）
    # ==========================================
    def init_tab_split(self):
        frm = ttk.Frame(self.tab_split)
        frm.pack(fill="both", expand=True, padx=10, pady=10)

        row1 = ttk.Frame(frm)
        row1.pack(fill="x", pady=5)
        ttk.Label(row1, text="入力MD:").pack(side="left")
        self.split_input_md_var = tk.StringVar()
        ttk.Entry(row1, textvariable=self.split_input_md_var).pack(side="left", fill="x", expand=True, padx=5)
        ttk.Button(row1, text="参照", command=self.select_split_input_md).pack(side="left")

        row2 = ttk.Frame(frm)
        row2.pack(fill="x", pady=5)
        ttk.Label(row2, text="分割数:").pack(side="left")
        self.split_count_var = tk.IntVar(value=16)
        ttk.Spinbox(row2, from_=2, to=100, textvariable=self.split_count_var, width=6).pack(side="left", padx=5)

        ttk.Label(row2, text="見出しレベル(1-6):").pack(side="left", padx=(20, 0))
        self.split_header_level_var = tk.IntVar(value=1)
        ttk.Spinbox(row2, from_=1, to=6, textvariable=self.split_header_level_var, width=6).pack(side="left", padx=5)

        row3 = ttk.Frame(frm)
        row3.pack(fill="x", pady=5)
        ttk.Label(row3, text="出力先(任意):").pack(side="left")
        self.split_output_dir_var = tk.StringVar()
        ttk.Entry(row3, textvariable=self.split_output_dir_var).pack(side="left", fill="x", expand=True, padx=5)
        ttk.Button(row3, text="参照", command=self.select_split_output_dir).pack(side="left")

        row4 = ttk.Frame(frm)
        row4.pack(fill="x", pady=10)
        ttk.Button(row4, text="テスト（サイズ確認）", command=self.run_split_preview).pack(side="left", padx=5)
        ttk.Button(row4, text="分割実行", command=self.run_split).pack(side="left", padx=5)
        ttk.Button(row4, text="JSON書き出し", command=self.export_split_preview_json).pack(side="left", padx=5)

        self.lbl_split_status = ttk.Label(row4, text="")
        self.lbl_split_status.pack(side="left", padx=10)

        preview_frame = ttk.LabelFrame(frm, text="分割テスト結果（書き出しなし） ※ダブルクリックで内容プレビュー")
        preview_frame.pack(fill="both", expand=True, pady=(5, 0))

        columns = ("no", "name", "kb", "bytes", "lines", "chars", "blocks", "exists")
        self.split_preview_tree = ttk.Treeview(preview_frame, columns=columns, show="headings", height=10)

        self.split_preview_tree.heading("no", text="No")
        self.split_preview_tree.heading("name", text="出力ファイル名")
        self.split_preview_tree.heading("kb", text="KB")
        self.split_preview_tree.heading("bytes", text="Bytes")
        self.split_preview_tree.heading("lines", text="行数")
        self.split_preview_tree.heading("chars", text="文字数")
        self.split_preview_tree.heading("blocks", text="ブロック数")
        self.split_preview_tree.heading("exists", text="既存")

        self.split_preview_tree.column("no", width=50, anchor="e")
        self.split_preview_tree.column("name", width=260, anchor="w")
        self.split_preview_tree.column("kb", width=90, anchor="e")
        self.split_preview_tree.column("bytes", width=110, anchor="e")
        self.split_preview_tree.column("lines", width=80, anchor="e")
        self.split_preview_tree.column("chars", width=90, anchor="e")
        self.split_preview_tree.column("blocks", width=90, anchor="e")
        self.split_preview_tree.column("exists", width=60, anchor="center")

        self.split_preview_tree.bind("<Double-Button-1>", self.on_split_preview_double_click)

        yscroll = ttk.Scrollbar(preview_frame, orient="vertical", command=self.split_preview_tree.yview)
        self.split_preview_tree.configure(yscrollcommand=yscroll.set)
        yscroll.pack(side="right", fill="y")
        self.split_preview_tree.pack(side="left", fill="both", expand=True)

        self.lbl_split_preview_summary = ttk.Label(frm, text="（テスト結果はここに集計表示）", foreground="gray")
        self.lbl_split_preview_summary.pack(anchor="w", pady=(6, 0))

    def select_split_input_md(self):
        p = filedialog.askopenfilename(filetypes=[("Markdown", "*.md"), ("All files", "*.*")])
        if p:
            self.split_input_md_var.set(p)

    def select_split_output_dir(self):
        d = filedialog.askdirectory()
        if d:
            self.split_output_dir_var.set(d)

    def _clear_split_preview(self):
        try:
            for item in self.split_preview_tree.get_children():
                self.split_preview_tree.delete(item)
        except Exception:
            pass
        try:
            self.lbl_split_preview_summary.config(text="（テスト結果はここに集計表示）", foreground="gray")
        except Exception:
            pass

    def run_split_preview(self):
        src_path = self.split_input_md_var.get()
        if not src_path or not os.path.exists(src_path):
            self.safe_showerror("エラー", "入力MDファイルを指定してください")
            return

        split_num = self.split_count_var.get()
        header_level = self.split_header_level_var.get()
        out_dir = self.split_output_dir_var.get().strip()
        if out_dir == "":
            out_dir = None

        try:
            plan = self.build_split_plan(src_path, split_num, header_level, out_dir, include_text=False)
            self.update_split_preview(plan)
            self.lbl_split_status.config(text="テスト完了")

            self._split_preview_params = (src_path, split_num, header_level, out_dir)
            self._last_split_preview_plan = plan

        except Exception as e:
            self.safe_showerror("エラー", str(e))
            self.lbl_split_status.config(text="テスト失敗")
            self._split_preview_params = None
            self._last_split_preview_plan = None

    def update_split_preview(self, plan):
        self._clear_split_preview()
        if not plan:
            self.lbl_split_preview_summary.config(text="分割案が生成されませんでした。", foreground="gray")
            return

        sizes = []
        exists_count = 0
        total_bytes = 0

        for it in plan:
            b = int(it.get("bytes", 0))
            kb = b / 1024.0
            total_bytes += b
            sizes.append(b)
            exists = "YES" if it.get("exists", False) else ""
            if exists:
                exists_count += 1

            self.split_preview_tree.insert(
                "",
                "end",
                values=(
                    it.get("index", ""),
                    it.get("name", ""),
                    f"{kb:,.1f}",
                    f"{b:,}",
                    f"{int(it.get('lines', 0)):,}",
                    f"{int(it.get('chars', 0)):,}",
                    f"{int(it.get('blocks', 0)):,}",
                    exists,
                ),
            )

        min_b = min(sizes) if sizes else 0
        max_b = max(sizes) if sizes else 0
        avg_b = (sum(sizes) / len(sizes)) if sizes else 0

        msg = (
            f"件数={len(plan)} / 合計={total_bytes/1024.0:,.1f} KB / "
            f"最小={min_b/1024.0:,.1f} KB / 平均={avg_b/1024.0:,.1f} KB / 最大={max_b/1024.0:,.1f} KB"
        )
        if exists_count > 0:
            msg += f" / 既存ファイル={exists_count}件（上書き確認あり）"

        self.lbl_split_preview_summary.config(text=msg, foreground="black")

    def export_split_preview_json(self):
        if not self._last_split_preview_plan:
            self.safe_showerror("エラー", "分割テスト結果がありません。先に「テスト（サイズ確認）」を実行してください。")
            return

        src_path = self.split_input_md_var.get()
        base_name = "split_preview.json"
        initialdir = None
        try:
            if src_path and os.path.exists(src_path):
                name_root = os.path.splitext(os.path.basename(src_path))[0]
                base_name = f"{name_root}_split_preview.json"
                initialdir = os.path.dirname(src_path)
        except Exception:
            pass

        save_path = filedialog.asksaveasfilename(
            defaultextension=".json",
            filetypes=[("JSON", "*.json"), ("All files", "*.*")],
            initialfile=base_name,
            initialdir=initialdir,
        )
        if not save_path:
            return

        params = self._split_preview_params
        if params:
            src_path_p, split_num_p, header_level_p, out_dir_p = params
        else:
            src_path_p = src_path
            split_num_p = self.split_count_var.get()
            header_level_p = self.split_header_level_var.get()
            out_dir_p = (self.split_output_dir_var.get().strip() or None)

        payload = {
            "type": "md_split_preview",
            "generated_at": time.strftime("%Y-%m-%d %H:%M:%S"),
            "source_md": src_path_p,
            "split_count": int(split_num_p),
            "header_level": int(header_level_p),
            "output_dir": out_dir_p,
            "rows": [],
        }

        for it in self._last_split_preview_plan:
            payload["rows"].append(
                {
                    "index": int(it.get("index", 0)),
                    "name": it.get("name"),
                    "path": it.get("path"),
                    "bytes": int(it.get("bytes", 0)),
                    "lines": int(it.get("lines", 0)),
                    "chars": int(it.get("chars", 0)),
                    "blocks": int(it.get("blocks", 0)),
                    "exists": bool(it.get("exists", False)),
                }
            )

        try:
            with open(save_path, "w", encoding="utf-8") as f:
                json.dump(payload, f, ensure_ascii=False, indent=2)
            self.safe_showinfo("完了", f"JSONを書き出しました:\n{save_path}")
            self.log(f"JSON書き出し: {save_path}")
        except Exception as e:
            self.safe_showerror("エラー", f"JSON書き出しに失敗しました:\n{e}")

    def on_split_preview_double_click(self, event=None):
        item_id = self.split_preview_tree.focus()
        if not item_id:
            return

        values = self.split_preview_tree.item(item_id, "values")
        if not values:
            return

        try:
            idx = int(values[0])
        except Exception:
            return

        if not self._split_preview_params:
            src_path = self.split_input_md_var.get()
            if not src_path or not os.path.exists(src_path):
                self.safe_showerror("エラー", "入力MDファイルを指定してください")
                return
            split_num = self.split_count_var.get()
            header_level = self.split_header_level_var.get()
            out_dir = self.split_output_dir_var.get().strip()
            if out_dir == "":
                out_dir = None
            params = (src_path, split_num, header_level, out_dir)
        else:
            params = self._split_preview_params

        src_path, split_num, header_level, out_dir = params

        try:
            plan_text = self.build_split_plan(src_path, split_num, header_level, out_dir, include_text=True)
            target = None
            for it in plan_text:
                if int(it.get("index", -1)) == idx:
                    target = it
                    break
            if not target:
                self.safe_showerror("エラー", "対象の分割データが見つかりませんでした")
                return

            b = int(target.get("bytes", 0))
            kb = b / 1024.0
            title = f"分割プレビュー No.{idx:02}  ({kb:,.1f} KB / {b:,} bytes)"
            self._open_merge_text_viewer(title, target.get("text", ""))

        except Exception as e:
            self.safe_showerror("エラー", f"プレビュー生成に失敗しました:\n{e}")

    def run_split(self):
        src_path = self.split_input_md_var.get()
        if not src_path or not os.path.exists(src_path):
            self.safe_showerror("エラー", "入力MDファイルを指定してください")
            return

        split_num = self.split_count_var.get()
        header_level = self.split_header_level_var.get()
        out_dir = self.split_output_dir_var.get().strip()
        if out_dir == "":
            out_dir = None

        try:
            self.split_markdown_file(src_path, split_num, header_level, out_dir)
            self.lbl_split_status.config(text="分割完了")
        except Exception as e:
            self.safe_showerror("エラー", str(e))
            self.lbl_split_status.config(text="分割失敗")

    def build_split_plan(self, src_path, split_num, header_level=6, output_dir=None, include_text=False):
        with open(src_path, "r", encoding="utf-8") as f:
            lines = f.readlines()

        try:
            header_level = int(header_level)
        except Exception:
            header_level = 6
        header_level = max(1, min(6, header_level))

        header_pattern = re.compile(r"^#{1," + str(header_level) + r"}\s")

        blocks = []
        current_block = []
        for line in lines:
            if header_pattern.match(line) and current_block:
                blocks.append("".join(current_block))
                current_block = []
            current_block.append(line)
        if current_block:
            blocks.append("".join(current_block))

        if len(blocks) < split_num:
            blocks = ["".join(lines)]

        blocks_per_file = math.ceil(len(blocks) / split_num)

        src_dir, base_name = os.path.split(src_path)
        name_root, ext = os.path.splitext(base_name)

        dir_name = output_dir.strip() if isinstance(output_dir, str) else output_dir
        if dir_name:
            dir_name = os.path.abspath(dir_name)
            if not os.path.isdir(dir_name):
                raise FileNotFoundError(f"出力先フォルダが存在しません: {dir_name}")
        else:
            dir_name = src_dir

        plan = []
        for i in range(split_num):
            chunk_blocks = blocks[i * blocks_per_file : (i + 1) * blocks_per_file]
            if not chunk_blocks:
                break

            text = "".join(chunk_blocks)
            out_name = f"{name_root}_{i+1:02}{ext}"
            out_path = os.path.join(dir_name, out_name)

            b = len(text.encode("utf-8"))
            item = {
                "index": i + 1,
                "name": out_name,
                "path": out_path,
                "bytes": b,
                "lines": text.count("\n") + 1 if text else 0,
                "chars": len(text),
                "blocks": len(chunk_blocks),
                "exists": os.path.exists(out_path),
            }
            if include_text:
                item["text"] = text
            plan.append(item)

        return plan

    def split_markdown_file(self, src_path, split_num, header_level=6, output_dir=None):
        plan = self.build_split_plan(src_path, split_num, header_level, output_dir, include_text=True)

        if not plan:
            raise RuntimeError("分割案が生成できませんでした。入力内容を確認してください。")

        existing = [it["path"] for it in plan if it.get("exists", False)]
        if existing:
            preview = "\n".join(os.path.basename(p) for p in existing[:10])
            if len(existing) > 10:
                preview += "\n..."
            msg = "以下のファイルが既に存在します。\n上書きして分割しますか？\n\n" + preview
            res = messagebox.askyesno("上書き確認", msg)
            if not res:
                self.log("分割を中止しました（上書き確認でキャンセル）。")
                return

        for it in plan:
            out_path = it["path"]
            text = it.get("text", "")
            with open(out_path, "w", encoding="utf-8") as f:
                f.write(text)
            self.log(f"  -> 作成: {os.path.basename(out_path)}")

        try:
            preview_plan = self.build_split_plan(src_path, split_num, header_level, output_dir, include_text=False)
            self.update_split_preview(preview_plan)
            self._last_split_preview_plan = preview_plan
            self._split_preview_params = (src_path, split_num, header_level, output_dir)
        except Exception:
            pass

        self.safe_showinfo("完了", f"分割しました。\n出力先: {os.path.dirname(plan[0]['path'])}")

    # ==========================================
    # Tab 3: MD結合（v14風）
    # ==========================================
    def init_tab_merge(self):
        frame = self.tab_merge

        self.merge_stack = []
        self.is_monitoring = False
        self.last_clipboard_hash = None

        container = tk.Frame(frame, pady=10)
        container.pack(fill="both", expand=True, padx=20)

        self.label_merge_status = tk.Label(container, text="待機中", fg="gray")
        self.label_merge_status.pack(pady=5)

        self.btn_monitor = tk.Button(container, text="監視開始", command=self.toggle_monitoring, bg="#fff9c4", height=2)
        self.btn_monitor.pack(fill="x", pady=5)

        tk.Label(container, text="スタック (プレビュー):").pack(anchor="w", pady=(10, 0))

        list_frame = tk.Frame(container)
        list_frame.pack(fill="both", expand=True, pady=5)

        self.listbox = tk.Listbox(list_frame, height=8, selectmode="extended")
        scrollbar = tk.Scrollbar(list_frame, orient="vertical", command=self.listbox.yview)
        self.listbox.config(yscrollcommand=scrollbar.set)
        scrollbar.pack(side="right", fill="y")
        self.listbox.pack(side="left", fill="both", expand=True)

        self.listbox.bind("<Double-Button-1>", self.preview_selected_item)
        tk.Label(container, text="※ ダブルクリックでプレビュー / 『選択編集』で内容を修正できます", fg="gray").pack(anchor="w", pady=(0, 5))

        btn_frame = tk.Frame(container)
        btn_frame.pack(fill="x", pady=5)

        tk.Button(btn_frame, text="手動追加", command=self.add_manual_item).pack(side="left")
        tk.Button(btn_frame, text="選択編集", command=self.edit_selected_item).pack(side="left", padx=10)
        tk.Button(btn_frame, text="選択削除", command=self.delete_selected_item).pack(side="left")
        tk.Button(btn_frame, text="↑ 上へ", command=self.move_item_up, width=6).pack(side="left", padx=10)
        tk.Button(btn_frame, text="↓ 下へ", command=self.move_item_down, width=6).pack(side="left")

        self.label_stack_count = tk.Label(btn_frame, text="件数: 0")
        self.label_stack_count.pack(side="right")

        tk.Button(container, text="結合して保存", command=self.save_merged_file, bg="#c8e6c9", height=2).pack(fill="x", pady=15)
        tk.Button(container, text="リセット", command=self.reset_stack).pack(anchor="e")

    def toggle_monitoring(self):
        if pyperclip is None:
            self.safe_showerror("エラー", "pyperclipがインストールされていません")
            return

        if not self.is_monitoring:
            try:
                c = pyperclip.paste()
                self.last_clipboard_hash = hashlib.sha256(c.encode("utf-8", "ignore")).hexdigest() if c else None
            except Exception:
                pass
            self.is_monitoring = True
            self.btn_monitor.config(text="監視停止", bg="#ffab91")
            self.label_merge_status.config(text="監視中...", fg="blue")
            threading.Thread(target=self.monitor_loop, daemon=True).start()
        else:
            self.is_monitoring = False
            self.btn_monitor.config(text="監視開始", bg="#fff9c4")
            self.label_merge_status.config(text="一時停止中", fg="gray")

    def monitor_loop(self):
        while self.is_monitoring and self.root.winfo_exists():
            try:
                content = pyperclip.paste()
                if content:
                    chash = hashlib.sha256(content.encode("utf-8", "ignore")).hexdigest()
                    if chash != self.last_clipboard_hash and content.strip():
                        self.merge_stack.append(content.rstrip("\n"))
                        self.last_clipboard_hash = chash
                        self.root.after(0, self.update_list_display)
            except Exception:
                pass
            time.sleep(1.0)

    def update_list_display(self, select_idx=None):
        self.listbox.delete(0, tk.END)
        for i, t in enumerate(self.merge_stack):
            one = t.replace("\r\n", "\n").replace("\n", " ")
            self.listbox.insert(tk.END, f"{i+1}: {one[:80]}{'...' if len(one) > 80 else ''}")

        self.label_stack_count.config(text=f"件数: {len(self.merge_stack)}")
        if select_idx is not None:
            try:
                self.listbox.selection_clear(0, tk.END)
                self.listbox.selection_set(select_idx)
                self.listbox.see(select_idx)
            except Exception:
                pass
        else:
            self.listbox.yview(tk.END)

    def move_item_up(self):
        idxs = self.listbox.curselection()
        if not idxs:
            return
        idx = idxs[0]
        if idx > 0:
            self.merge_stack[idx], self.merge_stack[idx - 1] = self.merge_stack[idx - 1], self.merge_stack[idx]
            self.update_list_display(select_idx=idx - 1)

    def move_item_down(self):
        idxs = self.listbox.curselection()
        if not idxs:
            return
        idx = idxs[0]
        if idx < len(self.merge_stack) - 1:
            self.merge_stack[idx], self.merge_stack[idx + 1] = self.merge_stack[idx + 1], self.merge_stack[idx]
            self.update_list_display(select_idx=idx + 1)

    def delete_selected_item(self):
        idxs = list(self.listbox.curselection())
        if not idxs:
            return
        for idx in reversed(idxs):
            try:
                del self.merge_stack[idx]
            except Exception:
                pass
        self.update_list_display()

    def preview_selected_item(self, event=None):
        idxs = self.listbox.curselection()
        if not idxs:
            return
        idx = idxs[0]
        self._open_merge_text_viewer(f"プレビュー（{idx+1}）", self.merge_stack[idx])

    def edit_selected_item(self):
        idxs = self.listbox.curselection()
        if not idxs:
            return
        if len(idxs) != 1:
            self.safe_showerror("エラー", "編集は1件だけ選択してください")
            return
        idx = idxs[0]
        original = self.merge_stack[idx]

        def _on_save(new_text: str):
            new_text = (new_text or "").rstrip("\n")
            if not new_text.strip():
                self.safe_showerror("エラー", "空の内容にはできません")
                return
            self.merge_stack[idx] = new_text
            self.update_list_display(select_idx=idx)

        self._open_merge_text_editor(f"選択編集（{idx+1}）", original, _on_save)

    def add_manual_item(self):
        initial = ""
        try:
            if pyperclip is not None:
                initial = pyperclip.paste() or ""
        except Exception:
            initial = ""

        def _on_save(new_text: str):
            new_text = (new_text or "").rstrip("\n")
            if not new_text.strip():
                return
            self.merge_stack.append(new_text)
            self.update_list_display(select_idx=len(self.merge_stack) - 1)

        self._open_merge_text_editor("手動追加", initial, _on_save)

    def _open_merge_text_viewer(self, title: str, text_value: str):
        win = tk.Toplevel(self.root)
        win.title(title)
        win.geometry("760x520")
        win.transient(self.root)
        win.grab_set()

        txt = scrolledtext.ScrolledText(win, wrap=tk.WORD)
        txt.pack(fill="both", expand=True, padx=10, pady=10)
        txt.insert("1.0", text_value)
        txt.config(state="disabled")

        btns = tk.Frame(win)
        btns.pack(fill="x", padx=10, pady=(0, 10))
        tk.Button(btns, text="閉じる", command=win.destroy, width=10).pack(side="right")

    def _open_merge_text_editor(self, title: str, initial_text: str, on_save):
        win = tk.Toplevel(self.root)
        win.title(title)
        win.geometry("760x560")
        win.transient(self.root)
        win.grab_set()

        txt = scrolledtext.ScrolledText(win, wrap=tk.WORD)
        txt.pack(fill="both", expand=True, padx=10, pady=10)
        if initial_text:
            txt.insert("1.0", initial_text)

        btns = tk.Frame(win)
        btns.pack(fill="x", padx=10, pady=(0, 10))

        def _apply():
            value = txt.get("1.0", "end-1c")
            try:
                on_save(value)
            except Exception as e:
                self.safe_showerror("エラー", str(e))
                return
            win.destroy()

        tk.Button(btns, text="適用", command=_apply, width=10).pack(side="right")
        tk.Button(btns, text="キャンセル", command=win.destroy, width=10).pack(side="right", padx=8)

    def save_merged_file(self):
        if not self.merge_stack:
            self.safe_showerror("エラー", "スタックが空です")
            return
        path = filedialog.asksaveasfilename(defaultextension=".md", filetypes=[("MD", "*.md"), ("All", "*.*")])
        if path:
            try:
                with open(path, "w", encoding="utf-8") as f:
                    f.write("\n\n\n".join(self.merge_stack).rstrip() + "\n")
                self.safe_showinfo("成功", f"保存しました。\n{path}")
            except Exception as e:
                self.safe_showerror("エラー", str(e))

    def reset_stack(self):
        if messagebox.askyesno("確認", "クリアしますか？"):
            self.merge_stack = []
            self.update_list_display()
            try:
                c = pyperclip.paste()
                if c:
                    self.last_clipboard_hash = hashlib.sha256(c.encode("utf-8", "ignore")).hexdigest()
            except Exception:
                pass

    # ==========================================
    # Tab 4: EPUB化
    # ==========================================
    def init_tab_epub(self):
        frm = ttk.Frame(self.tab_epub)
        frm.pack(fill="both", expand=True, padx=10, pady=10)

        row1 = ttk.Frame(frm)
        row1.pack(fill="x", pady=5)
        ttk.Label(row1, text="入力MD:").pack(side="left")
        self.epub_input_md_var = tk.StringVar()
        ttk.Entry(row1, textvariable=self.epub_input_md_var).pack(side="left", fill="x", expand=True, padx=5)
        ttk.Button(row1, text="参照", command=self.select_epub_input_md).pack(side="left")

        row2 = ttk.Frame(frm)
        row2.pack(fill="x", pady=5)
        ttk.Label(row2, text="出力EPUB:").pack(side="left")
        self.epub_output_path_var = tk.StringVar()
        ttk.Entry(row2, textvariable=self.epub_output_path_var).pack(side="left", fill="x", expand=True, padx=5)
        ttk.Button(row2, text="参照", command=self.select_epub_output_path).pack(side="left")

        row3 = ttk.Frame(frm)
        row3.pack(fill="x", pady=5)
        ttk.Label(row3, text="タイトル:").pack(side="left")
        self.epub_title_var = tk.StringVar(value="EPUB Title")
        ttk.Entry(row3, textvariable=self.epub_title_var).pack(side="left", fill="x", expand=True, padx=5)

        row4 = ttk.Frame(frm)
        row4.pack(fill="x", pady=5)
        ttk.Label(row4, text="著者:").pack(side="left")
        self.epub_author_var = tk.StringVar(value="Author")
        ttk.Entry(row4, textvariable=self.epub_author_var).pack(side="left", fill="x", expand=True, padx=5)

        row5 = ttk.Frame(frm)
        row5.pack(fill="x", pady=10)
        ttk.Button(row5, text="EPUB作成", command=self.run_epub).pack(side="left", padx=5)
        self.lbl_epub_status = ttk.Label(row5, text="")
        self.lbl_epub_status.pack(side="left", padx=10)

    def select_epub_input_md(self):
        p = filedialog.askopenfilename(filetypes=[("Markdown", "*.md"), ("All files", "*.*")])
        if p:
            self.epub_input_md_var.set(p)

    def select_epub_output_path(self):
        p = filedialog.asksaveasfilename(defaultextension=".epub", filetypes=[("EPUB", "*.epub"), ("All files", "*.*")])
        if p:
            self.epub_output_path_var.set(p)

    def run_epub(self):
        md_path = self.epub_input_md_var.get()
        out_epub_path = self.epub_output_path_var.get()
        title = self.epub_title_var.get()
        author = self.epub_author_var.get()

        if not md_path or not os.path.exists(md_path):
            self.safe_showerror("エラー", "入力MDを指定してください")
            return
        if not out_epub_path:
            self.safe_showerror("エラー", "出力EPUBを指定してください")
            return

        try:
            with open(md_path, "r", encoding="utf-8") as f:
                md_text = f.read()

            html = markdown_lib.markdown(md_text, extensions=["tables", "fenced_code"])

            book = epub.EpubBook()
            book.set_identifier(hashlib.md5((title + author).encode("utf-8")).hexdigest())
            book.set_title(title)
            book.add_author(author)
            book.set_language("ja")

            # ---- 表紙設定（Tab1で保存した無加工 cover.png があれば利用） ----
            try:
                cover_path = os.path.join(os.path.dirname(md_path), "cover.png")
                if os.path.exists(cover_path):
                    with open(cover_path, "rb") as f:
                        book.set_cover("cover.png", f.read())
                    self.log(f"表紙を設定しました: {cover_path}")
            except Exception as e:
                self.log(f"[WARN] 表紙設定に失敗しました: {e}")

            c1 = epub.EpubHtml(title=title, file_name="content.xhtml", lang="ja")
            c1.content = html

            book.add_item(c1)
            book.toc = (epub.Link("content.xhtml", title, "content"),)
            book.spine = ["nav", c1]
            book.add_item(epub.EpubNcx())
            book.add_item(epub.EpubNav())

            epub.write_epub(out_epub_path, book)

            self.lbl_epub_status.config(text="EPUB作成完了")
            self.safe_showinfo("完了", f"EPUBを作成しました:\n{out_epub_path}")
            self.log(f"EPUB作成: {out_epub_path}")

        except Exception as e:
            self.safe_showerror("エラー", str(e))
            self.lbl_epub_status.config(text="EPUB作成失敗")

    # ==========================================
    # Tab 5: その他
    # ==========================================
    def init_tab_tools(self):
        frm = ttk.Frame(self.tab_tools)
        frm.pack(fill="both", expand=True, padx=10, pady=10)

        ttk.Button(frm, text="Send to Kindle を開く", command=self.open_send_to_kindle).pack(anchor="w", pady=5)
        ttk.Button(frm, text="ログクリア", command=self.clear_log).pack(anchor="w", pady=5)

    def open_send_to_kindle(self):
        webbrowser.open("https://www.amazon.co.jp/sendtokindle")

    def clear_log(self):
        try:
            self.log_text.delete("1.0", tk.END)
        except Exception:
            pass


def main():
    root = tk.Tk()
    UnifiedYomitokuApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()
