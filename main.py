import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import json
import os
import sys
import subprocess
import tempfile
import threading
from pathlib import Path
from PIL import Image, ImageTk

# ── 設定檔路徑 ──────────────────────────────────────────────
APPDATA = os.environ.get("APPDATA", str(Path.home()))
CONFIG_DIR = Path(APPDATA) / "PrinterSelector"
CONFIG_FILE = CONFIG_DIR / "config.json"

def load_config():
    if CONFIG_FILE.exists():
        try:
            return json.loads(CONFIG_FILE.read_text(encoding="utf-8"))
        except Exception:
            pass
    return {"bw_printer": "", "color_printer": ""}

def save_config(cfg):
    CONFIG_DIR.mkdir(parents=True, exist_ok=True)
    CONFIG_FILE.write_text(json.dumps(cfg, ensure_ascii=False, indent=2), encoding="utf-8")

# ── 取得系統印表機清單 ────────────────────────────────────────
def get_printers():
    try:
        import win32print
        printers = [p[2] for p in win32print.EnumPrinters(
            win32print.PRINTER_ENUM_LOCAL | win32print.PRINTER_ENUM_CONNECTIONS
        )]
        return printers
    except Exception:
        return []

def get_default_printer():
    try:
        import win32print
        return win32print.GetDefaultPrinter()
    except Exception:
        return ""

def set_default_printer(name):
    try:
        import win32print
        win32print.SetDefaultPrinter(name)
        return True
    except Exception:
        return False

# ── 檔案轉 PDF（Office COM）────────────────────────────────────
def convert_office_to_pdf(filepath):
    """將 Word/Excel/PPT 透過 Office COM 轉成暫存 PDF，回傳暫存路徑"""
    ext = Path(filepath).suffix.lower()
    tmp = tempfile.NamedTemporaryFile(suffix=".pdf", delete=False)
    tmp.close()
    out_path = tmp.name

    try:
        if ext in (".doc", ".docx"):
            import comtypes.client
            word = comtypes.client.CreateObject("Word.Application")
            word.Visible = False
            doc = word.Documents.Open(str(Path(filepath).resolve()))
            doc.SaveAs(out_path, FileFormat=17)  # 17 = wdFormatPDF
            doc.Close(False)
            word.Quit()

        elif ext in (".xls", ".xlsx"):
            import comtypes.client
            excel = comtypes.client.CreateObject("Excel.Application")
            excel.Visible = False
            wb = excel.Workbooks.Open(str(Path(filepath).resolve()))
            wb.ExportAsFixedFormat(0, out_path)  # 0 = xlTypePDF
            wb.Close(False)
            excel.Quit()

        elif ext in (".ppt", ".pptx"):
            import comtypes.client
            ppt = comtypes.client.CreateObject("PowerPoint.Application")
            prs = ppt.Presentations.Open(str(Path(filepath).resolve()), WithWindow=False)
            prs.SaveAs(out_path, 32)  # 32 = ppSaveAsPDF
            prs.Close()
            ppt.Quit()

        else:
            return None

        return out_path
    except Exception as e:
        print(f"Office 轉換失敗: {e}")
        return None

# ── PDF → 預覽圖片 ────────────────────────────────────────────
def pdf_to_images(pdf_path, max_pages=10):
    from pdf2image import convert_from_path
    poppler = get_poppler_path()
    images = convert_from_path(pdf_path, dpi=120, first_page=1,
                               last_page=max_pages, poppler_path=poppler)
    return images

def get_poppler_path():
    """exe 打包後 poppler 在 _MEIPASS 下"""
    if getattr(sys, "frozen", False):
        return str(Path(sys._MEIPASS) / "poppler" / "bin")
    return None  # 開發時需 poppler 在 PATH

# ── 列印 ──────────────────────────────────────────────────────
def print_file(filepath, printer_name):
    """暫切預設印表機 → 送印 → 恢復"""
    original = get_default_printer()
    try:
        if printer_name and printer_name != original:
            set_default_printer(printer_name)
        _do_print(filepath)
    finally:
        if printer_name and printer_name != original:
            set_default_printer(original)

def _do_print(filepath):
    ext = Path(filepath).suffix.lower()
    if ext == ".pdf":
        # 用 SumatraPDF 靜默列印（若有），否則 ShellExecute
        sumatra = _find_sumatra()
        if sumatra:
            subprocess.run([sumatra, "-print-to-default", "-silent", filepath])
        else:
            os.startfile(filepath, "print")
    else:
        os.startfile(filepath, "print")

def _find_sumatra():
    candidates = [
        r"C:\Program Files\SumatraPDF\SumatraPDF.exe",
        r"C:\Program Files (x86)\SumatraPDF\SumatraPDF.exe",
    ]
    for c in candidates:
        if Path(c).exists():
            return c
    return None

# ══════════════════════════════════════════════════════════════
# 設定視窗
# ══════════════════════════════════════════════════════════════
class SettingsWindow(tk.Toplevel):
    def __init__(self, parent, cfg, on_save):
        super().__init__(parent)
        self.title("印表機設定")
        self.resizable(False, False)
        self.grab_set()
        self.cfg = cfg.copy()
        self.on_save = on_save

        printers = get_printers()
        if not printers:
            printers = ["（找不到印表機）"]

        pad = dict(padx=16, pady=8)

        tk.Label(self, text="⬛ 黑白印表機", font=("", 11, "bold")).grid(
            row=0, column=0, sticky="w", **pad)
        self.bw_var = tk.StringVar(value=cfg.get("bw_printer", ""))
        bw_cb = ttk.Combobox(self, textvariable=self.bw_var,
                             values=printers, width=45, state="readonly")
        bw_cb.grid(row=0, column=1, **pad)

        tk.Label(self, text="🟦 彩色印表機", font=("", 11, "bold")).grid(
            row=1, column=0, sticky="w", **pad)
        self.color_var = tk.StringVar(value=cfg.get("color_printer", ""))
        color_cb = ttk.Combobox(self, textvariable=self.color_var,
                                values=printers, width=45, state="readonly")
        color_cb.grid(row=1, column=1, **pad)

        btn_frame = tk.Frame(self)
        btn_frame.grid(row=2, column=0, columnspan=2, pady=12)
        tk.Button(btn_frame, text="儲存", width=10, command=self._save).pack(side="left", padx=8)
        tk.Button(btn_frame, text="取消", width=10, command=self.destroy).pack(side="left", padx=8)

    def _save(self):
        self.cfg["bw_printer"] = self.bw_var.get()
        self.cfg["color_printer"] = self.color_var.get()
        save_config(self.cfg)
        self.on_save(self.cfg)
        self.destroy()

# ══════════════════════════════════════════════════════════════
# 主視窗
# ══════════════════════════════════════════════════════════════
class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("列印選擇器")
        self.geometry("860x680")
        self.minsize(700, 500)
        self.configure(bg="#f5f5f5")

        self.cfg = load_config()
        self.current_file = None       # 原始選取的檔案
        self.preview_pdf = None        # 實際用來預覽的 PDF（可能是暫存）
        self.preview_images = []       # PIL Image list
        self.tk_images = []            # 保留參考避免 GC
        self.current_page = 0
        self._tmp_files = []           # 打開時建立的暫存檔，關閉時清理

        self._build_ui()

    # ── UI 建構 ──────────────────────────────────────────────
    def _build_ui(self):
        # 頂部工具列
        toolbar = tk.Frame(self, bg="#e0e0e0", pady=6)
        toolbar.pack(fill="x")

        tk.Button(toolbar, text="📂 選擇檔案", font=("", 11),
                  command=self._open_file).pack(side="left", padx=12)
        tk.Button(toolbar, text="⚙ 設定印表機", font=("", 11),
                  command=self._open_settings).pack(side="left", padx=4)

        self.file_label = tk.Label(toolbar, text="尚未選擇檔案",
                                   bg="#e0e0e0", font=("", 10), fg="#555")
        self.file_label.pack(side="left", padx=16)

        # 預覽區
        preview_frame = tk.Frame(self, bg="#888")
        preview_frame.pack(fill="both", expand=True, padx=12, pady=(8, 4))

        self.canvas = tk.Canvas(preview_frame, bg="#666", cursor="hand2")
        self.canvas.pack(fill="both", expand=True)

        self.placeholder = tk.Label(self.canvas,
                                    text="請選擇檔案以顯示預覽",
                                    font=("", 14), fg="#ccc", bg="#666")
        self.canvas.create_window(430, 300, window=self.placeholder)

        # 翻頁列
        nav = tk.Frame(self, bg="#f5f5f5")
        nav.pack(fill="x", padx=12)

        self.prev_btn = tk.Button(nav, text="◀", command=self._prev_page,
                                  state="disabled", width=4)
        self.prev_btn.pack(side="left")
        self.page_label = tk.Label(nav, text="", bg="#f5f5f5", font=("", 10))
        self.page_label.pack(side="left", padx=8)
        self.next_btn = tk.Button(nav, text="▶", command=self._next_page,
                                  state="disabled", width=4)
        self.next_btn.pack(side="left")

        # 底部列印按鈕
        btn_frame = tk.Frame(self, bg="#f5f5f5", pady=10)
        btn_frame.pack(fill="x", padx=12)

        self.bw_btn = tk.Button(
            btn_frame,
            text="⬛  黑白列印",
            font=("", 16, "bold"),
            fg="white", bg="#222",
            activebackground="#444",
            width=18, height=2,
            command=self._print_bw,
            state="disabled"
        )
        self.bw_btn.pack(side="left", expand=True, fill="x", padx=(0, 8))

        self.color_btn = tk.Button(
            btn_frame,
            text="🎨  彩色列印",
            font=("", 16, "bold"),
            fg="white",
            bg="#1565C0",
            activebackground="#1976D2",
            width=18, height=2,
            command=self._print_color,
            state="disabled"
        )
        self.color_btn.pack(side="left", expand=True, fill="x", padx=(8, 0))

        self.status_label = tk.Label(self, text="", bg="#f5f5f5",
                                     font=("", 9), fg="#777")
        self.status_label.pack(pady=(0, 6))

        self.bind("<Configure>", self._on_resize)

    # ── 開檔 ─────────────────────────────────────────────────
    def _open_file(self):
        filetypes = [
            ("所有支援格式", "*.pdf *.png *.jpg *.jpeg *.bmp *.gif *.tiff "
                            "*.doc *.docx *.xls *.xlsx *.ppt *.pptx"),
            ("PDF", "*.pdf"),
            ("圖片", "*.png *.jpg *.jpeg *.bmp *.gif *.tiff"),
            ("Word", "*.doc *.docx"),
            ("Excel", "*.xls *.xlsx"),
            ("PowerPoint", "*.ppt *.pptx"),
        ]
        path = filedialog.askopenfilename(filetypes=filetypes)
        if not path:
            return

        self.current_file = path
        self.file_label.config(text=Path(path).name)
        self.bw_btn.config(state="disabled")
        self.color_btn.config(state="disabled")
        self.preview_images = []
        self.current_page = 0
        self._set_status("載入中…")

        threading.Thread(target=self._load_preview, daemon=True).start()

    def _load_preview(self):
        path = self.current_file
        ext = Path(path).suffix.lower()

        try:
            if ext == ".pdf":
                self.preview_pdf = path
                images = pdf_to_images(path)

            elif ext in (".png", ".jpg", ".jpeg", ".bmp", ".gif", ".tiff"):
                images = [Image.open(path)]
                self.preview_pdf = None

            elif ext in (".doc", ".docx", ".xls", ".xlsx", ".ppt", ".pptx"):
                self._set_status("透過 Office 轉換中，請稍候…")
                tmp_pdf = convert_office_to_pdf(path)
                if tmp_pdf:
                    self._tmp_files.append(tmp_pdf)
                    self.preview_pdf = tmp_pdf
                    images = pdf_to_images(tmp_pdf)
                else:
                    self.after(0, lambda: messagebox.showwarning(
                        "無法預覽",
                        "需要安裝 Microsoft Office 才能預覽此格式。\n仍可直接送印。"
                    ))
                    images = []
                    self.preview_pdf = None
            else:
                images = []
                self.preview_pdf = None

            self.preview_images = images
            self.after(0, self._show_current_page)
            self.after(0, lambda: self.bw_btn.config(state="normal"))
            self.after(0, lambda: self.color_btn.config(state="normal"))
            self.after(0, lambda: self._set_status(""))

        except Exception as e:
            self.after(0, lambda: self._set_status(f"預覽失敗：{e}"))
            self.after(0, lambda: self.bw_btn.config(state="normal"))
            self.after(0, lambda: self.color_btn.config(state="normal"))

    # ── 預覽顯示 ──────────────────────────────────────────────
    def _show_current_page(self):
        if not self.preview_images:
            self._update_nav()
            return

        img = self.preview_images[self.current_page]
        self._render_image(img)
        self._update_nav()

    def _render_image(self, pil_img):
        cw = self.canvas.winfo_width() or 800
        ch = self.canvas.winfo_height() or 560

        # 等比縮放
        iw, ih = pil_img.size
        scale = min(cw / iw, ch / ih, 1.0)
        new_w, new_h = int(iw * scale), int(ih * scale)
        resized = pil_img.resize((new_w, new_h), Image.LANCZOS)

        tk_img = ImageTk.PhotoImage(resized)
        self.tk_images = [tk_img]

        self.canvas.delete("all")
        self.canvas.create_image(cw // 2, ch // 2, image=tk_img, anchor="center")

    def _on_resize(self, event):
        if self.preview_images:
            self._show_current_page()

    def _update_nav(self):
        total = len(self.preview_images)
        if total <= 1:
            self.page_label.config(text="")
            self.prev_btn.config(state="disabled")
            self.next_btn.config(state="disabled")
        else:
            self.page_label.config(text=f"第 {self.current_page + 1} / {total} 頁")
            self.prev_btn.config(state="normal" if self.current_page > 0 else "disabled")
            self.next_btn.config(state="normal" if self.current_page < total - 1 else "disabled")

    def _prev_page(self):
        if self.current_page > 0:
            self.current_page -= 1
            self._show_current_page()

    def _next_page(self):
        if self.current_page < len(self.preview_images) - 1:
            self.current_page += 1
            self._show_current_page()

    # ── 列印 ─────────────────────────────────────────────────
    def _print_bw(self):
        self._do_print(self.cfg.get("bw_printer", ""), "黑白")

    def _print_color(self):
        self._do_print(self.cfg.get("color_printer", ""), "彩色")

    def _do_print(self, printer_name, label):
        if not self.current_file:
            messagebox.showwarning("尚未選擇檔案", "請先選擇要列印的檔案。")
            return
        if not printer_name:
            messagebox.showwarning("未設定印表機",
                                   f"請先至「設定印表機」設定{label}印表機。")
            return

        self._set_status(f"送印至 {printer_name}…")
        threading.Thread(
            target=self._print_thread,
            args=(self.current_file, printer_name, label),
            daemon=True
        ).start()

    def _print_thread(self, filepath, printer_name, label):
        try:
            print_file(filepath, printer_name)
            self.after(0, lambda: self._set_status(f"✅ 已送至{label}印表機：{printer_name}"))
        except Exception as e:
            self.after(0, lambda: self._set_status(f"❌ 列印失敗：{e}"))

    # ── 設定 ─────────────────────────────────────────────────
    def _open_settings(self):
        SettingsWindow(self, self.cfg, self._on_cfg_updated)

    def _on_cfg_updated(self, new_cfg):
        self.cfg = new_cfg

    def _set_status(self, msg):
        self.status_label.config(text=msg)

    # ── 關閉清理 ──────────────────────────────────────────────
    def destroy(self):
        for f in self._tmp_files:
            try:
                os.unlink(f)
            except Exception:
                pass
        super().destroy()


if __name__ == "__main__":
    app = App()
    app.mainloop()
