#!/usr/bin/env python3
"""MarkItDown 2.0 — cross-platform GUI / Web / Menubar wrapper.

Convert documents (PDF, Word, PowerPoint, Excel, HTML, images + OCR, ...)
into Markdown. Based on Microsoft's markitdown library.

Modes:
  (default)      desktop GUI (tkinter, drag & drop)
  --web          local web UI at http://127.0.0.1:8741
  --menubar      macOS menu-bar agent (requires rumps)
  --selftest f.. headless smoke test
"""

import json
import os
import platform
import queue
import re
import subprocess
import sys
import tempfile
import threading
import traceback
import webbrowser

__version__ = "2.0.0"
GITHUB_REPO = "jaber1985/markitdown-app"
CONFIG_PATH = os.path.join(os.path.expanduser("~"), ".markitdown_gui.json")

# ---------------- optional dependencies ----------------
try:
    from tkinterdnd2 import TkinterDnD, DND_FILES
    _TkBase = TkinterDnD.Tk
    HAS_DND = True
except Exception:  # noqa: BLE001
    import tkinter as tk
    _TkBase = tk.Tk
    DND_FILES = None
    HAS_DND = False

import tkinter as tk
from tkinter import ttk, filedialog, messagebox

try:
    from rapidocr_onnxruntime import RapidOCR
    HAS_OCR = True
except Exception:  # noqa: BLE001
    RapidOCR = None
    HAS_OCR = False

from markitdown import MarkItDown

IMAGE_EXTS = {".png", ".jpg", ".jpeg", ".bmp", ".gif", ".webp", ".tif", ".tiff"}
SUPPORTED_EXTS = {
    ".pdf", ".docx", ".pptx", ".xlsx", ".xls", ".html", ".htm", ".csv",
    ".json", ".xml", ".zip", ".epub", ".txt", ".md", ".msg", ".ipynb",
    ".wav", ".mp3", ".m4a", ".flac",
} | IMAGE_EXTS

# ---------------- i18n ----------------
STRINGS = {
    "ar": {
        "title": "MarkItDown — تحويل الملفات إلى Markdown",
        "subtitle": "حوّل مستنداتك إلى Markdown. اسحب الملفات إلى القائمة مباشرة أو أضفها بالأزرار.",
        "supported": "الصيغ المدعومة: PDF, Word, PowerPoint, Excel, HTML, CSV, JSON, XML, ZIP, EPub, صور (مع OCR), صوت، وروابط صفحات و YouTube",
        "files_frame": "الملفات المحددة (اسحب وأفلت هنا)",
        "add_files": "➕ إضافة ملفات…",
        "add_folder": "📁 إضافة مجلد…",
        "remove": "إزالة المحدد",
        "clear": "مسح الكل",
        "url_placeholder": "أو ألصق رابط صفحة / فيديو YouTube هنا…",
        "convert_url": "تحويل الرابط",
        "out_frame": "مكان الحفظ",
        "same_folder": "بجانب الملف الأصلي",
        "custom_folder": "مجلد محدد:",
        "browse": "اختيار…",
        "merge": "دمج كل النتائج في ملف Markdown واحد",
        "convert": "🔄 تحويل",
        "preview": "👁 معاينة النتيجة",
        "log_frame": "السجل",
        "need_file": "أضف ملفاً أو رابطاً واحداً على الأقل أولاً.",
        "need_out": "اختر مجلد الحفظ أولاً.",
        "converting": "تحويل",
        "init": "جارٍ تهيئة المحوّل…",
        "ok": "✔ تم",
        "fail": "✖ فشل",
        "done_ok": "تم تحويل {} ملف بنجاح ✅",
        "done_mix": "نجح {} وفشل {}. راجع السجل.",
        "done_log": "—— اكتمل: {} نجح، {} فشل ——",
        "merged": "📑 ملف مدمج",
        "ocr_hdr": "النص المستخرج (OCR)",
        "no_preview": "لا توجد نتائج للمعاينة بعد — نفّذ تحويلاً أولاً.",
        "menu_lang": "English",
        "menu_update": "التحقق من التحديثات",
        "update_new": "إصدار جديد متاح: {} (الحالي: {}). فتح صفحة التنزيل؟",
        "update_none": "لديك أحدث إصدار ({}) ✅",
        "update_err": "تعذّر التحقق من التحديثات (تحقق من الاتصال).",
        "drop_hint": "⬇ اسحب الملفات وأفلتها هنا ⬇",
        "url_invalid": "الرجاء إدخال رابط يبدأ بـ http:// أو https://",
        "credit": "تطوير: Jaber Aleida",
        "menu_about": "عن البرنامج",
        "about_text": "MarkItDown {}\n\nأداة لتحويل الملفات إلى Markdown\nمبنية على مكتبة markitdown من Microsoft (MIT)\n\nتطوير وتوقيع: Jaber Aleida",
    },
    "en": {
        "title": "MarkItDown — Convert files to Markdown",
        "subtitle": "Convert your documents to Markdown. Drag files onto the list or use the buttons.",
        "supported": "Supported: PDF, Word, PowerPoint, Excel, HTML, CSV, JSON, XML, ZIP, EPub, images (with OCR), audio, web pages & YouTube links",
        "files_frame": "Selected files (drag & drop here)",
        "add_files": "➕ Add files…",
        "add_folder": "📁 Add folder…",
        "remove": "Remove selected",
        "clear": "Clear all",
        "url_placeholder": "…or paste a web page / YouTube URL here",
        "convert_url": "Convert URL",
        "out_frame": "Output location",
        "same_folder": "Next to the source file",
        "custom_folder": "Custom folder:",
        "browse": "Browse…",
        "merge": "Merge all results into a single Markdown file",
        "convert": "🔄 Convert",
        "preview": "👁 Preview result",
        "log_frame": "Log",
        "need_file": "Add at least one file or URL first.",
        "need_out": "Choose an output folder first.",
        "converting": "Converting",
        "init": "Initializing converter…",
        "ok": "✔ done",
        "fail": "✖ failed",
        "done_ok": "Converted {} file(s) successfully ✅",
        "done_mix": "{} succeeded, {} failed. See log.",
        "done_log": "—— finished: {} ok, {} failed ——",
        "merged": "📑 merged file",
        "ocr_hdr": "Extracted text (OCR)",
        "no_preview": "Nothing to preview yet — run a conversion first.",
        "menu_lang": "العربية",
        "menu_update": "Check for updates",
        "update_new": "New version available: {} (current: {}). Open download page?",
        "update_none": "You have the latest version ({}) ✅",
        "update_err": "Could not check for updates (check connection).",
        "drop_hint": "⬇ Drag & drop files here ⬇",
        "url_invalid": "Please enter a URL starting with http:// or https://",
        "credit": "Developed by Jaber Aleida",
        "menu_about": "About",
        "about_text": "MarkItDown {}\n\nConvert files to Markdown\nBuilt on Microsoft's markitdown library (MIT)\n\nDeveloped & signed by: Jaber Aleida",
    },
}


def load_config() -> dict:
    try:
        with open(CONFIG_PATH, "r", encoding="utf-8") as f:
            return json.load(f)
    except Exception:  # noqa: BLE001
        return {}


def save_config(cfg: dict):
    try:
        with open(CONFIG_PATH, "w", encoding="utf-8") as f:
            json.dump(cfg, f, ensure_ascii=False, indent=2)
    except Exception:  # noqa: BLE001
        pass


_CFG = load_config()
LANG = _CFG.get("lang", "ar") if _CFG.get("lang", "ar") in STRINGS else "ar"


def t(key: str) -> str:
    return STRINGS[LANG].get(key, STRINGS["en"].get(key, key))


# ---------------- conversion core ----------------
_converter = None
_ocr_engine = None


def get_converter() -> MarkItDown:
    global _converter
    if _converter is None:
        _converter = MarkItDown(enable_plugins=False)
    return _converter


def get_ocr():
    global _ocr_engine
    if _ocr_engine is None and HAS_OCR:
        _ocr_engine = RapidOCR()
    return _ocr_engine


def unique_out_path(target_dir: str, base: str, ext: str = ".md") -> str:
    out_path = os.path.join(target_dir, base + ext)
    n = 1
    while os.path.exists(out_path):
        out_path = os.path.join(target_dir, f"{base}-{n}{ext}")
        n += 1
    return out_path


def ocr_image(path: str) -> str:
    engine = get_ocr()
    if engine is None:
        return ""
    try:
        result, _ = engine(path)
        if not result:
            return ""
        return "\n".join(line[1] for line in result if len(line) > 1)
    except Exception:  # noqa: BLE001
        return ""


def convert_one(path: str, out_dir: str = None) -> str:
    """Convert one file; returns the .md output path."""
    md = get_converter()
    result = md.convert(path)
    text = result.text_content or ""

    ext = os.path.splitext(path)[1].lower()
    if ext in IMAGE_EXTS:
        ocr_text = ocr_image(path)
        if ocr_text.strip():
            text = f"{text}\n\n## {t('ocr_hdr')}\n\n{ocr_text}\n"

    target_dir = out_dir or os.path.dirname(os.path.abspath(path))
    base = os.path.splitext(os.path.basename(path))[0]
    out_path = unique_out_path(target_dir, base)
    with open(out_path, "w", encoding="utf-8") as f:
        f.write(text)
    return out_path


def convert_url(url: str, out_dir: str) -> str:
    md = get_converter()
    result = md.convert(url)
    slug = re.sub(r"[^A-Za-z0-9ء-ي]+", "-", url.split("//")[-1]).strip("-")[:60] or "page"
    out_path = unique_out_path(out_dir, slug)
    with open(out_path, "w", encoding="utf-8") as f:
        f.write(result.text_content or "")
    return out_path


def merge_outputs(paths: list, out_dir: str) -> str:
    merged = unique_out_path(out_dir, "merged")
    with open(merged, "w", encoding="utf-8") as out:
        for p in paths:
            out.write(f"\n\n---\n\n# {os.path.basename(p)}\n\n")
            with open(p, "r", encoding="utf-8") as f:
                out.write(f.read())
    return merged


# ---------------- update check ----------------
def check_update():
    """Return (latest_version, download_url) or None."""
    try:
        import urllib.request
        req = urllib.request.Request(
            f"https://api.github.com/repos/{GITHUB_REPO}/releases/latest",
            headers={"Accept": "application/vnd.github+json", "User-Agent": "markitdown-app"},
        )
        with urllib.request.urlopen(req, timeout=8) as r:
            data = json.loads(r.read().decode("utf-8"))
        latest = data.get("tag_name", "").lstrip("v")
        if not latest or latest == __version__:
            return None
        system = platform.system()
        want = ".dmg" if system == "Darwin" else ".exe"
        url = data.get("html_url", "")
        for asset in data.get("assets", []):
            if asset.get("name", "").endswith(want):
                url = asset.get("browser_download_url", url)
                break
        return latest, url
    except Exception:  # noqa: BLE001
        return False


# ---------------- desktop GUI ----------------
class MarkItDownApp(_TkBase):
    def __init__(self):
        super().__init__()
        self.files: list[str] = []
        self.outputs: list[str] = []
        self.jobs: queue.Queue = queue.Queue()
        self.output_dir = tk.StringVar(value=_CFG.get("output_dir", ""))
        self.same_folder = tk.BooleanVar(value=_CFG.get("same_folder", True))
        self.merge_var = tk.BooleanVar(value=_CFG.get("merge", False))

        self.title(t("title"))
        self.geometry("760x640")
        self.minsize(640, 520)

        self._build_menu()
        self._build_ui()
        self._setup_dnd()
        self.after(100, self._poll_jobs)
        self.protocol("WM_DELETE_WINDOW", self._on_close)
        threading.Thread(target=self._silent_update_check, daemon=True).start()

    # ---------- menu ----------
    def _build_menu(self):
        menubar = tk.Menu(self)
        app_menu = tk.Menu(menubar, tearoff=0)
        app_menu.add_command(label=t("menu_lang"), command=self.toggle_language)
        app_menu.add_command(label=t("menu_update"), command=self.manual_update_check)
        app_menu.add_separator()
        app_menu.add_command(label=t("menu_about"), command=self.show_about)
        menubar.add_cascade(label="MarkItDown", menu=app_menu)
        self.config(menu=menubar)
        self._menubar = menubar
        self._app_menu = app_menu

    def toggle_language(self):
        global LANG
        LANG = "en" if LANG == "ar" else "ar"
        self._persist()
        # إعادة بناء الواجهة باللغة الجديدة
        for w in self.winfo_children():
            w.destroy()
        self._build_menu()
        self._build_ui()
        self._setup_dnd()
        self.title(t("title"))

    # ---------- UI ----------
    def _build_ui(self):
        pad = {"padx": 10, "pady": 5}

        ttk.Label(self, text="MarkItDown", font=("Helvetica", 22, "bold")).pack(anchor="w", **pad)
        self.lbl_sub = ttk.Label(self, text=t("subtitle") + "\n" + t("supported"),
                                 justify="left", foreground="#444444", wraplength=720)
        self.lbl_sub.pack(anchor="w", padx=10)
        ttk.Label(self, text=t("credit"), justify="left",
                  foreground="#888888", font=("Helvetica", 10)).pack(anchor="w", padx=10)

        # file list
        self.list_frame = ttk.LabelFrame(self, text=t("files_frame"))
        self.list_frame.pack(fill="both", expand=True, **pad)
        self.listbox = tk.Listbox(self.list_frame, selectmode=tk.EXTENDED, activestyle="none")
        scroll = ttk.Scrollbar(self.list_frame, orient="vertical", command=self.listbox.yview)
        self.listbox.configure(yscrollcommand=scroll.set)
        self.listbox.pack(side="left", fill="both", expand=True, padx=(8, 0), pady=8)
        scroll.pack(side="right", fill="y", padx=(0, 8), pady=8)
        self.listbox.insert("end", t("drop_hint") if HAS_DND else "")

        btns = ttk.Frame(self)
        btns.pack(fill="x", **pad)
        self.btn_add = ttk.Button(btns, text=t("add_files"), command=self.add_files)
        self.btn_add.pack(side="left")
        self.btn_folder = ttk.Button(btns, text=t("add_folder"), command=self.add_folder)
        self.btn_folder.pack(side="left", padx=6)
        self.btn_remove = ttk.Button(btns, text=t("remove"), command=self.remove_selected)
        self.btn_remove.pack(side="left")
        self.btn_clear = ttk.Button(btns, text=t("clear"), command=self.clear_all)
        self.btn_clear.pack(side="left", padx=6)

        # URL row
        url_frame = ttk.Frame(self)
        url_frame.pack(fill="x", **pad)
        self.url_entry = ttk.Entry(url_frame)
        self.url_entry.insert(0, t("url_placeholder"))
        self.url_entry.configure(foreground="#888888")
        self.url_entry.bind("<FocusIn>", self._url_focus_in)
        self.url_entry.bind("<FocusOut>", self._url_focus_out)
        self.url_entry.pack(side="left", fill="x", expand=True)
        self.btn_url = ttk.Button(url_frame, text=t("convert_url"), command=self.convert_url_clicked)
        self.btn_url.pack(side="left", padx=6)
        self._url_has_hint = True

        # output options
        self.out_frame = ttk.LabelFrame(self, text=t("out_frame"))
        self.out_frame.pack(fill="x", **pad)
        self.rb_same = ttk.Radiobutton(self.out_frame, text=t("same_folder"),
                                       variable=self.same_folder, value=True, command=self._toggle_out)
        self.rb_same.pack(side="left", padx=8, pady=6)
        self.rb_custom = ttk.Radiobutton(self.out_frame, text=t("custom_folder"),
                                         variable=self.same_folder, value=False, command=self._toggle_out)
        self.rb_custom.pack(side="left", padx=8)
        self.out_entry = ttk.Entry(self.out_frame, textvariable=self.output_dir, width=26)
        self.out_entry.pack(side="left", padx=4)
        self.out_btn = ttk.Button(self.out_frame, text=t("browse"), command=self.pick_output)
        self.out_btn.pack(side="left", padx=4)
        self._toggle_out()

        opt = ttk.Frame(self)
        opt.pack(fill="x", padx=10)
        self.chk_merge = ttk.Checkbutton(opt, text=t("merge"), variable=self.merge_var)
        self.chk_merge.pack(side="left")

        # actions
        action = ttk.Frame(self)
        action.pack(fill="x", **pad)
        self.convert_btn = ttk.Button(action, text=t("convert"), command=self.start_conversion)
        self.convert_btn.pack(side="left")
        self.preview_btn = ttk.Button(action, text=t("preview"), command=self.show_preview)
        self.preview_btn.pack(side="left", padx=6)
        self.progress = ttk.Progressbar(action, mode="determinate")
        self.progress.pack(side="left", fill="x", expand=True, padx=10)

        self.log_frame = ttk.LabelFrame(self, text=t("log_frame"))
        self.log_frame.pack(fill="both", expand=True, **pad)
        self.log = tk.Text(self.log_frame, height=7, state="disabled", wrap="word")
        self.log.pack(fill="both", expand=True, padx=8, pady=8)

    # ---------- drag & drop ----------
    def _setup_dnd(self):
        if not HAS_DND:
            return
        try:
            self.listbox.drop_target_register(DND_FILES)
            self.listbox.dnd_bind("<<Drop>>", self._on_drop)
        except Exception:  # noqa: BLE001
            pass

    def _on_drop(self, event):
        try:
            paths = self.tk.splitlist(event.data)
        except Exception:  # noqa: BLE001
            paths = [event.data]
        for p in paths:
            if os.path.isdir(p):
                self._add_folder_path(p)
            elif os.path.isfile(p):
                self._add_file(p)

    # ---------- helpers ----------
    def _persist(self):
        save_config({
            "lang": LANG,
            "same_folder": self.same_folder.get(),
            "output_dir": self.output_dir.get(),
            "merge": self.merge_var.get(),
        })

    def _on_close(self):
        self._persist()
        self.destroy()

    def _toggle_out(self):
        state = "disabled" if self.same_folder.get() else "normal"
        self.out_entry.configure(state=state)
        self.out_btn.configure(state=state)

    def _url_focus_in(self, _):
        if self._url_has_hint:
            self.url_entry.delete(0, "end")
            self.url_entry.configure(foreground="#000000")
            self._url_has_hint = False

    def _url_focus_out(self, _):
        if not self.url_entry.get().strip():
            self.url_entry.insert(0, t("url_placeholder"))
            self.url_entry.configure(foreground="#888888")
            self._url_has_hint = True

    def _refresh_placeholder(self):
        self.listbox.delete(0, "end")
        if not self.files and HAS_DND:
            self.listbox.insert("end", t("drop_hint"))
        for p in self.files:
            self.listbox.insert("end", os.path.basename(p) + f"    —    {os.path.dirname(p)}")

    def _add_file(self, path: str):
        if path not in self.files:
            self.files.append(path)
            self._refresh_placeholder()

    def log_msg(self, msg: str):
        self.log.configure(state="normal")
        self.log.insert("end", msg + "\n")
        self.log.see("end")
        self.log.configure(state="disabled")

    # ---------- adding ----------
    def add_files(self):
        for p in filedialog.askopenfilenames(title=t("add_files")):
            self._add_file(p)

    def add_folder(self):
        d = filedialog.askdirectory(title=t("add_folder"))
        if d:
            self._add_folder_path(d)

    def _add_folder_path(self, folder: str):
        added = 0
        for root, _dirs, names in os.walk(folder):
            for name in names:
                if os.path.splitext(name)[1].lower() in SUPPORTED_EXTS:
                    self._add_file(os.path.join(root, name))
                    added += 1
        self.log_msg(f"📁 {folder}: {added} file(s)")

    def remove_selected(self):
        for idx in reversed(self.listbox.curselection()):
            if idx < len(self.files):
                del self.files[idx]
        self._refresh_placeholder()

    def clear_all(self):
        self.files.clear()
        self._refresh_placeholder()

    def pick_output(self):
        d = filedialog.askdirectory(title=t("browse"))
        if d:
            self.output_dir.set(d)

    # ---------- URL ----------
    def convert_url_clicked(self):
        url = "" if self._url_has_hint else self.url_entry.get().strip()
        if not (url.startswith("http://") or url.startswith("https://")):
            messagebox.showinfo("MarkItDown", t("url_invalid"))
            return
        out_dir = None if self.same_folder.get() else self.output_dir.get()
        if out_dir is None:
            out_dir = os.path.expanduser("~/Documents")
        elif not out_dir:
            messagebox.showinfo("MarkItDown", t("need_out"))
            return
        self.convert_btn.configure(state="disabled")
        self.progress.configure(maximum=1, value=0)
        threading.Thread(target=self._url_worker, args=(url, out_dir), daemon=True).start()

    def _url_worker(self, url, out_dir):
        try:
            self.jobs.put(("log", f"🌐 {url}"))
            out = convert_url(url, out_dir)
            self.jobs.put(("output", out))
            self.jobs.put(("log", f"  {t('ok')}: {out}"))
            self.jobs.put(("step", None))
            self.jobs.put(("done", (1, 0)))
        except Exception as e:  # noqa: BLE001
            self.jobs.put(("log", f"  {t('fail')}: {e}"))
            self.jobs.put(("step", None))
            self.jobs.put(("done", (0, 1)))

    # ---------- conversion ----------
    def start_conversion(self):
        if not self.files:
            messagebox.showinfo("MarkItDown", t("need_file"))
            return
        if not self.same_folder.get() and not self.output_dir.get():
            messagebox.showinfo("MarkItDown", t("need_out"))
            return
        self._persist()
        self.convert_btn.configure(state="disabled")
        self.progress.configure(maximum=len(self.files), value=0)
        threading.Thread(
            target=self._worker,
            args=(list(self.files), None if self.same_folder.get() else self.output_dir.get(),
                  self.merge_var.get()),
            daemon=True,
        ).start()

    def _worker(self, files, out_dir, merge):
        try:
            self.jobs.put(("log", t("init")))
            get_converter()
            ok, fail, produced = 0, 0, []
            for path in files:
                try:
                    self.jobs.put(("log", f"{t('converting')}: {os.path.basename(path)} …"))
                    out = convert_one(path, out_dir)
                    produced.append(out)
                    self.jobs.put(("output", out))
                    self.jobs.put(("log", f"  {t('ok')}: {out}"))
                    ok += 1
                except Exception as e:  # noqa: BLE001
                    self.jobs.put(("log", f"  {t('fail')} {os.path.basename(path)}: {e}"))
                    fail += 1
                self.jobs.put(("step", None))
            if merge and produced:
                target = out_dir or os.path.dirname(os.path.abspath(produced[0]))
                merged = merge_outputs(produced, target)
                self.jobs.put(("output", merged))
                self.jobs.put(("log", f"{t('merged')}: {merged}"))
            self.jobs.put(("done", (ok, fail)))
        except Exception:
            self.jobs.put(("log", traceback.format_exc()))
            self.jobs.put(("done", (0, len(files))))

    def _poll_jobs(self):
        try:
            while True:
                kind, payload = self.jobs.get_nowait()
                if kind == "log":
                    self.log_msg(payload)
                elif kind == "output":
                    self.outputs.append(payload)
                elif kind == "step":
                    self.progress["value"] += 1
                elif kind == "done":
                    ok, fail = payload
                    self.convert_btn.configure(state="normal")
                    self.log_msg(t("done_log").format(ok, fail))
                    if fail == 0 and ok:
                        messagebox.showinfo("MarkItDown", t("done_ok").format(ok))
                    elif fail:
                        messagebox.showwarning("MarkItDown", t("done_mix").format(ok, fail))
                elif kind == "update_msg":
                    tag, data = payload
                    if tag == "err":
                        messagebox.showerror("MarkItDown", t("update_err"))
                    elif tag == "none":
                        messagebox.showinfo("MarkItDown", t("update_none").format(__version__))
                    else:
                        latest, url = data
                        if messagebox.askyesno("MarkItDown", t("update_new").format(latest, __version__)):
                            webbrowser.open(url)
        except queue.Empty:
            pass
        self.after(100, self._poll_jobs)

    # ---------- preview ----------
    def show_preview(self):
        if not self.outputs:
            messagebox.showinfo("MarkItDown", t("no_preview"))
            return
        win = tk.Toplevel(self)
        win.title(t("preview"))
        win.geometry("700x500")
        top = ttk.Frame(win)
        top.pack(fill="x", padx=8, pady=6)
        combo = ttk.Combobox(top, values=self.outputs, state="readonly")
        combo.pack(fill="x")
        combo.current(len(self.outputs) - 1)
        text = tk.Text(win, wrap="word")
        text.pack(fill="both", expand=True, padx=8, pady=8)

        def load(_=None):
            try:
                with open(combo.get(), "r", encoding="utf-8") as f:
                    text.delete("1.0", "end")
                    text.insert("1.0", f.read())
            except Exception as e:  # noqa: BLE001
                text.delete("1.0", "end")
                text.insert("1.0", str(e))

        combo.bind("<<ComboboxSelected>>", load)
        load()

    # ---------- about ----------
    def show_about(self):
        messagebox.showinfo(t("menu_about"), t("about_text").format(__version__))

    # ---------- updates ----------
    def _silent_update_check(self):
        res = check_update()
        if res and res is not False:
            latest, url = res
            self.jobs.put(("log", f"🆕 v{latest}: {url}"))

    def manual_update_check(self):
        def run():
            res = check_update()
            if res is False:
                self.jobs.put(("update_msg", ("err", None)))
            elif res is None:
                self.jobs.put(("update_msg", ("none", None)))
            else:
                self.jobs.put(("update_msg", ("new", res)))
        threading.Thread(target=run, daemon=True).start()


# ---------------- local web UI ----------------
WEB_HTML = """<!doctype html>
<html lang="ar" dir="rtl">
<head>
<meta charset="utf-8">
<title>MarkItDown</title>
<style>
  body { font-family: -apple-system, "Segoe UI", Tahoma, sans-serif; max-width: 760px;
         margin: 40px auto; padding: 0 16px; color: #222; }
  h1 { margin-bottom: 4px; }
  .sub { color: #666; font-size: 14px; margin-bottom: 20px; }
  #drop { border: 2px dashed #4a90c2; border-radius: 14px; padding: 46px;
          text-align: center; color: #4a90c2; font-size: 18px; background: #f7fbfe; }
  #drop.over { background: #e3f0fa; }
  #url { width: 70%; padding: 8px; border: 1px solid #bbb; border-radius: 8px; }
  button { padding: 9px 18px; border: 0; border-radius: 8px; background: #2d7fb8;
           color: #fff; font-size: 15px; cursor: pointer; }
  label { font-size: 14px; }
  #out { margin-top: 22px; font-family: ui-monospace, monospace; font-size: 13px;
         white-space: pre-wrap; background: #fafafa; border: 1px solid #eee;
         border-radius: 8px; padding: 12px; min-height: 60px; }
  input[type=file] { display: none; }
</style>
</head>
<body>
<h1>MarkItDown</h1>
<div class="sub">حوّل PDF / Word / PowerPoint / Excel / HTML / الصور (OCR) إلى Markdown</div>
<div id="drop">⬇ اسحب الملفات هنا أو انقر للاختيار ⬇</div>
<input id="file" type="file" multiple>
<p>
  <input id="url" placeholder="أو ألصق رابط صفحة / YouTube هنا">
  <button onclick="convertUrl()">تحويل الرابط</button>
</p>
<p><label><input type="checkbox" id="merge"> دمج كل النتائج في ملف واحد</label></p>
<div id="out">النتائج تظهر هنا…</div>
<p style="color:#999;font-size:12px;margin-top:24px">تطوير: Jaber Aleida — MarkItDown 2.0</p>
<script>
const drop = document.getElementById('drop'), file = document.getElementById('file'),
      out = document.getElementById('out');
drop.onclick = () => file.click();
drop.ondragover = e => { e.preventDefault(); drop.classList.add('over'); };
drop.ondragleave = () => drop.classList.remove('over');
drop.ondrop = e => { e.preventDefault(); drop.classList.remove('over'); send(e.dataTransfer.files); };
file.onchange = () => send(file.files);

async function send(files) {
  if (!files.length) return;
  const fd = new FormData();
  for (const f of files) fd.append('files', f);
  fd.append('merge', document.getElementById('merge').checked);
  out.textContent = '⏳ جارٍ التحويل…';
  const r = await fetch('/api/convert', { method: 'POST', body: fd });
  show(await r.json());
}
async function convertUrl() {
  const u = document.getElementById('url').value.trim();
  if (!/^https?:\\/\\//.test(u)) { out.textContent = 'أدخل رابطاً صحيحاً'; return; }
  const fd = new FormData(); fd.append('url', u);
  out.textContent = '⏳ جارٍ التحويل…';
  const r = await fetch('/api/convert', { method: 'POST', body: fd });
  show(await r.json());
}
function show(res) {
  out.textContent = res.results.map(x =>
    x.ok ? ('✔ ' + x.name + '\\n   → ' + x.out) : ('✖ ' + x.name + ': ' + x.error)
  ).join('\\n') + (res.merged ? ('\\n📑 مدمج: ' + res.merged) : '');
}
</script>
</body>
</html>
"""


def run_web(host: str = "127.0.0.1", port: int = 8741):
    from fastapi import FastAPI, Request
    from fastapi.responses import HTMLResponse, JSONResponse
    import uvicorn

    app = FastAPI(title="MarkItDown")
    out_root = _CFG.get("output_dir") or os.path.join(os.path.expanduser("~"), "MarkItDown-Converted")
    os.makedirs(out_root, exist_ok=True)

    @app.get("/", response_class=HTMLResponse)
    def index():
        return WEB_HTML

    @app.post("/api/convert")
    async def api_convert(request: Request):
        form = await request.form()
        results, produced = [], []
        merge = str(form.get("merge", "")).lower() == "true"
        url = form.get("url")
        try:
            if url:
                out = convert_url(str(url), out_root)
                results.append({"name": str(url), "ok": True, "out": out})
                produced.append(out)
            for up in form.getlist("files"):
                filename = getattr(up, "filename", None)
                if not filename:
                    continue
                suffix = os.path.splitext(filename)[1]
                fd, tmp = tempfile.mkstemp(suffix=suffix)
                with os.fdopen(fd, "wb") as f:
                    f.write(await up.read())
                try:
                    # احتفظ بالاسم الأصلي للمخرجات
                    named = os.path.join(tempfile.gettempdir(), filename)
                    os.replace(tmp, named)
                    out = convert_one(named, out_root)
                    results.append({"name": filename, "ok": True, "out": out})
                    produced.append(out)
                except Exception as e:  # noqa: BLE001
                    results.append({"name": filename, "ok": False, "error": str(e)})
        except Exception as e:  # noqa: BLE001
            results.append({"name": "url", "ok": False, "error": str(e)})
        merged = merge_outputs(produced, out_root) if merge and len(produced) > 1 else None
        return JSONResponse({"results": results, "merged": merged, "out_dir": out_root})

    print(f"MarkItDown web UI → http://{host}:{port}  (outputs: {out_root})")
    webbrowser.open(f"http://{host}:{port}")
    uvicorn.run(app, host=host, port=port, log_level="warning")


# ---------------- macOS menu bar ----------------
def run_menubar():
    try:
        import rumps
    except Exception:  # noqa: BLE001
        print("rumps غير متوفر — شغّل التطبيق بدون --menubar")
        sys.exit(1)

    class MenuBarApp(rumps.App):
        def __init__(self):
            super().__init__("M↓", title=None, quit_button=None)
            self.menu = ["فتح نافذة التحويل", "تحويل رابط من الحافظة", "فتح واجهة الويب", "إنهاء"]

        @rumps.clicked("فتح نافذة التحويل")
        def open_gui(self, _):
            if getattr(sys, "frozen", False):
                subprocess.Popen([sys.executable])
            else:
                subprocess.Popen([sys.executable, os.path.abspath(__file__)])

        @rumps.clicked("تحويل رابط من الحافظة")
        def from_clipboard(self, _):
            try:
                url = subprocess.check_output(["pbpaste"]).decode("utf-8").strip()
                if not url.startswith("http"):
                    rumps.notification("MarkItDown", "لا يوجد رابط", "الحافظة لا تحتوي رابط http")
                    return
                out_dir = _CFG.get("output_dir") or os.path.expanduser("~/Documents")
                out = convert_url(url, out_dir)
                rumps.notification("MarkItDown", "تم التحويل ✔", os.path.basename(out))
            except Exception as e:  # noqa: BLE001
                rumps.notification("MarkItDown", "فشل التحويل", str(e)[:200])

        @rumps.clicked("فتح واجهة الويب")
        def open_web(self, _):
            if getattr(sys, "frozen", False):
                subprocess.Popen([sys.executable, "--web"])
            else:
                subprocess.Popen([sys.executable, os.path.abspath(__file__), "--web"])

        @rumps.clicked("إنهاء")
        def quit_app(self, _):
            rumps.quit_application()

    MenuBarApp().run()


# ---------------- entry ----------------
def _selftest():
    failures = 0
    for path in sys.argv[2:]:
        try:
            out = convert_one(path, tempfile.mkdtemp())
            size = os.path.getsize(out)
            print(f"OK   {path} -> {size} bytes")
        except Exception as e:  # noqa: BLE001
            print(f"FAIL {path}: {e}")
            failures += 1
    print("SELFTEST", "FAILED" if failures else "PASSED")
    sys.exit(1 if failures else 0)


def main():
    args = sys.argv[1:]
    if args and args[0] == "--selftest":
        _selftest()
    elif args and args[0] == "--web":
        run_web()
    elif args and args[0] == "--menubar":
        run_menubar()
    else:
        MarkItDownApp().mainloop()


if __name__ == "__main__":
    main()
