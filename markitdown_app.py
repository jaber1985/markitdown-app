#!/usr/bin/env python3
"""MarkItDown — macOS GUI wrapper.

Convert documents (PDF, Word, PowerPoint, Excel, HTML, images, audio, ...)
into Markdown files. Based on Microsoft's markitdown library.
"""

import os
import queue
import sys
import threading
import traceback
import tkinter as tk
from tkinter import ttk, filedialog, messagebox

from markitdown import MarkItDown

APP_TITLE = "MarkItDown — تحويل الملفات إلى Markdown"

SUPPORTED = (
    "PDF, Word (docx), PowerPoint (pptx), Excel (xlsx/xls), HTML, CSV, JSON, XML, "
    "صور (مع وصف EXIF), صوت (مع بيانات EXIF), ZIP, EPub, وروابط YouTube"
)


class MarkItDownApp(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title(APP_TITLE)
        self.geometry("720x560")
        self.minsize(620, 480)

        self.files: list[str] = []
        self.output_dir = tk.StringVar(value="")
        self.same_folder = tk.BooleanVar(value=True)
        self.jobs: queue.Queue = queue.Queue()
        self.converter = None

        self._build_ui()
        self.after(100, self._poll_jobs)

    # ---------------- UI ----------------
    def _build_ui(self):
        pad = {"padx": 10, "pady": 6}

        header = ttk.Label(
            self,
            text="MarkItDown",
            font=("Helvetica", 22, "bold"),
        )
        header.pack(anchor="w", **pad)
        ttk.Label(
            self,
            text="حوّل مستنداتك إلى ملفات Markdown بضغطة واحدة.\nالصيغ المدعومة: " + SUPPORTED,
            justify="left",
            foreground="#444444",
        ).pack(anchor="w", padx=10)

        # File list
        list_frame = ttk.LabelFrame(self, text="الملفات المحددة")
        list_frame.pack(fill="both", expand=True, **pad)

        self.listbox = tk.Listbox(list_frame, selectmode=tk.EXTENDED, activestyle="none")
        scroll = ttk.Scrollbar(list_frame, orient="vertical", command=self.listbox.yview)
        self.listbox.configure(yscrollcommand=scroll.set)
        self.listbox.pack(side="left", fill="both", expand=True, padx=(8, 0), pady=8)
        scroll.pack(side="right", fill="y", padx=(0, 8), pady=8)

        btns = ttk.Frame(self)
        btns.pack(fill="x", **pad)
        ttk.Button(btns, text="➕ إضافة ملفات…", command=self.add_files).pack(side="left")
        ttk.Button(btns, text="إزالة المحدد", command=self.remove_selected).pack(side="left", padx=6)
        ttk.Button(btns, text="مسح الكل", command=self.clear_all).pack(side="left")

        # Output options
        out_frame = ttk.LabelFrame(self, text="مكان الحفظ")
        out_frame.pack(fill="x", **pad)
        ttk.Radiobutton(
            out_frame,
            text="بجانب الملف الأصلي",
            variable=self.same_folder,
            value=True,
            command=self._toggle_out,
        ).pack(side="left", padx=8, pady=6)
        ttk.Radiobutton(
            out_frame,
            text="مجلد محدد:",
            variable=self.same_folder,
            value=False,
            command=self._toggle_out,
        ).pack(side="left", padx=8)
        self.out_entry = ttk.Entry(out_frame, textvariable=self.output_dir, state="disabled", width=28)
        self.out_entry.pack(side="left", padx=4)
        self.out_btn = ttk.Button(out_frame, text="اختيار…", command=self.pick_output, state="disabled")
        self.out_btn.pack(side="left", padx=4)

        # Convert
        action = ttk.Frame(self)
        action.pack(fill="x", **pad)
        self.convert_btn = ttk.Button(action, text="🔄 تحويل", command=self.start_conversion)
        self.convert_btn.pack(side="left")
        self.progress = ttk.Progressbar(action, mode="determinate")
        self.progress.pack(side="left", fill="x", expand=True, padx=10)

        # Log
        log_frame = ttk.LabelFrame(self, text="السجل")
        log_frame.pack(fill="both", expand=True, **pad)
        self.log = tk.Text(log_frame, height=8, state="disabled", wrap="word")
        self.log.pack(fill="both", expand=True, padx=8, pady=8)

    def _toggle_out(self):
        state = "disabled" if self.same_folder.get() else "normal"
        self.out_entry.configure(state=state)
        self.out_btn.configure(state=state)

    def log_msg(self, msg: str):
        self.log.configure(state="normal")
        self.log.insert("end", msg + "\n")
        self.log.see("end")
        self.log.configure(state="disabled")

    # ---------------- File handling ----------------
    def add_files(self):
        paths = filedialog.askopenfilenames(title="اختر الملفات للتحويل")
        for p in paths:
            if p not in self.files:
                self.files.append(p)
                self.listbox.insert("end", os.path.basename(p) + f"    —    {os.path.dirname(p)}")

    def remove_selected(self):
        for idx in reversed(self.listbox.curselection()):
            self.listbox.delete(idx)
            del self.files[idx]

    def clear_all(self):
        self.files.clear()
        self.listbox.delete(0, "end")

    def pick_output(self):
        d = filedialog.askdirectory(title="اختر مجلد الحفظ")
        if d:
            self.output_dir.set(d)

    # ---------------- Conversion ----------------
    def start_conversion(self):
        if not self.files:
            messagebox.showinfo(APP_TITLE, "أضف ملفاً واحداً على الأقل أولاً.")
            return
        if not self.same_folder.get() and not self.output_dir.get():
            messagebox.showinfo(APP_TITLE, "اختر مجلد الحفظ أولاً.")
            return

        self.convert_btn.configure(state="disabled")
        self.progress.configure(maximum=len(self.files), value=0)
        files = list(self.files)
        out_dir = None if self.same_folder.get() else self.output_dir.get()
        threading.Thread(target=self._worker, args=(files, out_dir), daemon=True).start()

    def _worker(self, files, out_dir):
        try:
            if self.converter is None:
                self.jobs.put(("log", "جارٍ تهيئة المحوّل…"))
                self.converter = MarkItDown(enable_plugins=False)
            ok, fail = 0, 0
            for path in files:
                try:
                    self.jobs.put(("log", f"تحويل: {os.path.basename(path)} …"))
                    result = self.converter.convert(path)
                    target_dir = out_dir or os.path.dirname(os.path.abspath(path))
                    base = os.path.splitext(os.path.basename(path))[0]
                    out_path = os.path.join(target_dir, base + ".md")
                    n = 1
                    while os.path.exists(out_path):
                        out_path = os.path.join(target_dir, f"{base}-{n}.md")
                        n += 1
                    with open(out_path, "w", encoding="utf-8") as f:
                        f.write(result.text_content)
                    self.jobs.put(("log", f"  ✔ تم: {out_path}"))
                    ok += 1
                except Exception as e:  # noqa: BLE001
                    self.jobs.put(("log", f"  ✖ فشل {os.path.basename(path)}: {e}"))
                    fail += 1
                self.jobs.put(("step", None))
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
                elif kind == "step":
                    self.progress["value"] += 1
                elif kind == "done":
                    ok, fail = payload
                    self.convert_btn.configure(state="normal")
                    self.log_msg(f"—— اكتمل: {ok} نجح، {fail} فشل ——")
                    if fail == 0:
                        messagebox.showinfo(APP_TITLE, f"تم تحويل {ok} ملف بنجاح ✅")
                    else:
                        messagebox.showwarning(APP_TITLE, f"نجح {ok} وفشل {fail}. راجع السجل.")
        except queue.Empty:
            pass
        self.after(100, self._poll_jobs)


def _selftest():
    """Headless smoke test: python markitdown_app.py --selftest <file> [file...]"""
    md = MarkItDown(enable_plugins=False)
    failures = 0
    for path in sys.argv[2:]:
        try:
            result = md.convert(path)
            print(f"OK   {path} -> {len(result.text_content)} chars")
        except Exception as e:  # noqa: BLE001
            print(f"FAIL {path}: {e}")
            failures += 1
    print("SELFTEST", "FAILED" if failures else "PASSED")
    sys.exit(1 if failures else 0)


if __name__ == "__main__":
    if len(sys.argv) > 1 and sys.argv[1] == "--selftest":
        _selftest()
    app = MarkItDownApp()
    app.mainloop()
