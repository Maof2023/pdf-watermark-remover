# -*- coding: utf-8 -*-
"""PDF 去水印转 Word —— 图形界面（Python 自带 tkinter，无需额外安装 UI 库）。

用法：python app.py
"""
from __future__ import annotations

import logging
import os
import queue
import subprocess
import sys
import threading
import tkinter as tk
from pathlib import Path
from tkinter import filedialog, messagebox, scrolledtext, ttk

from core import batch_convert, find_repeated_texts

APP_TITLE = "PDF 去水印转 Word 工具"


class QueueLogHandler(logging.Handler):
    """把 pdf2docx 等库的日志转发到 UI 消息队列（线程安全）。"""

    def __init__(self, q: queue.Queue):
        super().__init__()
        self.q = q

    def emit(self, record: logging.LogRecord) -> None:
        try:
            self.q.put(("log", self.format(record)))
        except Exception:
            pass


class App:
    def __init__(self, root: tk.Tk):
        self.root = root
        self.root.title(APP_TITLE)
        self.root.geometry("840x620")
        self.root.minsize(720, 520)

        self.pdfs = []
        self.q = queue.Queue()
        self.worker = None
        self.last_out = None

        self.src_var = tk.StringVar(value="未选择（可多选 PDF 文件或整个文件夹）")
        self.out_var = tk.StringVar()
        self.wm_var = tk.StringVar()
        self.auto_var = tk.BooleanVar(value=True)

        self._build()
        self.root.after(100, self._drain)

    # ------------------------------------------------ UI 构建
    def _build(self):
        pad = {"padx": 8, "pady": 4}
        top = ttk.Frame(self.root)
        top.pack(fill="x", **pad)
        ttk.Button(top, text="选择 PDF 文件…", command=self.pick_files).pack(side="left")
        ttk.Button(top, text="选择文件夹…", command=self.pick_folder).pack(side="left", padx=(6, 8))
        ttk.Label(top, textvariable=self.src_var, foreground="#555").pack(
            side="left", fill="x", expand=True)

        row2 = ttk.Frame(self.root)
        row2.pack(fill="x", **pad)
        ttk.Label(row2, text="输出目录:").pack(side="left")
        ttk.Entry(row2, textvariable=self.out_var).pack(
            side="left", fill="x", expand=True, padx=6)
        ttk.Button(row2, text="浏览…", command=self.pick_out).pack(side="left")
        ttk.Label(row2, text="留空 = 源目录\\word").pack(side="left", padx=(6, 0))

        row3 = ttk.Frame(self.root)
        row3.pack(fill="x", **pad)
        ttk.Label(row3, text="水印文本:").pack(side="left")
        self.wm_combo = ttk.Combobox(row3, textvariable=self.wm_var)
        self.wm_combo.pack(side="left", fill="x", expand=True, padx=6)
        self.detect_btn = ttk.Button(row3, text="检测水印", command=self.detect)
        self.detect_btn.pack(side="left")
        ttk.Checkbutton(row3, text="开始时自动检测",
                        variable=self.auto_var).pack(side="left", padx=(8, 0))

        ttk.Label(
            self.root,
            text="水印文本留空且不勾选自动检测时，仅移除 Acrobat 水印表单（无损）。\n"
                 "注意：烧在图片像素内部的水印无法去除，会原样保留。",
            foreground="#888", wraplength=800, justify="left",
        ).pack(fill="x", padx=10)

        row4 = ttk.Frame(self.root)
        row4.pack(fill="x", **pad)
        self.start_btn = ttk.Button(row4, text="开始转换", command=self.start)
        self.start_btn.pack(side="left")
        self.bar = ttk.Progressbar(row4, mode="determinate", maximum=100)
        self.bar.pack(side="left", fill="x", expand=True, padx=8)
        self.pct = ttk.Label(row4, text="0%", width=5)
        self.pct.pack(side="left")
        self.open_btn = ttk.Button(row4, text="打开输出目录",
                                   command=self.open_out, state="disabled")
        self.open_btn.pack(side="left", padx=(8, 0))

        self.log = scrolledtext.ScrolledText(self.root, height=18, state="disabled",
                                             wrap="word")
        self.log.pack(fill="both", expand=True, padx=8, pady=(0, 8))

    # ------------------------------------------------ 选择与工具
    def _update_source_label(self):
        if not self.pdfs:
            self.src_var.set("未选择")
        elif len(self.pdfs) == 1:
            self.src_var.set(f"已选 1 个文件：{self.pdfs[0].name}")
        else:
            self.src_var.set(f"已选 {len(self.pdfs)} 个文件（目录：{self.pdfs[0].parent}）")

    def _default_out(self):
        if not self.out_var.get().strip() and self.pdfs:
            self.out_var.set(str(self.pdfs[0].parent / "word"))

    def pick_files(self):
        paths = filedialog.askopenfilenames(
            title="选择 PDF 文件", filetypes=[("PDF 文件", "*.pdf")])
        if paths:
            self.pdfs = sorted(Path(p) for p in paths)
            self._update_source_label()
            self._default_out()

    def pick_folder(self):
        d = filedialog.askdirectory(title="选择包含 PDF 的文件夹")
        if d:
            self.pdfs = sorted(Path(d).glob("*.pdf"))
            self._update_source_label()
            if not self.pdfs:
                messagebox.showwarning(APP_TITLE, "该文件夹下没有 PDF 文件")
            self._default_out()

    def pick_out(self):
        d = filedialog.askdirectory(title="选择输出目录")
        if d:
            self.out_var.set(d)

    def open_out(self):
        d = self.last_out or self.out_var.get().strip()
        if not d or not Path(d).exists():
            messagebox.showinfo(APP_TITLE, "输出目录不存在")
            return
        try:
            os.startfile(d)  # Windows
        except AttributeError:
            opener = "open" if sys.platform == "darwin" else "xdg-open"
            subprocess.Popen([opener, d])

    def _append_log(self, msg: str):
        self.log.configure(state="normal")
        self.log.insert("end", msg + "\n")
        self.log.see("end")
        self.log.configure(state="disabled")

    def _post_log(self, msg: str):
        self.q.put(("log", msg))

    # ------------------------------------------------ 水印检测
    def detect(self):
        if not self.pdfs:
            messagebox.showinfo(APP_TITLE, "请先选择 PDF 文件或文件夹")
            return
        self._post_log(f"正在检测水印（样本：{self.pdfs[0].name}）…")
        self.detect_btn.config(state="disabled")

        def work():
            try:
                self.q.put(("cands", find_repeated_texts(self.pdfs[0])))
            except Exception as e:  # noqa: BLE001
                self._post_log(f"检测失败: {type(e).__name__}: {e}")
                self.q.put(("cands", []))

        threading.Thread(target=work, daemon=True).start()

    # ------------------------------------------------ 开始转换
    def start(self):
        if self.worker and self.worker.is_alive():
            messagebox.showinfo(APP_TITLE, "已有转换任务在进行中")
            return
        if not self.pdfs:
            messagebox.showinfo(APP_TITLE, "请先选择 PDF 文件或文件夹")
            return
        out_dir = self.out_var.get().strip() or str(self.pdfs[0].parent / "word")
        wm = self.wm_var.get().strip() or None
        auto = self.auto_var.get()
        pdfs = list(self.pdfs)

        self.start_btn.config(state="disabled")
        self.open_btn.config(state="disabled")
        self.bar.configure(value=0)
        self.pct.configure(text="0%")

        handler = QueueLogHandler(self.q)
        handler.setFormatter(logging.Formatter("%(message)s"))
        root_logger = logging.getLogger()
        root_logger.addHandler(handler)
        prev_level = root_logger.level
        root_logger.setLevel(logging.INFO)

        def work():
            try:
                if wm is None and auto:
                    self._post_log("正在自动检测水印…")
                    cands = [c for c in find_repeated_texts(pdfs[0])
                             if len(c["text"]) >= 4]
                    if cands:
                        use = cands[0]["text"]
                        self._post_log(
                            f"自动检测到水印: {use}（{cands[0]['pages']} 页, "
                            f"{cands[0]['ratio']:.0%}）。如不正确，请在输入框修改后重试。")
                    else:
                        use = None
                        self._post_log("未检测到重复文本水印，将仅移除 Acrobat 水印表单。")
                else:
                    use = wm
                results = batch_convert(
                    pdfs, out_dir, use,
                    progress=lambda done, total, cur:
                        self.q.put(("progress", done, total, cur)))
                self.q.put(("done", out_dir, results))
            except Exception as e:  # noqa: BLE001
                self._post_log(f"发生未预期的错误: {type(e).__name__}: {e}")
                self.q.put(("done", out_dir, []))
            finally:
                root_logger.removeHandler(handler)
                root_logger.setLevel(prev_level)

        self.worker = threading.Thread(target=work, daemon=True)
        self.worker.start()

    # ------------------------------------------------ 队列消费（主线程）
    def _drain(self):
        try:
            while True:
                item = self.q.get_nowait()
                kind = item[0]
                if kind == "log":
                    self._append_log(item[1])
                elif kind == "progress":
                    done, total, cur = item[1], item[2], item[3]
                    if isinstance(cur, dict):
                        flag = "OK  " if cur.get("ok") else "FAIL"
                        msg = f"[{done}/{total}] {flag} {cur['file']}"
                        if cur.get("ok"):
                            msg += (f"  页数={cur.get('pages')} 去除水印={cur.get('removed')}"
                                    f" 残留={'否' if cur.get('watermark_gone') else '是'}"
                                    f" 耗时={cur.get('seconds')}s")
                        else:
                            msg += f"  -> {cur.get('error', '')}"
                        self._append_log(msg)
                    else:
                        self._append_log(f"[{done + 1}/{total}] 正在转换: {cur} …")
                    pct = int(done * 100 / total) if total else 100
                    self.bar.configure(value=pct)
                    self.pct.configure(text=f"{pct}%")
                elif kind == "cands":
                    self.detect_btn.config(state="normal")
                    cands = item[1]
                    if cands:
                        texts = [c["text"] for c in cands]
                        self.wm_combo["values"] = texts
                        if not self.wm_var.get().strip():
                            self.wm_var.set(texts[0])
                        for c in cands:
                            self._append_log(
                                f"  候选水印: {c['text']}  出现 {c['pages']} 页 ({c['ratio']:.0%})")
                        self._append_log("已填入最可能的候选，可下拉切换或手动修改。")
                    else:
                        self._append_log("未发现跨页重复的文本水印候选。")
                elif kind == "done":
                    out_dir, results = item[1], item[2]
                    self.start_btn.config(state="normal")
                    self.last_out = out_dir
                    self.open_btn.config(state="normal")
                    ok = sum(1 for r in results if r.get("ok"))
                    self._append_log(f"===== 完成: {ok}/{len(results)} 成功，输出目录: {out_dir} =====")
                    if results:
                        fails = [r["file"] for r in results if not r.get("ok")]
                        if fails:
                            messagebox.showwarning(
                                APP_TITLE,
                                f"{ok}/{len(results)} 成功。\n失败文件:\n" + "\n".join(fails))
                        else:
                            messagebox.showinfo(
                                APP_TITLE, f"全部完成：{ok}/{len(results)} 成功。\n输出目录: {out_dir}")
        except queue.Empty:
            pass
        self.root.after(100, self._drain)


def main():
    root = tk.Tk()
    App(root)
    root.mainloop()


if __name__ == "__main__":
    main()
