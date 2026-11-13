# -*- coding: utf-8 -*-
"""워드 → 한글 변환기 (창 화면)"""
import os
import queue
import sys
import threading
import traceback
import tkinter as tk
from tkinter import filedialog, messagebox, ttk

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, HERE)

try:
    from tkinterdnd2 import DND_FILES, TkinterDnD
    BaseTk = TkinterDnD.Tk
    HAS_DND = True
except Exception:
    BaseTk = tk.Tk
    HAS_DND = False

import word2hwp_core as core

WORD_EXT = (".docx", ".docm", ".dotx", ".dotm", ".doc")
BG = "#F5F7FB"
ACCENT = "#2B6CE6"
FONT = "맑은 고딕"


class App(BaseTk):
    def __init__(self, initial_files=None):
        super().__init__()
        self.title("워드 → 한글 변환기")
        self.geometry("640x560")
        self.minsize(560, 480)
        self.configure(bg=BG)
        self.q = queue.Queue()
        self.busy = False
        self.last_out_dir = None

        self.fmt = tk.StringVar(value="hwp")
        self.where = tk.StringVar(value="same")
        self.out_dir = tk.StringVar(value="")

        self._build()
        self.after(100, self._poll)
        if initial_files:
            self.after(300, lambda: self.start(initial_files))

    # ------------------------------------------------------------------ UI
    def _build(self):
        style = ttk.Style(self)
        try:
            style.theme_use("vista")
        except tk.TclError:
            pass
        style.configure("TLabel", background=BG, font=(FONT, 10))
        style.configure("TRadiobutton", background=BG, font=(FONT, 10))
        style.configure("Big.TButton", font=(FONT, 11, "bold"), padding=8)
        style.configure("TButton", font=(FONT, 10), padding=5)
        style.configure("Treeview", font=(FONT, 10), rowheight=26)
        style.configure("Treeview.Heading", font=(FONT, 10, "bold"))

        tk.Label(self, text="워드 → 한글 변환기", bg=BG, fg="#1A2B4C",
                 font=(FONT, 17, "bold")).pack(pady=(16, 2))
        tk.Label(self, text="Word 파일(.docx)을 한글 파일로 바꿔 줍니다. 원본은 그대로 둡니다.",
                 bg=BG, fg="#5A6478", font=(FONT, 10)).pack()

        # 드롭 영역
        self.drop = tk.Label(
            self, bg="white", fg=ACCENT, cursor="hand2", relief="groove", bd=2,
            font=(FONT, 13, "bold"),
            text=("여기에 Word 파일을 끌어다 놓으세요\n\n(또는 여기를 눌러 파일 고르기)"
                  if HAS_DND else "여기를 눌러 Word 파일을 고르세요"),
            height=6)
        self.drop.pack(fill="x", padx=20, pady=14)
        self.drop.bind("<Button-1>", lambda e: self.pick_files())
        if HAS_DND:
            self.drop.drop_target_register(DND_FILES)
            self.drop.dnd_bind("<<Drop>>", self._on_drop)
            self.drop.dnd_bind("<<DropEnter>>", lambda e: self.drop.configure(bg="#E8F0FE"))
            self.drop.dnd_bind("<<DropLeave>>", lambda e: self.drop.configure(bg="white"))

        btns = tk.Frame(self, bg=BG)
        btns.pack(fill="x", padx=20)
        ttk.Button(btns, text="📄 파일 고르기", style="Big.TButton",
                   command=self.pick_files).pack(side="left", expand=True, fill="x", padx=(0, 6))
        ttk.Button(btns, text="📁 폴더 통째로 변환", style="Big.TButton",
                   command=self.pick_folder).pack(side="left", expand=True, fill="x", padx=(6, 0))

        # 옵션
        opt = tk.Frame(self, bg=BG)
        opt.pack(fill="x", padx=22, pady=(12, 4))
        ttk.Label(opt, text="저장 형식:").grid(row=0, column=0, sticky="w")
        ttk.Radiobutton(opt, text=".hwp (기본 한글 파일)", value="hwp", variable=self.fmt).grid(row=0, column=1, sticky="w", padx=6)
        ttk.Radiobutton(opt, text=".hwpx (새 한글 형식)", value="hwpx", variable=self.fmt).grid(row=0, column=2, sticky="w", padx=6)
        ttk.Label(opt, text="저장 위치:").grid(row=1, column=0, sticky="w", pady=(6, 0))
        ttk.Radiobutton(opt, text="원본과 같은 폴더", value="same", variable=self.where).grid(row=1, column=1, sticky="w", padx=6, pady=(6, 0))
        ttk.Radiobutton(opt, text="다른 폴더 지정…", value="other", variable=self.where,
                        command=self.pick_out_dir).grid(row=1, column=2, sticky="w", padx=6, pady=(6, 0))
        self.out_label = ttk.Label(opt, text="", foreground="#5A6478")
        self.out_label.grid(row=2, column=1, columnspan=2, sticky="w", padx=6)

        # 결과 목록
        frame = tk.Frame(self, bg=BG)
        frame.pack(fill="both", expand=True, padx=20, pady=(6, 4))
        self.tree = ttk.Treeview(frame, columns=("file", "status"), show="headings", height=6)
        self.tree.heading("file", text="파일")
        self.tree.heading("status", text="상태")
        self.tree.column("file", width=360)
        self.tree.column("status", width=180)
        self.tree.tag_configure("ok", foreground="#1E8E3E")
        self.tree.tag_configure("err", foreground="#D93025")
        self.tree.tag_configure("warn", foreground="#B06000")
        sb = ttk.Scrollbar(frame, orient="vertical", command=self.tree.yview)
        self.tree.configure(yscrollcommand=sb.set)
        self.tree.pack(side="left", fill="both", expand=True)
        sb.pack(side="right", fill="y")
        self.tree.bind("<Double-1>", self._open_selected)

        self.progress = ttk.Progressbar(self, mode="determinate")
        self.progress.pack(fill="x", padx=20, pady=(4, 2))
        bottom = tk.Frame(self, bg=BG)
        bottom.pack(fill="x", padx=20, pady=(2, 12))
        self.status = ttk.Label(bottom, text="변환할 파일을 넣어 주세요.")
        self.status.pack(side="left")
        self.open_btn = ttk.Button(bottom, text="결과 폴더 열기", command=self.open_out, state="disabled")
        self.open_btn.pack(side="right")

        engine = "한글 프로그램으로 변환" if core.hancom_available() else "내장 변환기 사용 (한글 프로그램 없이 변환)"
        tk.Label(self, text=engine, bg=BG, fg="#8A93A6", font=(FONT, 8)).pack(side="bottom", pady=(0, 4))

    # -------------------------------------------------------------- 입력
    def _on_drop(self, event):
        self.drop.configure(bg="white")
        self.start(self.tk.splitlist(event.data))

    def pick_files(self):
        if self.busy:
            return
        files = filedialog.askopenfilenames(title="변환할 Word 파일 고르기",
                                            filetypes=[("Word 문서", "*.docx *.doc *.docm *.dotx"), ("모든 파일", "*.*")])
        if files:
            self.start(files)

    def pick_folder(self):
        if self.busy:
            return
        d = filedialog.askdirectory(title="Word 파일이 들어 있는 폴더 고르기")
        if d:
            self.start([d])

    def pick_out_dir(self):
        d = filedialog.askdirectory(title="한글 파일을 저장할 폴더")
        if d:
            self.out_dir.set(d)
            self.out_label.configure(text=d)
        elif not self.out_dir.get():
            self.where.set("same")

    def expand(self, paths):
        files = []
        for p in paths:
            p = p.strip('"')
            if os.path.isdir(p):
                for root, _, names in os.walk(p):
                    for n in sorted(names):
                        if n.lower().endswith(WORD_EXT) and not n.startswith("~$"):
                            files.append(os.path.join(root, n))
            elif os.path.isfile(p):
                files.append(p)
        return files

    # -------------------------------------------------------------- 변환
    def start(self, paths):
        if self.busy:
            messagebox.showinfo("잠시만요", "지금 변환 중입니다. 끝난 뒤 다시 넣어 주세요.")
            return
        files = self.expand(paths)
        words = [f for f in files if f.lower().endswith(WORD_EXT) and not os.path.basename(f).startswith("~$")]
        skipped = [f for f in files if f not in words]
        for f in skipped:
            msg = "Word 파일이 아님"
            self.tree.insert("", "end", values=(os.path.basename(f), msg), tags=("err",))
        if not words:
            if not skipped:
                messagebox.showinfo("알림", "Word 파일(.docx)을 찾지 못했습니다.")
            return
        out_dir = self.out_dir.get() if self.where.get() == "other" and self.out_dir.get() else None
        rows = [self.tree.insert("", "end", values=(os.path.basename(f), "대기 중"), tags=()) for f in words]
        self.busy = True
        self.progress.configure(maximum=len(words), value=0)
        threading.Thread(target=self._worker, args=(words, rows, out_dir, self.fmt.get()), daemon=True).start()

    def _worker(self, files, rows, out_dir, fmt):
        ok = 0
        for i, (f, row) in enumerate(zip(files, rows)):
            self.q.put(("row", row, "변환 중…", ""))
            self.q.put(("status", f"변환 중 ({i + 1}/{len(files)}): {os.path.basename(f)}"))
            try:
                def prog(x, i=i):
                    self.q.put(("progress", i + x))
                out, warns = core.convert_file(f, out_dir=out_dir, fmt=fmt, progress=prog)
                ok += 1
                notes = [w for w in warns if "한글 프로그램으로 변환" not in w]
                label = "완료 ✔" if not notes else f"완료 (참고 {len(notes)}건)"
                self.q.put(("row", row, label, "warn" if notes else "ok", out, notes))
                self.last_out_dir = os.path.dirname(out)
            except ValueError as e:      # 예상된 문제 (손상·암호·빈 파일 등): 이유만 보여 줌
                self.q.put(("row", row, f"실패: {e}", "err"))
            except Exception as e:
                traceback.print_exc()
                self.q.put(("row", row, f"실패(예상 못한 오류): {e}", "err"))
            self.q.put(("progress", i + 1))
        self.q.put(("done", ok, len(files)))

    def _poll(self):
        try:
            while True:
                msg = self.q.get_nowait()
                kind = msg[0]
                if kind == "row":
                    _, row, text, tag, *rest = msg
                    vals = self.tree.item(row, "values")
                    self.tree.item(row, values=(vals[0], text), tags=(tag,) if tag else ())
                    if rest:
                        self.outputs = getattr(self, "outputs", {})
                        self.outputs[row] = rest
                elif kind == "status":
                    self.status.configure(text=msg[1])
                elif kind == "progress":
                    self.progress.configure(value=msg[1])
                elif kind == "done":
                    self.busy = False
                    ok, total = msg[1], msg[2]
                    self.status.configure(text=f"끝났습니다: {total}개 중 {ok}개 성공. (목록을 두 번 누르면 파일이 열립니다)")
                    if self.last_out_dir:
                        self.open_btn.configure(state="normal")
        except queue.Empty:
            pass
        self.after(100, self._poll)

    def _open_selected(self, _event):
        sel = self.tree.selection()
        outs = getattr(self, "outputs", {})
        if sel and sel[0] in outs:
            out, notes = outs[sel[0]]
            if notes:
                messagebox.showinfo("참고 사항", "\n".join(notes[:15]))
            try:
                os.startfile(out)
            except OSError:
                os.startfile(os.path.dirname(out))

    def open_out(self):
        if self.last_out_dir:
            os.startfile(self.last_out_dir)


if __name__ == "__main__":
    App(sys.argv[1:]).mainloop()
