# -*- coding: utf-8 -*-
"""실행 화면. 파일 두 개를 고르고 실행하면 끝나는 한 흐름으로 만든다.

무거운 일은 작업 스레드에서 하고, 화면은 큐로만 갱신한다.
"""
from __future__ import annotations

import os
import queue
import subprocess
import threading
import time
import traceback
import tkinter as tk
from tkinter import filedialog

from core.paths import app_dir, config_dir
from core.pipeline import DEFAULT_SHEET, START_ROW, run
from core.readers import JournalReadError
from core.template import OUTPUT_DIRNAME, TemplateError, find_template
from core.writer_excel import ExcelWriteError

from . import theme as T
from .widgets import PathField, PillButton, RoundedBox, SegmentProgress, Switch

APP_TITLE = "조회처 추출"

# (로그 문장 앞머리, 진행률, 상태 문장) — 실제로 끝난 단계만 올린다.
# 시간으로 진행률을 추정하지 않는다. 단계가 끝나야 그 값이 된다.
PROGRESS_STEPS: tuple[tuple[str, int, str], ...] = (
    ("설정 버전", 5, "설정을 읽었습니다"),
    ("분개장 읽는 중", 12, "분개장을 읽는 중"),
    ("양식 '", 25, "분개장 해석 완료"),
    ("키워드 검색", 38, "키워드 검색 완료"),
    ("후보 ", 48, "후보 정리 완료"),
    ("양식 복사", 55, "양식을 복사했습니다"),
    ("기존 Control Sheet", 55, "기존 파일에 씁니다"),
    ("백업 생성", 62, "백업을 만들었습니다"),
    ("기존 상태 수집", 70, "원본 상태(수식·VBA·도형) 수집 중"),
    ("기준:", 75, "원본 상태 수집 완료"),
    ("기존 목록 정리", 78, "이전 목록 정리"),
    ("조회처 ", 84, "조회처를 썼습니다"),
    ("임시본 저장 완료", 92, "다시 열어 검증하는 중"),
    ("검증 결과", 98, "검증 통과 — 파일 교체"),
)
APP_SUBTITLE = "분개장을 고르면 조회처를 찾아 Control Sheet를 만들어 드립니다."


class AppWindow:
    def __init__(self) -> None:
        self.root = tk.Tk()
        self.root.title(APP_TITLE)
        self.root.geometry("1020x940")
        self.root.minsize(940, 780)
        self.root.configure(bg=T.CANVAS)
        self.fonts = T.Fonts()

        self.journal = tk.StringVar()
        self.outdir = tk.StringVar()
        self.sheet = tk.StringVar(value=DEFAULT_SHEET)
        self.dry = tk.BooleanVar(value=False)
        self.last_output = ""
        self.progress = tk.DoubleVar(value=0.0)
        self.percent_text = tk.StringVar(value="0% · 대기 중")
        self.status_text = tk.StringVar(value="분개장을 고르고 실행을 누르세요.")
        self.elapsed_text = tk.StringVar(value="")
        self._started_at = 0.0
        self._tick_job = None
        self.journal.trace_add("write", lambda *_: self._refresh_outdir())
        self.queue: "queue.Queue[tuple[str, object]]" = queue.Queue()
        self.running = False

        self._build_header()
        self._build_statusbar()      # 본문보다 먼저 붙여야 항상 맨 아래에 남는다
        self._build_body()
        self.root.after(120, self._pump)

    # ------------------------------------------------------------------ 머리말
    def _build_header(self) -> None:
        bar = tk.Frame(self.root, bg=T.WHITE)
        bar.pack(fill="x")
        inner = tk.Frame(bar, bg=T.WHITE)
        inner.pack(fill="x", padx=T.PAD_PAGE, pady=(20, 16))
        tk.Label(inner, text=APP_TITLE, bg=T.WHITE, fg=T.TEXT,
                 font=self.fonts.title, anchor="w").pack(fill="x")
        tk.Label(inner, text=APP_SUBTITLE, bg=T.WHITE, fg=T.TEXT_SUB,
                 font=self.fonts.subtitle, anchor="w").pack(fill="x", pady=(6, 0))
        tk.Frame(bar, bg=T.BORDER, height=1).pack(fill="x")

    # ------------------------------------------------------------ 상태 표시줄
    def _build_statusbar(self) -> None:
        """화면 맨 아래 줄. 진행률·상태·경과 시간을 한자리에서 보여 준다."""
        bar = tk.Frame(self.root, bg=T.WHITE)
        bar.pack(side="bottom", fill="x")
        tk.Frame(bar, bg=T.BORDER, height=1).pack(fill="x")

        inner = tk.Frame(bar, bg=T.WHITE)
        inner.pack(fill="x", padx=T.PAD_PAGE, pady=(10, 12))

        SegmentProgress(inner, self.progress).pack(fill="x", pady=(0, 7))

        line = tk.Frame(inner, bg=T.WHITE)
        line.pack(fill="x")
        self.dot = tk.Canvas(line, bg=T.WHITE, highlightthickness=0, bd=0,
                             width=12, height=12)
        self.dot.pack(side="left", pady=(2, 0))
        tk.Label(line, textvariable=self.percent_text, bg=T.WHITE, fg=T.TEXT,
                 font=self.fonts.body_bold).pack(side="left", padx=(6, 10))
        tk.Label(line, textvariable=self.elapsed_text, bg=T.WHITE, fg=T.TEXT_MUTED,
                 font=self.fonts.caption).pack(side="right")
        tk.Label(line, textvariable=self.status_text, bg=T.WHITE, fg=T.TEXT_SUB,
                 font=self.fonts.body, anchor="w").pack(side="left", fill="x", expand=True)
        self._set_dot(T.TEXT_MUTED)

    def _set_dot(self, color: str) -> None:
        self.dot.delete("all")
        self.dot.create_oval(2, 3, 10, 11, fill=color, outline="")

    def _progress(self, value: float, label: str, color: str = T.BLUE) -> None:
        value = max(0.0, min(100.0, value))
        if value >= self.progress.get() or value == 0:
            self.progress.set(value)
        self.percent_text.set(f"{int(self.progress.get())}% · {label}")
        self._set_dot(color)

    def _read_step(self, message: str) -> None:
        """로그 한 줄을 보고 끝난 단계를 반영한다."""
        for head, value, label in PROGRESS_STEPS:
            if message.startswith(head):
                self._progress(value, label)
                break
        self.status_text.set(message if len(message) <= 120 else message[:117] + "…")

    def _tick(self) -> None:
        if not self.running:
            return
        seconds = int(time.monotonic() - self._started_at)
        self.elapsed_text.set(f"경과 {seconds // 60:02d}:{seconds % 60:02d}")
        self._tick_job = self.root.after(1000, self._tick)

    # -------------------------------------------------------------------- 본문
    def _build_body(self) -> None:
        page = tk.Frame(self.root, bg=T.CANVAS)
        page.pack(fill="both", expand=True, padx=T.PAD_PAGE, pady=T.PAD_PAGE)

        self._card_files(page)
        self._card_options(page)
        self._action_row(page)
        # 맨 아래 한 줄을 먼저 붙여 두어야 기록 카드가 늘어나도 밀리지 않는다.
        tk.Label(page, text="분개장은 읽기만 합니다. Control Sheet는 양식을 복사해 새로 만들고, "
                            "검증을 통과한 결과만 남깁니다.",
                 bg=T.CANVAS, fg=T.TEXT_MUTED, font=self.fonts.caption,
                 anchor="w").pack(side="bottom", fill="x", pady=(T.GAP, 0))
        self._card_log(page)

    def _section(self, parent: tk.Misc, number: str, title: str) -> tk.Frame:
        head = tk.Frame(parent, bg=T.WHITE)
        head.pack(fill="x", pady=(0, 14))
        tk.Label(head, text=number, bg=T.WHITE, fg=T.BLUE,
                 font=self.fonts.body_bold).pack(side="left", padx=(0, 8))
        tk.Label(head, text=title, bg=T.WHITE, fg=T.TEXT,
                 font=self.fonts.section).pack(side="left")
        return head

    def _card_files(self, page: tk.Frame) -> None:
        card = RoundedBox(page)
        card.pack(fill="x")
        body = card.body
        self._section(body, "1", "분개장 고르기")
        PathField(body, "분개장", self.journal, self.fonts, self._pick_journal,
                  "회계 분개장 (.xlsx, .xlsm)").pack(fill="x", pady=(0, 10))
        PathField(body, "저장 위치", self.outdir, self.fonts, self._pick_outdir,
                  "분개장을 고르면 그 옆 output 폴더로 정해집니다").pack(fill="x")
        note = tk.Frame(body, bg=T.WHITE)
        note.pack(fill="x", pady=(10, 0))
        tk.Label(note, text="", bg=T.WHITE, width=13).pack(side="left")
        self.template_label = tk.Label(note, text="", bg=T.WHITE, fg=T.TEXT_MUTED,
                                       font=self.fonts.caption, anchor="w", justify="left")
        self.template_label.pack(side="left", fill="x", expand=True)
        self._show_template()

    def _show_template(self) -> None:
        try:
            path = find_template()
            self.template_label.configure(
                text="Control Sheet는 양식을 복사해 새로 만듭니다 — 쓰는 양식: "
                     + os.path.basename(path))
        except TemplateError as exc:
            self.template_label.configure(text=f"양식을 찾지 못했습니다 — {exc}", fg=T.RED)

    def _refresh_outdir(self) -> None:
        journal = (self.journal.get() or "").strip()
        if journal:
            self.outdir.set(os.path.join(os.path.dirname(os.path.abspath(journal)),
                                         OUTPUT_DIRNAME))

    def _card_options(self, page: tk.Frame) -> None:
        """설정은 거의 바꾸지 않는다. 한 줄로 줄여 진행 기록에 자리를 넘긴다."""
        card = RoundedBox(page)
        card.pack(fill="x", pady=(T.GAP, 0))
        body = card.body

        row = tk.Frame(body, bg=T.WHITE)
        row.pack(fill="x")
        tk.Label(row, text="2", bg=T.WHITE, fg=T.BLUE,
                 font=self.fonts.body_bold).pack(side="left", padx=(0, 8))
        tk.Label(row, text="설정", bg=T.WHITE, fg=T.TEXT,
                 font=self.fonts.section).pack(side="left", padx=(0, 20))

        tk.Label(row, text="저장하지 않고 검색만", bg=T.WHITE, fg=T.TEXT,
                 font=self.fonts.body).pack(side="right", padx=(10, 0))
        Switch(row, self.dry).pack(side="right")

        tk.Label(row, text="출력 시트", bg=T.WHITE, fg=T.TEXT_SUB,
                 font=self.fonts.body).pack(side="left", padx=(0, 8))
        box = RoundedBox(row, radius=T.RADIUS_FIELD, pad=9, height=38)
        box.pack(side="left", fill="x", expand=True, padx=(0, 20))
        tk.Entry(box.body, textvariable=self.sheet, bg=T.WHITE, fg=T.TEXT,
                 font=self.fonts.body, relief="flat", highlightthickness=0,
                 insertbackground=T.TEXT).pack(fill="both", expand=True)

    def _action_row(self, page: tk.Frame) -> None:
        row = tk.Frame(page, bg=T.CANVAS)
        row.pack(fill="x", pady=(T.GAP + 4, T.GAP))
        self.run_btn = PillButton(row, "실행", self._start, kind="primary",
                                  height=42, min_width=132, font=self.fonts.button)
        self.run_btn.pack(side="left")
        self.open_btn = PillButton(row, "만든 파일 열기", self._open_result, kind="quiet",
                                   height=42, min_width=140, font=self.fonts.button)
        self.open_btn.pack(side="left", padx=(10, 0))
        self.open_btn.configure(state="disabled")
        self.folder_btn = PillButton(row, "폴더 열기", self._open_folder, kind="quiet",
                                     height=42, min_width=110, font=self.fonts.button)
        self.folder_btn.pack(side="left", padx=(8, 0))
        self.folder_btn.configure(state="disabled")

    def _card_log(self, page: tk.Frame) -> None:
        card = RoundedBox(page, grow=True)
        # 요청 높이를 최소 높이로 쓴다. 창이 커지면 expand로 더 늘어난다.
        # 로그가 50줄 넘게 쌓이므로 처음부터 넉넉히 보여 준다.
        card.configure(height=360)
        card.pack(fill="both", expand=True)
        body = card.body
        head = self._section(body, "3", "진행 기록")
        tk.Label(head, text="무엇을 읽고, 무엇을 찾고, 무엇을 검증했는지 줄줄이 남습니다",
                 bg=T.WHITE, fg=T.TEXT_MUTED, font=self.fonts.caption).pack(side="left",
                                                                            padx=(10, 0))
        wrap = tk.Frame(body, bg=T.WHITE)
        wrap.pack(fill="both", expand=True)
        self.log = tk.Text(wrap, bg=T.WHITE, fg=T.TEXT, font=self.fonts.mono,
                           relief="flat", highlightthickness=0, wrap="none",
                           spacing1=1, spacing3=1, padx=0, pady=0)
        self.log.pack(side="left", fill="both", expand=True)
        bar = tk.Scrollbar(wrap, command=self.log.yview, width=10,
                           bg=T.WHITE, troughcolor=T.WHITE, activebackground=T.BORDER,
                           relief="flat", borderwidth=0, highlightthickness=0)
        bar.pack(side="right", fill="y")
        self.log.configure(yscrollcommand=bar.set, state="disabled")
        self.log.tag_configure("head", foreground=T.TEXT, font=self.fonts.body_bold)
        self.log.tag_configure("ok", foreground=T.GREEN)
        self.log.tag_configure("bad", foreground=T.RED)
        self.log.tag_configure("muted", foreground=T.TEXT_MUTED)
        self.log.tag_configure("hint", foreground=T.TEXT_MUTED,
                               font=self.fonts.body)
        self._show_hint()

    def _show_hint(self) -> None:
        """실행 전 빈 화면에 무엇이 나올지 적어 둔다."""
        self.log.configure(state="normal")
        self.log.delete("1.0", "end")
        for line in ("실행하면 여기에 이렇게 남습니다.",
                     "",
                     "   · 분개장을 어떤 양식으로 읽었는지, 거래행 몇 건을 읽고 몇 건을 뺐는지",
                     "   · 키워드·기관사전으로 찾은 조회처 목록",
                     "   · A열에서 뺀 후보와 그 근거(4대보험·PG 등)",
                     "   · 만든 Control Sheet 경로와 백업 위치",
                     "   · 저장 뒤 검증 14항목 — 통과 / 실패 / 미검증",
                     "",
                     "실행 중에는 맨 아래 상태 표시줄에서 진행률과 경과 시간을 봅니다."):
            self.log.insert("end", line + "\n", "hint")
        self.log.configure(state="disabled")

    # -------------------------------------------------------------------- 동작
    def _write(self, text: str, tag: str = "") -> None:
        self.log.configure(state="normal")
        self.log.insert("end", text + "\n", tag or ())
        self.log.see("end")
        self.log.configure(state="disabled")

    def _clear_log(self) -> None:
        self.log.configure(state="normal")
        self.log.delete("1.0", "end")
        self.log.configure(state="disabled")

    def _pick_journal(self) -> None:
        p = filedialog.askopenfilename(
            title="분개장 선택",
            filetypes=[("Excel 통합 문서", "*.xlsx *.xlsm"), ("모든 파일", "*.*")])
        if p:
            self.journal.set(p)

    def _pick_outdir(self) -> None:
        p = filedialog.askdirectory(title="Control Sheet를 저장할 폴더")
        if p:
            self.outdir.set(p)

    def _open_folder(self) -> None:
        folder = (self.outdir.get() or "").strip() or app_dir()
        if self.last_output and os.path.exists(self.last_output):
            subprocess.Popen(["explorer", "/select,", os.path.normpath(self.last_output)])
        elif os.path.isdir(folder):
            subprocess.Popen(["explorer", os.path.normpath(folder)])

    def _open_result(self) -> None:
        if self.last_output and os.path.exists(self.last_output):
            os.startfile(self.last_output)

    def _start(self) -> None:
        if self.running:
            return
        journal = (self.journal.get() or "").strip()
        outdir = (self.outdir.get() or "").strip()
        if not journal or not os.path.exists(journal):
            self._progress(0, "대기 중", T.RED)
            self.status_text.set("분개장 파일을 먼저 고르세요.")
            return

        self.last_output = ""
        self._clear_log()
        self.running = True
        self.run_btn.configure(state="disabled", text="실행 중")
        self._started_at = time.monotonic()
        self.elapsed_text.set("경과 00:00")
        self._progress(0, "시작")
        self.status_text.set("시작합니다…")
        self._tick()
        threading.Thread(target=self._worker,
                         args=(journal, outdir, self.sheet.get().strip(), self.dry.get()),
                         daemon=True).start()

    def _worker(self, journal: str, outdir: str, sheet: str, dry: bool) -> None:
        try:
            res = run(journal, None, config_dir(), sheet_name=sheet,
                      start_row=START_ROW, save=not dry,
                      log=lambda m: self.queue.put(("log", m)),
                      out_dir=outdir or None)
            self.queue.put(("done", res))
        except JournalReadError as exc:
            self.queue.put(("error", f"분개장을 읽지 못했습니다 — {exc}"))
        except TemplateError as exc:
            self.queue.put(("error", f"양식 문제로 파일을 만들지 못했습니다 — {exc}"))
        except ExcelWriteError as exc:
            head = ("Excel과의 연결이 끊겨 저장하지 못했습니다"
                    if getattr(exc, "transient", False) else "저장하지 않았습니다(원본 유지)")
            self.queue.put(("error", f"{head} — {exc}"))
        except Exception as exc:  # noqa: BLE001
            self.queue.put(("error", f"{exc}\n{traceback.format_exc()}"))

    def _pump(self) -> None:
        try:
            while True:
                kind, payload = self.queue.get_nowait()
                if kind == "log":
                    self._write(str(payload), "muted")
                    self._read_step(str(payload))
                elif kind == "error":
                    self._progress(self.progress.get(), "작업 실패", T.RED)
                    self.status_text.set(str(payload).splitlines()[0][:120])
                    self._finish(False)
                    self._write("")
                    self._write("실패", "bad")
                    self._write(str(payload), "bad")
                elif kind == "done":
                    self._progress(100, "완료", T.GREEN)
                    self._finish(True)
                    self._report(payload)
        except queue.Empty:
            pass
        self.root.after(120, self._pump)

    def _finish(self, ok: bool) -> None:
        self.running = False
        if self._tick_job is not None:
            try:
                self.root.after_cancel(self._tick_job)
            except Exception:  # noqa: BLE001
                pass
            self._tick_job = None
        seconds = int(time.monotonic() - self._started_at) if self._started_at else 0
        self.elapsed_text.set(f"처리 시간 {seconds // 60:02d}:{seconds % 60:02d}")
        self.run_btn.configure(state="normal", text="실행")
        self.folder_btn.configure(state="normal" if ok else "disabled")
        if not ok:
            self.open_btn.configure(state="disabled")

    def _report(self, res) -> None:
        self.last_output = getattr(res, "output_path", "") or ""
        made = os.path.basename(self.last_output) if self.last_output else "저장 안 함"
        self.status_text.set(f"조회처 {len(res.names)}건 · {made}")
        if self.last_output:
            self.open_btn.configure(state="normal")
        self._write("")
        self._write(f"조회처 {len(res.names)}건", "head")
        for i, name in enumerate(res.names, 1):
            self._write(f"{i:3d}.  {name}")
        if res.report:
            rep = res.report
            self._write("")
            self._write(f"검증 {rep.summary()}", "head")
            for name, state, note in rep.checks:
                tag = {"PASS": "ok", "FAIL": "bad"}.get(state, "muted")
                self._write(f"  [{state:^10}] {name} — {note}", tag)
            self._write("")
            self._write(f"백업: {rep.backup}", "muted")
        self._write("")
        if self.last_output:
            self._write(f"만든 파일: {self.last_output}", "head")
        self._write(res.message, "head")

    def run(self) -> int:
        self.root.mainloop()
        return 0


def launch() -> int:
    return AppWindow().run()
