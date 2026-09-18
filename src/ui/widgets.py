# -*- coding: utf-8 -*-
"""둥근 모서리 카드·알약 단추·토글. tkinter 기본 위젯에는 없어 Canvas로 그린다.

모든 색·글자 값은 theme에서 가져온다. 여기서 새 색을 만들지 않는다.
"""
from __future__ import annotations

import math
import tkinter as tk
from typing import Callable

from . import theme as T


def round_points(x1: float, y1: float, x2: float, y2: float,
                 r: float, seg: int = 10) -> list[float]:
    """둥근 사각형의 외곽선 좌표. 반지름이 너무 크면 알약이 된다."""
    r = max(0.0, min(r, (x2 - x1) / 2, (y2 - y1) / 2))
    pts: list[float] = []
    corners = ((x2 - r, y1 + r, -90.0), (x2 - r, y2 - r, 0.0),
               (x1 + r, y2 - r, 90.0), (x1 + r, y1 + r, 180.0))
    for cx, cy, start in corners:
        for i in range(seg + 1):
            a = math.radians(start + i * 90.0 / seg)
            pts.extend((cx + r * math.cos(a), cy + r * math.sin(a)))
    return pts


def draw_round(canvas: tk.Canvas, tag: str, x1: float, y1: float, x2: float, y2: float,
               radius: float, fill: str, outline: str = "") -> None:
    canvas.delete(tag)
    canvas.create_polygon(round_points(x1, y1, x2, y2, radius),
                          fill=fill, outline=outline or fill,
                          width=1, joinstyle="round", tags=tag)


class RoundedBox(tk.Canvas):
    """내용 프레임을 담는 둥근 상자. 높이는 내용에 맞춰 늘어난다."""

    def __init__(self, master: tk.Misc, *, radius: int = T.RADIUS_CARD,
                 pad: int = T.PAD_CARD, fill: str = T.WHITE,
                 outline: str = T.BORDER, height: int | None = None,
                 grow: bool = False) -> None:
        super().__init__(master, bg=master.cget("bg"), highlightthickness=0, bd=0)
        self._radius, self._pad = radius, pad
        self._fill, self._outline = fill, outline
        self._fixed = height
        self._grow = grow
        self.body = tk.Frame(self, bg=fill)
        self._win = self.create_window(pad, pad, anchor="nw", window=self.body)
        self._last = (0, 0)
        if height:
            self.configure(height=height)
        self.body.bind("<Configure>", lambda _e: self._sync())
        self.bind("<Configure>", lambda _e: self._sync())

    def _sync(self) -> None:
        w = self.winfo_width()
        if w <= 1:
            return
        inner_w = max(1, w - self._pad * 2)
        if self._grow:
            h = self.winfo_height()
            if h <= 1:
                return
            self.itemconfigure(self._win, width=inner_w, height=h - self._pad * 2)
            self.coords(self._win, self._pad, self._pad)
            draw_round(self, "box", 0.5, 0.5, w - 0.5, h - 0.5,
                       self._radius, self._fill, self._outline)
            self.tag_lower("box")
            return
        h = self._fixed or (self.body.winfo_reqheight() + self._pad * 2)
        if (w, h) != self._last:
            self._last = (w, h)
            self.configure(height=h)
        # height=0 은 Tk에서 '내용이 요구하는 높이'를 뜻한다.
        self.itemconfigure(self._win, width=inner_w,
                           height=(h - self._pad * 2) if self._fixed else 0)
        self.coords(self._win, self._pad, self._pad)
        draw_round(self, "box", 0.5, 0.5, w - 0.5, h - 0.5,
                   self._radius, self._fill, self._outline)
        self.tag_lower("box")


class PillButton(tk.Canvas):
    """apple.com 단추. kind: primary(파랑) / secondary(연회색) / quiet(테두리만)."""

    _STYLE = {
        "primary": (T.BLUE, T.BLUE_HOVER, T.BLUE_PRESS, T.WHITE, ""),
        "secondary": (T.FILL_SOFT, T.FILL_SOFT_HOVER, T.BORDER, T.TEXT, ""),
        "quiet": (T.WHITE, T.CANVAS, T.FILL_SOFT, T.LINK, T.BORDER),
    }

    def __init__(self, master: tk.Misc, text: str, command: Callable[[], None] | None = None,
                 *, kind: str = "primary", height: int = 38, min_width: int = 96,
                 font: tuple | None = None) -> None:
        super().__init__(master, bg=master.cget("bg"), highlightthickness=0, bd=0,
                         height=height, cursor="hand2")
        self._font = font or ("Malgun Gothic", 10, "bold")
        self._text = text
        self._kind = kind if kind in self._STYLE else "primary"
        self._command = command
        self._h = height
        self._state = "normal"
        self._hover = False
        self._press = False
        pad = 22
        w = max(min_width, self._measure(text) + pad * 2)
        self.configure(width=w)
        self.bind("<Enter>", self._on_enter)
        self.bind("<Leave>", self._on_leave)
        self.bind("<ButtonPress-1>", self._on_press)
        self.bind("<ButtonRelease-1>", self._on_release)
        self.bind("<Configure>", lambda _e: self._draw())

    def _measure(self, text: str) -> int:
        import tkinter.font as tkfont
        try:
            return tkfont.Font(font=self._font).measure(text)
        except Exception:  # noqa: BLE001
            return len(text) * 9

    # ---------------------------------------------------------------- 상태
    def configure(self, **kw):  # type: ignore[override]
        if "text" in kw:
            self._text = kw.pop("text")
        if "command" in kw:
            self._command = kw.pop("command")
        if "state" in kw:
            self._state = kw.pop("state")
            self.configure(cursor="arrow" if self._state == "disabled" else "hand2")
        res = super().configure(**kw) if kw else None
        self._draw()
        return res

    config = configure

    def _on_enter(self, _e) -> None:
        self._hover = True
        self._draw()

    def _on_leave(self, _e) -> None:
        self._hover = self._press = False
        self._draw()

    def _on_press(self, _e) -> None:
        if self._state == "disabled":
            return
        self._press = True
        self._draw()

    def _on_release(self, _e) -> None:
        was = self._press
        self._press = False
        self._draw()
        if was and self._state != "disabled" and self._command:
            self._command()

    def _draw(self) -> None:
        w, h = self.winfo_width(), self._h
        if w <= 1:
            w = int(self.cget("width"))
        base, hover, press, fg, line = self._STYLE[self._kind]
        if self._state == "disabled":
            fill, fg, line = T.DISABLED_FILL, T.DISABLED_TEXT, ""
        elif self._press:
            fill = press
        elif self._hover:
            fill = hover
        else:
            fill = base
        draw_round(self, "pill", 0.5, 0.5, w - 0.5, h - 0.5, h / 2, fill, line)
        self.delete("label")
        self.create_text(w / 2, h / 2, text=self._text, fill=fg,
                         font=self._font, tags="label")


class Switch(tk.Canvas):
    """켜고 끄는 토글. tk.BooleanVar와 묶어 쓴다."""

    W, H = 44, 26

    def __init__(self, master: tk.Misc, variable: tk.BooleanVar,
                 command: Callable[[], None] | None = None) -> None:
        super().__init__(master, bg=master.cget("bg"), highlightthickness=0, bd=0,
                         width=self.W, height=self.H, cursor="hand2")
        self.var = variable
        self._command = command
        self.bind("<Button-1>", self._toggle)
        self.var.trace_add("write", lambda *_: self._draw())
        self._draw()

    def _toggle(self, _e) -> None:
        self.var.set(not self.var.get())
        if self._command:
            self._command()

    def _draw(self) -> None:
        on = bool(self.var.get())
        draw_round(self, "track", 0.5, 0.5, self.W - 0.5, self.H - 0.5,
                   self.H / 2, T.BLUE if on else T.FILL_SOFT,
                   T.BLUE if on else T.BORDER)
        self.delete("knob")
        r = (self.H - 8) / 2
        cx = (self.W - r - 4.5) if on else (r + 4.5)
        self.create_oval(cx - r, self.H / 2 - r, cx + r, self.H / 2 + r,
                         fill=T.WHITE, outline="", tags="knob")


class StatusDot(tk.Canvas):
    """상태 점 + 글자. apple.com 재고 표시와 같은 모양."""

    def __init__(self, master: tk.Misc, font: tuple) -> None:
        super().__init__(master, bg=master.cget("bg"), highlightthickness=0, bd=0,
                         width=210, height=22)
        self._font = font
        self.set("대기 중", T.TEXT_MUTED)

    def set(self, text: str, color: str = T.TEXT_SUB) -> None:
        self.delete("all")
        self.create_oval(2, 8, 10, 16, fill=color, outline="")
        self.create_text(16, 11, text=text, anchor="w", fill=T.TEXT_SUB, font=self._font)


class PathField(tk.Frame):
    """[라벨] [경로 상자] [선택 단추] 한 줄."""

    def __init__(self, master: tk.Misc, label: str, variable: tk.StringVar,
                 fonts: T.Fonts, on_pick: Callable[[], None],
                 placeholder: str = "파일을 선택하세요") -> None:
        super().__init__(master, bg=master.cget("bg"))
        self.var = variable
        self._placeholder = placeholder
        tk.Label(self, text=label, bg=self.cget("bg"), fg=T.TEXT,
                 font=fonts.body_bold, anchor="w", width=13).pack(side="left")
        box = RoundedBox(self, radius=T.RADIUS_FIELD, pad=12, fill=T.WHITE,
                         outline=T.BORDER, height=42)
        box.pack(side="left", fill="x", expand=True, padx=(0, 10))
        self._label = tk.Label(box.body, text=placeholder, bg=T.WHITE, fg=T.TEXT_MUTED,
                               font=fonts.body, anchor="w")
        self._label.pack(fill="both", expand=True)
        PillButton(self, "선택", on_pick, kind="secondary", height=38,
                   min_width=84, font=fonts.button).pack(side="left")
        self.var.trace_add("write", lambda *_: self._refresh())
        self._refresh()

    def _refresh(self) -> None:
        path = (self.var.get() or "").strip()
        if not path:
            self._label.configure(text=self._placeholder, fg=T.TEXT_MUTED)
        else:
            self._label.configure(text=path, fg=T.TEXT)


class SegmentProgress(tk.Canvas):
    """숫자 진행률에 비례해 칸을 채운다. **시간으로 진행률을 추정하지 않는다.**

    회계감사 자동화 프로그램(`wp4000a/app/theme.py`)의 같은 이름 위젯을 옮겨 왔다.
    색만 이 화면의 색으로 바꿨다.
    """

    H = 8

    def __init__(self, parent: tk.Misc, variable: tk.DoubleVar, **kw) -> None:
        super().__init__(parent, height=self.H, highlightthickness=0, bd=0,
                         background=parent.cget("bg"), **kw)
        self.variable = variable
        self._trace = variable.trace_add("write", self._draw)
        self.bind("<Configure>", self._draw)
        self.bind("<Destroy>", self._cleanup)

    def _draw(self, *_a) -> None:
        self.delete("all")
        width = max(1, self.winfo_width())
        try:
            value = max(0.0, min(100.0, float(self.variable.get())))
        except (ValueError, tk.TclError):
            value = 0.0
        count = max(1, width // 12)
        unit = width / count
        boundary = width * value / 100
        for i in range(count):
            left = i * unit
            right = (i + 1) * unit - 3
            draw_round(self, f"e{i}", left, 0, right, self.H, 2, T.FILL_SOFT)
            if boundary > left:
                draw_round(self, f"f{i}", left, 0, min(right, boundary), self.H,
                           2, T.BLUE)

    def _cleanup(self, event) -> None:
        if event.widget is self:
            try:
                self.variable.trace_remove("write", self._trace)
            except tk.TclError:
                pass
