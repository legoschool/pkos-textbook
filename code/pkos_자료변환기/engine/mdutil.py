# -*- coding: utf-8 -*-
"""
마크다운 조각 만들기 — 표와 셀 값
===================================

여러 읽기 함수가 똑같이 쓰는 것을 한 곳에 둔다.

표
    · 모든 행이 비어 있는 **오른쪽 열은 버린다.** 엑셀은 쓰지 않은 열까지
      범위에 잡는 일이 흔해서, 그대로 두면 `|  |  |  |  |` 가 줄줄이 붙는다.
    · 칸 안의 | 는 \\| 로, 줄바꿈은 <br> 로 적는다. (깃허브·옵시디언 모두 읽는다)
    · 너무 긴 표는 자르되 자른 사실을 적는다.

셀 값
    · 3500000.0 → 3500000     (엑셀은 정수도 실수로 준다)
    · 2026-03-02 00:00:00 → 2026-03-02

PKOS(개인지식운영체계) 프로젝트
"""

from __future__ import annotations

import re
import datetime

MAX_ROWS = 2000


def fmt_value(v) -> str:
    """엑셀·표 칸의 값을 사람이 읽는 글로."""
    if v is None:
        return ""
    if isinstance(v, bool):
        return "TRUE" if v else "FALSE"
    if isinstance(v, datetime.datetime):
        if (v.hour, v.minute, v.second) == (0, 0, 0):
            return v.strftime("%Y-%m-%d")
        return v.strftime("%Y-%m-%d %H:%M")
    if isinstance(v, datetime.date):
        return v.isoformat()
    if isinstance(v, datetime.time):
        return v.strftime("%H:%M")
    if isinstance(v, float):
        if v != v:                                  # NaN
            return ""
        if v.is_integer() and abs(v) < 1e15:
            return str(int(v))
        return f"{v:.10g}"
    return str(v)


def _cell(v) -> str:
    s = fmt_value(v).strip()
    s = s.replace("\\", "\\\\").replace("|", "\\|")
    s = re.sub(r"\s*\r?\n\s*", "<br>", s)
    return s


def rows_to_md(rows, max_rows: int = MAX_ROWS) -> str:
    """2차원 목록을 마크다운 표로. 첫 행을 머리로 쓴다."""
    rows = [[_cell(c) for c in r] for r in rows if r is not None]
    rows = [r for r in rows if any(r)]
    if not rows:
        return ""
    width = max(len(r) for r in rows)
    rows = [r + [""] * (width - len(r)) for r in rows]
    while width > 1 and not any(r[width - 1] for r in rows):
        width -= 1
    rows = [r[:width] for r in rows]

    shown = rows[:max_rows]
    out = ["| " + " | ".join(shown[0]) + " |",
           "|" + "|".join(["---"] * width) + "|"]
    out += ["| " + " | ".join(r) + " |" for r in shown[1:]]
    if len(rows) > max_rows:
        out.append(f"\n*(전체 {len(rows):,}행 중 {max_rows:,}행만 옮김)*")
    return "\n".join(out)


def join_blocks(blocks) -> str:
    """문단 덩어리들을 빈 줄 하나로 잇는다. 빈 덩어리는 버린다."""
    out = []
    for b in blocks:
        s = (b or "").strip("\n")
        if s.strip():
            out.append(s.rstrip())
    return "\n\n".join(out)
