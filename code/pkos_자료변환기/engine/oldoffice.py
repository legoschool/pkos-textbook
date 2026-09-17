# -*- coding: utf-8 -*-
"""
옛 오피스 형식 읽기 — .doc / .ppt / .xls
=========================================

기존 도구는 이 세 형식을 거절했다. 흩어진 옛 자료를 그대로 넣어야 하므로 직접 읽는다.
모두 마이크로소프트가 공개한 명세([MS-DOC] [MS-PPT] [MS-XLS])를 따른다.

.ppt — 처음 만든 판은 파일을 앞에서부터 훑으며 글자 레코드를 모두 모았다.
실제 수업자료로 돌려 보니 엉망이었다.

    ___PPT10                         ← 프로그램 내부 태그 이름
    마스터 제목 스타일 편집            ← 슬라이드 마스터의 틀 문구
    C:\\Users\\bk\\Desktop\\…\\영상.wmv ← 남의 PC 경로 (동영상 연결)

그래서 명세대로 따라간다.

    CurrentUser → UserEditAtom → PersistDirectory (저장 위치 목록)
    → Document → SlideListWithText(슬라이드 목록, 순서가 여기 있다)
    → 각 슬라이드 본체 → 그림 영역 안의 글상자

슬라이드 순서와 제목은 목록에서, 자유 글상자 글은 본체에서, 발표자 노트는
노트 목록에서 가져온다. 마스터 목록과 CString(태그·링크·글꼴 이름)은 읽지 않는다.

.doc — 조각표(piece table)를 풀어 본문을 잇는다. 압축된 조각은 명세대로
cp1252 로 푼다(cp949 로 풀면 “ ” 같은 문장부호가 뒤 글자를 먹고 깨진다).
필드 코드(HYPERLINK "…" 따위)는 걷어내고 결과 글자만 남기며, 칸 표시(\\x07)를
따라 표를 되살린다.

.xls — xlrd 로 읽는다. 정수는 정수로, 날짜는 날짜로, 오류 칸은 오류로 적는다.

PKOS(개인지식운영체계) 프로젝트
"""

from __future__ import annotations

import re
import struct

from .readers_legacy import ReadResult
from .fsutil import long_path
from .mdutil import rows_to_md, fmt_value, join_blocks


def _need_ole():
    try:
        import olefile
        return olefile
    except ImportError:
        return None


def _u32(b: bytes, off: int) -> int:
    return struct.unpack_from("<I", b, off)[0]


# ═════════════════════════════════════════════════════════════
# 엑셀 (.xls)
# ═════════════════════════════════════════════════════════════
def read_xls(path: str) -> ReadResult:
    kind = "엑셀(옛형식)"
    try:
        import xlrd
    except ImportError:
        return ReadResult(False, kind=kind, error="xlrd 설치 필요 — pip install xlrd")
    try:
        wb = xlrd.open_workbook(long_path(path), on_demand=True)
    except Exception as e:
        msg = str(e)
        if "encrypt" in msg.lower():
            return ReadResult(False, kind=kind, error="암호가 걸린 엑셀 파일입니다. 엑셀에서 "
                                                      "암호를 풀고 저장한 뒤 다시 넣어 주세요.")
        return ReadResult(False, kind=kind, error=f"열지 못했습니다: {msg}. 엑셀에서 열어 "
                                                  f".xlsx 로 다시 저장해 보세요.")
    try:
        blocks, n_rows = [], 0
        for idx in range(wb.nsheets):
            sh = wb.sheet_by_index(idx)
            rows = []
            for r in range(sh.nrows):
                row = []
                for c in range(sh.ncols):
                    t, v = sh.cell_type(r, c), sh.cell_value(r, c)
                    if t == xlrd.XL_CELL_DATE:
                        try:
                            v = xlrd.xldate.xldate_as_datetime(v, wb.datemode)
                        except Exception:
                            pass
                    elif t == xlrd.XL_CELL_ERROR:
                        v = xlrd.error_text_from_code.get(v, "#오류")
                    elif t == xlrd.XL_CELL_BOOLEAN:
                        v = bool(v)
                    row.append(fmt_value(v))
                rows.append(row)
            n_rows += sh.nrows
            table = rows_to_md(rows)
            if table:
                hidden = " (숨긴 시트)" if getattr(sh, "visibility", 0) else ""
                blocks += [f"## {sh.name}{hidden}", table]
            wb.unload_sheet(idx)
        return ReadResult(True, join_blocks(blocks), kind,
                          {"시트": wb.nsheets, "행": n_rows})
    except Exception as e:
        return ReadResult(False, kind=kind, error=str(e))
    finally:
        try:
            wb.release_resources()
        except Exception:
            pass


# ═════════════════════════════════════════════════════════════
# 파워포인트 (.ppt) — [MS-PPT]
# ═════════════════════════════════════════════════════════════
RT_DOCUMENT = 0x03E8
RT_SLIDE = 0x03EE
RT_SLIDE_ATOM = 0x03EF
RT_NOTES = 0x03F0
RT_SLIDE_PERSIST = 0x03F3
RT_SLIDE_LIST_WITH_TEXT = 0x0FF0
RT_TEXT_HEADER = 0x0F9F
RT_TEXT_CHARS = 0x0FA0
RT_TEXT_BYTES = 0x0FA8
RT_USER_EDIT = 0x0FF5
RT_CURRENT_USER = 0x0FF6
RT_PERSIST_DIRECTORY = 0x1772

_TITLE_TYPES = {0, 6}          # Title, CenterTitle
_NOTES_TYPE = 2
_MAX_DEPTH = 16


def _records(data: bytes, start: int, end: int):
    """[start, end) 안의 레코드를 (버전, 인스턴스, 종류, 본문시작, 길이) 로."""
    pos = start
    end = min(end, len(data))
    while pos + 8 <= end:
        vi, rt, ln = struct.unpack_from("<HHI", data, pos)
        body = pos + 8
        if body + ln > end:
            return
        yield vi & 0x0F, vi >> 4, rt, body, ln
        pos = body + ln


def _header(data: bytes, pos: int):
    if pos < 0 or pos + 8 > len(data):
        return None
    vi, rt, ln = struct.unpack_from("<HHI", data, pos)
    return vi & 0x0F, vi >> 4, rt, pos + 8, ln


def _texts_in(data: bytes, start: int, end: int, depth: int = 0) -> list[tuple[int | None, str]]:
    """컨테이너 안(그림 영역의 글상자 포함)의 글을 (글종류, 글) 로 모은다."""
    out, ttype = [], None
    for ver, _inst, rt, body, ln in _records(data, start, end):
        if ver == 0x0F:
            if depth < _MAX_DEPTH:
                out += _texts_in(data, body, body + ln, depth + 1)
        elif rt == RT_TEXT_HEADER and ln >= 4:
            ttype = _u32(data, body)
        elif rt == RT_TEXT_CHARS:
            out.append((ttype, data[body:body + ln].decode("utf-16-le", "ignore")))
        elif rt == RT_TEXT_BYTES:
            out.append((ttype, data[body:body + ln].decode("latin-1", "ignore")))
    return out


def _persist_directory(stream: bytes, current_user: bytes) -> tuple[dict, int | None]:
    """저장 위치 목록(영속 ID → 스트림 안 위치)과 Document 의 영속 ID."""
    first = _header(current_user, 0)
    if not first or first[2] != RT_CURRENT_USER or first[4] < 12:
        raise ValueError("CurrentUser 레코드가 없습니다")
    off_edit = _u32(current_user, first[3] + 8)

    edits, seen = [], set()
    while off_edit and off_edit not in seen:
        seen.add(off_edit)
        h = _header(stream, off_edit)
        if not h or h[2] != RT_USER_EDIT or h[4] < 20:
            break
        body = h[3]
        edits.append((_u32(stream, body + 12), _u32(stream, body + 16)))
        off_edit = _u32(stream, body + 8)
    if not edits:
        raise ValueError("UserEditAtom 을 찾지 못했습니다")

    directory: dict[int, int] = {}
    for off_dir, _ in reversed(edits):             # 오래된 편집부터 — 새것이 덮어쓴다
        h = _header(stream, off_dir)
        if not h or h[2] != RT_PERSIST_DIRECTORY:
            continue
        p, end = h[3], h[3] + h[4]
        while p + 4 <= end:
            w = _u32(stream, p)
            p += 4
            pid, count = w & 0xFFFFF, w >> 20
            for k in range(count):
                if p + 4 > end:
                    break
                directory[pid + k] = _u32(stream, p)
                p += 4
    return directory, edits[0][1]


def _clean_ppt_text(s: str) -> list[str]:
    s = s.replace("\r", "\n").replace("\x0b", "\n")
    lines = []
    for ln in s.split("\n"):
        t = "".join(ch for ch in ln if ch >= " " or ch == "\t").strip()
        if t and t != "*":                         # '*' 는 슬라이드 번호 자리표시
            lines.append(t)
    return lines


def _slide_lists(stream: bytes, doc_body: int, doc_len: int):
    """Document 안의 슬라이드 목록(인스턴스 0)과 노트 목록(2)."""
    slides, notes = [], []
    for _v, inst, rt, body, ln in _records(stream, doc_body, doc_body + doc_len):
        if rt != RT_SLIDE_LIST_WITH_TEXT or inst not in (0, 2):
            continue                               # 1 = 마스터 목록 — 읽지 않는다
        cur, ttype = None, None
        for _v2, _i2, rt2, b2, l2 in _records(stream, body, body + ln):
            if rt2 == RT_SLIDE_PERSIST and l2 >= 16:
                cur = {"ref": _u32(stream, b2), "id": _u32(stream, b2 + 12), "texts": []}
                (slides if inst == 0 else notes).append(cur)
                ttype = None
            elif rt2 == RT_TEXT_HEADER and l2 >= 4:
                ttype = _u32(stream, b2)
            elif rt2 in (RT_TEXT_CHARS, RT_TEXT_BYTES) and cur is not None:
                enc = "utf-16-le" if rt2 == RT_TEXT_CHARS else "latin-1"
                cur["texts"].append((ttype, stream[b2:b2 + l2].decode(enc, "ignore")))
    return slides, notes


def read_ppt(path: str) -> ReadResult:
    kind = "파워포인트(옛형식)"
    olefile = _need_ole()
    if olefile is None:
        return ReadResult(False, kind=kind, error="olefile 설치 필요 — pip install olefile")
    try:
        ole = olefile.OleFileIO(long_path(path))
        try:
            if ole.exists("EncryptedSummary"):
                return ReadResult(False, kind=kind,
                                  error="암호가 걸린 파워포인트입니다. 암호를 풀고 저장한 뒤 넣어 주세요.")
            if not ole.exists("PowerPoint Document"):
                return ReadResult(False, kind=kind, error="PowerPoint Document 스트림이 없습니다")
            stream = ole.openstream("PowerPoint Document").read()
            current = ole.openstream("Current User").read() if ole.exists("Current User") else b""
        finally:
            ole.close()
    except Exception as e:
        return ReadResult(False, kind=kind, error=str(e))

    # ① 저장 위치 목록으로 Document 를 찾는다. 안 되면 앞에서부터 찾는다.
    directory, doc_ref, doc_pos = {}, None, None
    try:
        directory, doc_ref = _persist_directory(stream, current)
        doc_pos = directory.get(doc_ref)
    except Exception:
        pass
    h = _header(stream, doc_pos) if doc_pos is not None else None
    if not h or h[2] != RT_DOCUMENT:
        h = next(((v, i, rt, b, l) for v, i, rt, b, l in _records(stream, 0, len(stream))
                  if rt == RT_DOCUMENT), None)
    if not h:
        return ReadResult(False, kind=kind, error="Document 레코드를 찾지 못했습니다. "
                                                  ".pptx 로 다시 저장해 보세요.")
    slides, notes = _slide_lists(stream, h[3], h[4])

    # ② 노트: 노트 목록의 글 + 노트 본체의 글 → 노트 id 로 묶는다
    notes_by_id: dict[int, list[str]] = {}
    for nt in notes:
        texts = [t for _, t in nt["texts"]]
        pos = directory.get(nt["ref"])
        nh = _header(stream, pos) if pos is not None else None
        if nh and nh[2] == RT_NOTES:
            texts += [t for tt, t in _texts_in(stream, nh[3], nh[3] + nh[4])
                      if tt in (None, _NOTES_TYPE, 1, 4)]
        lines = []
        for t in texts:
            for ln in _clean_ppt_text(t):
                if ln not in lines and not ln.isdigit():
                    lines.append(ln)
        notes_by_id[nt["id"]] = lines

    # ③ 슬라이드마다: 목록의 글(제목·본문) + 본체 글상자의 글 + 노트
    blocks, n_notes = [], 0
    for n, sl in enumerate(slides, 1):
        texts = list(sl["texts"])
        notes_ref = None
        pos = directory.get(sl["ref"])
        sh = _header(stream, pos) if pos is not None else None
        if sh and sh[2] == RT_SLIDE:
            for _v, _i, rt, body, ln in _records(stream, sh[3], sh[3] + sh[4]):
                if rt == RT_SLIDE_ATOM and ln >= 20:
                    notes_ref = _u32(stream, body + 16)
            texts += _texts_in(stream, sh[3], sh[3] + sh[4])

        title, body_lines = "", []
        for tt, t in texts:
            lines = _clean_ppt_text(t)
            if not lines:
                continue
            if tt in _TITLE_TYPES and not title:
                title = " ".join(lines)
                continue
            for ln in lines:
                if ln not in body_lines and ln != title:
                    body_lines.append(ln)

        blocks.append(f"## {n}. {title}" if title else f"## 슬라이드 {n}")
        blocks += body_lines
        note = notes_by_id.get(notes_ref) if notes_ref else None
        if note:
            n_notes += 1
            blocks.append("> **발표자 노트**\n" + "\n".join("> " + x for x in note))

    text = join_blocks(blocks)
    if not slides or len(re.sub(r"##[^\n]*", "", text).strip()) < 5:
        return ReadResult(False, kind=kind,
                          error="슬라이드에서 글자를 찾지 못했습니다. 그림만 있는 발표자료거나 "
                                "구조가 특이한 파일입니다. .pptx 로 다시 저장해 보세요.")
    return ReadResult(True, text, kind, {"슬라이드": len(slides), "노트": n_notes})


# ═════════════════════════════════════════════════════════════
# 워드 (.doc) — [MS-DOC]
# ═════════════════════════════════════════════════════════════
_FIB_IDENT = 0xA5EC


def _doc_fib(wd: bytes) -> dict:
    """FIB 에서 필요한 값만 읽는다. 크기가 다른 FIB 도 csw/cslw 를 따라 읽는다."""
    if len(wd) < 0x40 or struct.unpack_from("<H", wd, 0)[0] != _FIB_IDENT:
        raise ValueError("Word 문서 표시(0xA5EC)가 없습니다")
    n_fib = struct.unpack_from("<H", wd, 2)[0]
    flags = struct.unpack_from("<H", wd, 0x0A)[0]
    csw = struct.unpack_from("<H", wd, 32)[0]
    lw = 32 + 2 + csw * 2
    cslw = struct.unpack_from("<H", wd, lw)[0]
    lw_start = lw + 2
    def lw_at(i):
        return _u32(wd, lw_start + i * 4) if i < cslw else 0
    fc_start = lw_start + cslw * 4 + 2

    def pair(i):
        return _u32(wd, fc_start + i * 8), _u32(wd, fc_start + i * 8 + 4)

    fc_clx, lcb_clx = pair(33)                  # 조각표
    fc_papx, lcb_papx = pair(13)                # 문단 속성 위치표 (PlcBtePapx)
    return {
        "nfib": n_fib,
        "encrypted": bool(flags & 0x0100),
        "table": "1Table" if flags & 0x0200 else "0Table",
        "ccp": [lw_at(i) for i in (3, 4, 5, 6, 7, 8, 9, 10)],   # 본문 각주 머리 매크로 메모 미주 글상자 머리글상자
        "fc_clx": fc_clx, "lcb_clx": lcb_clx,
        "fc_papx": fc_papx, "lcb_papx": lcb_papx,
        "fc_min": _u32(wd, 0x18), "fc_mac": _u32(wd, 0x1C),
    }


def _doc_pieces(table: bytes, fc_clx: int, lcb_clx: int) -> bytes | None:
    """CLX 에서 조각표(PlcPcd)를 꺼낸다."""
    clx = table[fc_clx:fc_clx + lcb_clx]
    i = 0
    while i < len(clx):
        if clx[i] == 0x01 and i + 3 <= len(clx):          # Prc — 건너뛴다
            i += 3 + struct.unpack_from("<H", clx, i + 1)[0]
        elif clx[i] == 0x02 and i + 5 <= len(clx):        # Pcdt
            lcb = _u32(clx, i + 1)
            return clx[i + 5:i + 5 + lcb]
        else:
            break
    return None


def _doc_text(wd: bytes, plc: bytes) -> tuple[str, list[tuple[int, int, int, bool]]]:
    """조각표로 글 전체를 잇는다 → (글, [(cp시작, cp끝, 파일위치, 압축여부)]).

    글자 하나 = CP 하나가 되도록 맞춘다. 문단 속성은 파일 위치(FC)로 적혀 있어서,
    CP 와 FC 를 오가려면 이 대응이 한 글자도 어긋나면 안 된다.
    (이모지 같은 대리 쌍도 두 글자로 둔다 — 합치는 것은 맨 마지막에 한다)
    """
    n = (len(plc) - 4) // 12
    if n <= 0:
        return "", []
    cps = struct.unpack_from(f"<{n + 1}I", plc, 0)
    base = (n + 1) * 4
    parts, pieces, cp = [], [], 0
    for k in range(n):
        fc = _u32(plc, base + k * 8 + 2)
        count = cps[k + 1] - cps[k]
        if count <= 0:
            continue
        compressed = bool(fc & 0x40000000)
        if compressed:                                   # 압축 조각 → cp1252
            off = (fc & 0x3FFFFFFF) // 2
            s = wd[off:off + count].decode("cp1252", "replace")
        else:
            off = fc & 0x3FFFFFFF
            raw = wd[off:off + count * 2]
            try:
                s = raw.decode("utf-16-le", "surrogatepass")
            except UnicodeDecodeError:
                s = raw.decode("utf-16-le", "replace")
        s = s[:count].ljust(count, "\x00")
        pieces.append((cp, cp + count, off, compressed))
        parts.append(s)
        cp += count
    return "".join(parts), pieces


# ── 문단 속성 (PAPX) — 이 문단이 표 안에 있는가, 행 끝인가
_SPRM_PFINTABLE = 0x2416
_SPRM_PFTTP = 0x2417
_SPRM_PITAP = 0x6649


def _operand_size(sprm: int, grp: bytes, pos: int) -> int | None:
    spra = sprm >> 13
    if spra in (0, 1):
        return 1
    if spra in (2, 4, 5):
        return 2
    if spra == 3:
        return 4
    if spra == 7:
        return 3
    if sprm in (0xD608, 0xD606):                  # sprmTDefTable(10): 앞 2바이트가 '나머지+1'
        return struct.unpack_from("<H", grp, pos)[0] + 1 if pos + 2 <= len(grp) else None
    if pos >= len(grp) or (sprm == 0xC615 and grp[pos] == 255):
        return None                               # sprmPChgTabs 특수형 — 더 읽지 않는다
    return grp[pos] + 1


def _papx_flags(grp: bytes) -> tuple[bool, bool, int]:
    in_table, ttp, itap, pos = False, False, 0, 0
    while pos + 2 <= len(grp):
        sprm = struct.unpack_from("<H", grp, pos)[0]
        pos += 2
        size = _operand_size(sprm, grp, pos)
        if size is None or pos + size > len(grp):
            break
        if sprm == _SPRM_PFINTABLE:
            in_table = grp[pos] == 1
        elif sprm == _SPRM_PFTTP:
            ttp = grp[pos] == 1
        elif sprm == _SPRM_PITAP and size == 4:
            itap = struct.unpack_from("<i", grp, pos)[0]
        pos += size
    if in_table and itap <= 0:
        itap = 1
    return (in_table or itap > 0), ttp, itap


def _papx_runs(wd: bytes, table: bytes, fc: int, lcb: int) -> list[tuple]:
    """PlcBtePapx → 문단 속성 페이지(FKP)들 → [(fc시작, fc끝, 표안, 행끝, 깊이)]"""
    runs = []
    if not lcb or fc + lcb > len(table):
        return runs
    plc = table[fc:fc + lcb]
    n = (len(plc) - 4) // 8
    for i in range(max(n, 0)):
        pn = _u32(plc, (n + 1) * 4 + i * 4) & 0x3FFFFF
        page = wd[pn * 512:(pn + 1) * 512]
        if len(page) < 512:
            continue
        crun = page[511]
        if not crun or 4 * (crun + 1) + 13 * crun > 511:
            continue
        rgfc = [_u32(page, 4 * k) for k in range(crun + 1)]
        for k in range(crun):
            boff = page[4 * (crun + 1) + 13 * k]
            flags = (False, False, 0)
            if boff:
                p = boff * 2
                cb = page[p]
                start, length = (p + 1, 2 * cb - 1) if cb else (p + 2, 2 * page[p + 1])
                flags = _papx_flags(page[start + 2:start + length])     # istd 2바이트 뒤
            runs.append((rgfc[k], rgfc[k + 1], *flags))
    runs.sort()
    return runs


class _ParaIndex:
    """글자 위치(CP) → 그 문단의 표 속성"""

    def __init__(self, runs, pieces):
        import bisect
        self._bisect = bisect.bisect_right
        self.runs, self.starts = runs, [r[0] for r in runs]
        self.pieces, self.pstarts = pieces, [p[0] for p in pieces]

    def flags(self, cp: int) -> tuple[bool, bool, int]:
        i = self._bisect(self.pstarts, cp) - 1
        if i < 0 or not self.runs:
            return False, False, 0
        cs, ce, fc0, comp = self.pieces[i]
        if cp >= ce:
            return False, False, 0
        fc = fc0 + (cp - cs) * (1 if comp else 2)
        j = self._bisect(self.starts, fc) - 1
        if j < 0 or fc >= self.runs[j][1]:
            return False, False, 0
        return self.runs[j][2], self.runs[j][3], self.runs[j][4]


# ── 글자 정리 — 글자 수를 바꾸지 않고 '안 보일 글자'를 \x00 으로 표시한다
_DOC_CHARMAP = {"\x0b": "\n", "\x0c": "\n", "\x0e": "\n", "\x1e": "-", "\xa0": " "}


def _visible(text: str) -> str:
    """필드 명령(\\x13 명령 \\x14 결과 \\x15)과 제어문자를 \\x00 으로 바꾼다.
    길이를 그대로 두어 CP 위치 계산이 어긋나지 않게 한다."""
    out, stack = list(text), []
    for i, ch in enumerate(text):
        if ch == "\x13":
            stack.append(True)
            out[i] = "\x00"
        elif ch == "\x14":
            if stack:
                stack[-1] = False
            out[i] = "\x00"
        elif ch == "\x15":
            if stack:
                stack.pop()
            out[i] = "\x00"
        elif ch in "\r\x07":
            continue                              # 문단·칸 표시는 필드 안이어도 남긴다
        elif any(stack):
            out[i] = "\x00"
        elif ch in _DOC_CHARMAP:
            out[i] = _DOC_CHARMAP[ch]
        elif ch < " " and ch not in "\t\n":
            out[i] = "\x00"
    return "".join(out)


def _fix_text(s: str) -> str:
    s = s.replace("\x00", "")
    try:
        return s.encode("utf-16-le", "surrogatepass").decode("utf-16-le", "replace")
    except Exception:
        return s


def _doc_story_to_md(text: str, cp0: int, idx: _ParaIndex | None) -> str:
    """이야기 하나(본문·각주…)를 문단과 표로. 문단 속성으로 칸·행을 가른다."""
    vis = _visible(text)
    blocks, rows, row, cell, para = [], [], [], [], []

    def close_row():
        if cell:
            row.append(_fix_text("".join(cell)).strip())
            cell.clear()
        if row:
            rows.append(row[:])
            row.clear()

    def close_table():
        close_row()
        if rows:
            blocks.append(rows_to_md(rows))
            rows.clear()

    for i, ch in enumerate(vis):
        if ch not in "\r\x07":
            para.append(ch)
            continue
        seg = "".join(para)
        para.clear()
        in_t, ttp, itap = idx.flags(cp0 + i) if idx else (False, False, 0)

        if in_t and ttp:                              # 행 끝 표시
            if seg.replace("\x00", "").strip():
                cell.append(seg)
            close_row()
        elif in_t and itap >= 2:                      # 표 속 표 — 바깥 칸의 글로 잇는다
            cell.append(seg + (" / " if ch == "\x07" else "\n"))
        elif in_t:
            cell.append(seg)
            if ch == "\x07":                          # 칸 끝
                row.append(_fix_text("".join(cell)).strip())
                cell.clear()
            else:
                cell.append("\n")                     # 칸 안 줄바꿈
        elif idx is None and ch == "\x07":            # 문단 속성을 못 읽은 판 — 칸을 문단처럼
            t = _fix_text(seg).strip()
            if t:
                blocks.append(t)
        else:                                         # 표 밖 문단
            close_table()
            t = _fix_text(seg).strip()
            if t:
                blocks.append(t)
    close_table()
    t = _fix_text("".join(para)).strip()
    if t:
        blocks.append(t)
    return join_blocks(blocks)


def read_doc(path: str) -> ReadResult:
    kind = "워드(옛형식)"
    olefile = _need_ole()
    if olefile is None:
        return ReadResult(False, kind=kind, error="olefile 설치 필요 — pip install olefile")
    try:
        ole = olefile.OleFileIO(long_path(path))
        try:
            if not ole.exists("WordDocument"):
                return ReadResult(False, kind=kind, error="WordDocument 스트림이 없습니다")
            wd = ole.openstream("WordDocument").read()
            fib = _doc_fib(wd)
            table = ole.openstream(fib["table"]).read() if ole.exists(fib["table"]) else b""
        finally:
            ole.close()
    except Exception as e:
        return ReadResult(False, kind=kind, error=str(e))

    if fib["encrypted"]:
        return ReadResult(False, kind=kind, error="암호가 걸린 워드 문서입니다. 워드에서 암호를 "
                                                  "풀고 저장한 뒤 다시 넣어 주세요.")
    full, pieces, idx = "", [], None
    if fib["nfib"] >= 0x00C1 and table and fib["lcb_clx"]:     # Word 97 이후
        try:
            plc = _doc_pieces(table, fib["fc_clx"], fib["lcb_clx"])
            if plc:
                full, pieces = _doc_text(wd, plc)
                idx = _ParaIndex(_papx_runs(wd, table, fib["fc_papx"], fib["lcb_papx"]), pieces)
        except Exception:
            full, pieces, idx = "", [], None
    if not full and 0 < fib["fc_min"] < fib["fc_mac"] <= len(wd):
        full = wd[fib["fc_min"]:fib["fc_mac"]].decode("cp949", "replace")   # Word 6/95

    if not full.replace("\x00", "").strip():
        return ReadResult(False, kind=kind, error="본문을 찾지 못했습니다. 워드에서 열어 "
                                                  ".docx 로 다시 저장한 뒤 넣어 주세요.")

    # 이야기(story) 구간: 본문 · 각주 · 머리글 · 매크로 · 메모 · 미주 · 글상자 · 머리글상자
    ccp = fib["ccp"]
    names = ["본문", "각주", None, None, "메모", "미주", "글상자", None]
    blocks, pos = [], 0
    if ccp[0] <= 0 or not pieces:
        blocks.append(_doc_story_to_md(full, 0, idx))
    else:
        for name, count in zip(names, ccp):
            seg = full[pos:pos + count]
            if name and seg.strip("\x00\r\x07 "):
                md = _doc_story_to_md(seg, pos, idx)
                if md:
                    blocks.append(md if name == "본문" else f"## {name}\n\n{md}")
            pos += count
    body = join_blocks(blocks)
    if len(body) < 5:
        return ReadResult(False, kind=kind, error="글자가 거의 없습니다. 그림 위주 문서이거나 "
                                                  "구조가 특이합니다. .docx 로 다시 저장해 보세요.")
    return ReadResult(True, body, kind, {"글자": len(body)})
