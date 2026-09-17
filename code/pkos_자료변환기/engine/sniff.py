# -*- coding: utf-8 -*-
"""
형식 판별기 — 파일 이름이 아니라 '내용'으로 무엇인지 알아낸다
==============================================================

흩어진 자료는 이름이 믿을 게 못 된다.

    문서1.pdf          ← 실제로는 워드 파일
    스캔.jpg           ← 실제로는 PNG
    자료                ← 확장자가 아예 없음
    보고서.hwp         ← 실제로는 한글이 아니라 hwpx(ZIP)

그래서 이 모듈은 **이진 형식은 확장자를 믿지 않는다.** 앞부분 바이트(매직 넘버)를
보고, ZIP·OLE 같은 '껍데기' 형식은 안까지 열어서 진짜 정체를 확인한다.

단, **글자파일끼리는 확장자를 존중한다.** 글자파일은 매직 넘버가 없어서 내용만으로
가르면 억지가 된다 — 쉼표가 몇 개 든 메모(.txt)가 표(CSV)로 둔갑되는 식이다.
그래서 HTML·XML·JSON·자막처럼 **우연히 나올 수 없는 구조**만 확장자를 이기고,
나머지는 확장자가 없거나 모를 때만 내용으로 짐작한다.

    from engine.sniff import sniff
    k = sniff("문서1.pdf")
    k.fmt        # 'docx'
    k.label      # '워드'
    k.mismatch   # True  ← 이름은 pdf인데 속은 워드
    k.why        # '내용(ZIP 안에 word/document.xml)'

PKOS(개인지식운영체계) 프로젝트
"""

from __future__ import annotations

import io
import os
import re
import json
import struct
import zipfile
from dataclasses import dataclass

from .fsutil import long_path


# ─────────────────────────────────────────────────────────────
# 형식 이름표
# ─────────────────────────────────────────────────────────────
# fmt(내부 키) → (한글 이름, 갈래)
LABELS = {
    "hwp":    ("한글", "문서"),
    "hwpx":   ("한글", "문서"),
    "hwp3":   ("한글(97 이전)", "문서"),
    "hml":    ("한글(HML)", "문서"),
    "docx":   ("워드", "문서"),
    "doc":    ("워드(옛형식)", "문서"),
    "rtf":    ("서식글", "문서"),
    "odt":    ("오픈도큐먼트 문서", "문서"),
    "pdf":    ("PDF", "문서"),
    "epub":   ("전자책", "문서"),
    "encrypted_office": ("암호 걸린 오피스", "문서"),

    "pptx":   ("파워포인트", "발표"),
    "ppt":    ("파워포인트(옛형식)", "발표"),
    "odp":    ("오픈도큐먼트 발표", "발표"),

    "xlsx":   ("엑셀", "표"),
    "xls":    ("엑셀(옛형식)", "표"),
    "xlsb":   ("엑셀(이진)", "표"),
    "ods":    ("오픈도큐먼트 표", "표"),
    "csv":    ("표 데이터", "표"),
    "tsv":    ("표 데이터", "표"),

    "md":     ("마크다운", "글"),
    "txt":    ("텍스트", "글"),
    "html":   ("웹문서", "글"),
    "xml":    ("XML", "글"),
    "json":   ("JSON", "글"),
    "srt":    ("자막", "글"),
    "code":   ("소스코드", "글"),

    "jpeg":   ("사진", "그림"),
    "png":    ("그림", "그림"),
    "gif":    ("움직이는 그림", "그림"),
    "webp":   ("그림", "그림"),
    "bmp":    ("그림", "그림"),
    "tiff":   ("그림", "그림"),
    "heic":   ("아이폰 사진", "그림"),
    "svg":    ("벡터 그림", "그림"),
    "ico":    ("아이콘", "그림"),

    "zip":    ("압축", "압축"),
    "gzip":   ("압축", "압축"),
    "bzip2":  ("압축", "압축"),
    "xz":     ("압축", "압축"),
    "tar":    ("압축", "압축"),
    "7z":     ("압축(7z)", "압축"),
    "rar":    ("압축(RAR)", "압축"),

    "gdoc":   ("구글문서", "문서"),
    "gsheet": ("구글시트", "표"),
    "gslides": ("구글슬라이드", "발표"),

    "mp3":    ("소리", "소리영상"),
    "wav":    ("소리", "소리영상"),
    "mp4":    ("영상", "소리영상"),
    "ogg":    ("소리", "소리영상"),
    "flac":   ("소리", "소리영상"),

    "lnk":    ("바로가기", "기타"),
    "sqlite": ("데이터베이스", "기타"),
    "exe":    ("실행파일", "기타"),
    "empty":  ("빈 파일", "기타"),
    "unknown": ("알 수 없음", "기타"),
}

# 확장자 → 기대되는 fmt. 이진 형식에서는 **대조용**이고, 글자파일에서는 판단에 쓴다.
EXT_HINT = {
    ".hwp": "hwp", ".hwt": "hwp", ".hwpx": "hwpx", ".hwtx": "hwpx", ".hml": "hml",
    ".docx": "docx", ".docm": "docx", ".dotx": "docx", ".dotm": "docx",
    ".doc": "doc", ".dot": "doc", ".rtf": "rtf", ".odt": "odt",
    ".pdf": "pdf", ".epub": "epub",
    ".pptx": "pptx", ".pptm": "pptx", ".ppsx": "pptx", ".ppsm": "pptx",
    ".potx": "pptx", ".potm": "pptx", ".ppt": "ppt", ".pps": "ppt", ".pot": "ppt",
    ".odp": "odp",
    ".xlsx": "xlsx", ".xlsm": "xlsx", ".xltx": "xlsx", ".xltm": "xlsx",
    ".xls": "xls", ".xlt": "xls", ".xlsb": "xlsb", ".ods": "ods",
    ".csv": "csv", ".tsv": "tsv", ".tab": "tsv",
    ".md": "md", ".markdown": "md", ".txt": "txt", ".text": "txt", ".log": "txt",
    ".html": "html", ".htm": "html", ".xhtml": "html", ".mht": "html",
    ".xml": "xml", ".json": "json", ".srt": "srt", ".vtt": "srt", ".smi": "srt",
    ".jpg": "jpeg", ".jpeg": "jpeg", ".jpe": "jpeg",
    ".png": "png", ".gif": "gif", ".webp": "webp", ".bmp": "bmp",
    ".tif": "tiff", ".tiff": "tiff", ".heic": "heic", ".heif": "heic",
    ".svg": "svg", ".ico": "ico",
    ".zip": "zip", ".gz": "gzip", ".bz2": "bzip2", ".xz": "xz",
    ".tar": "tar", ".7z": "7z", ".rar": "rar",
    ".gdoc": "gdoc", ".gsheet": "gsheet", ".gslides": "gslides",
    ".mp3": "mp3", ".wav": "wav", ".mp4": "mp4", ".m4a": "mp3",
    ".ogg": "ogg", ".flac": "flac",
    ".py": "code", ".js": "code", ".ts": "code", ".java": "code",
    ".c": "code", ".cpp": "code", ".h": "code", ".cs": "code",
    ".r": "code", ".sql": "code", ".sh": "code", ".ipynb": "json",
    ".css": "code", ".yml": "code", ".yaml": "code", ".toml": "code",
    ".ini": "txt", ".cfg": "txt", ".bat": "code", ".cmd": "code", ".ps1": "code",
    ".lnk": "lnk", ".db": "sqlite", ".sqlite": "sqlite", ".exe": "exe",
}

TEXT_FMTS = {"txt", "md", "csv", "tsv", "code", "srt", "json", "xml", "html", "svg", "hml"}

# 확장자가 달라도 같은 것으로 보는 묶음 (mismatch 로 치지 않는다)
_SAME = [
    TEXT_FMTS,
    {"jpeg", "heic"},          # 아이폰 사진은 변환돼 저장되는 경우가 잦다
    {"zip", "gzip", "tar"},
    {"doc", "rtf"},            # 워드가 .doc 이름으로 RTF 를 저장하는 일이 흔하다
]


# ─────────────────────────────────────────────────────────────
# 변환 대상이 아닌 잡동사니
# ─────────────────────────────────────────────────────────────
_JUNK_NAMES = {"thumbs.db", "desktop.ini", ".ds_store", "ehthumbs.db",
               "ehthumbs_vista.db", "icon\r"}
_JUNK_EXT = {".lnk", ".tmp", ".temp", ".crdownload", ".part", ".partial",
             ".download", ".bak~", ".swp", ".~lock"}


def is_junk_file(name: str) -> bool:
    """운영체제·프로그램이 만든 부산물이라 자료가 아닌 파일"""
    low = name.lower()
    if low in _JUNK_NAMES:
        return True
    if low.startswith(("~$", ".~lock.", "._")) or low.startswith("."):
        return True
    return os.path.splitext(low)[1] in _JUNK_EXT


# ─────────────────────────────────────────────────────────────
@dataclass
class FileKind:
    fmt: str = "unknown"          # 내부 형식 키
    label: str = "알 수 없음"      # 사람이 읽는 이름
    group: str = "기타"            # 갈래
    why: str = ""                 # 어떻게 판별했는지
    ext: str = ""                 # 원래 확장자 (없으면 "")
    ext_fmt: str = ""             # 확장자가 주장하던 형식
    mismatch: bool = False        # 이름과 속이 다른가
    encoding: str = ""            # 글자파일일 때 알아낸 인코딩
    note: str = ""                # 덧붙일 말

    @property
    def is_text(self) -> bool:
        return self.fmt in TEXT_FMTS

    @property
    def is_image(self) -> bool:
        return self.group == "그림"

    @property
    def is_archive(self) -> bool:
        return self.group == "압축"

    def __str__(self) -> str:
        s = f"{self.label}({self.fmt})"
        if self.mismatch:
            s += f" ⚠ 이름은 {self.ext}"
        return s


def _kind(fmt: str, why: str, **kw) -> FileKind:
    label, group = LABELS.get(fmt, ("알 수 없음", "기타"))
    return FileKind(fmt=fmt, label=label, group=group, why=why, **kw)


# ─────────────────────────────────────────────────────────────
# 1단계 — 매직 넘버 (파일 앞부분 바이트)
# ─────────────────────────────────────────────────────────────
OLE_MAGIC = b"\xd0\xcf\x11\xe0\xa1\xb1\x1a\xe1"
ZIP_MAGIC = b"PK\x03\x04"
ZIP_EMPTY = b"PK\x05\x06"

_MAGIC = [
    (0, b"%PDF",                    "pdf"),
    (0, b"{\\rtf",                  "rtf"),
    (0, b"HWP Document File",       "hwp3"),
    (0, b"\x89PNG\r\n\x1a\n",       "png"),
    (0, b"\xff\xd8\xff",            "jpeg"),
    (0, b"GIF87a",                  "gif"),
    (0, b"GIF89a",                  "gif"),
    (0, b"II*\x00",                 "tiff"),
    (0, b"MM\x00*",                 "tiff"),
    (0, b"\x00\x00\x01\x00",        "ico"),
    (0, b"\x1f\x8b",                "gzip"),
    (0, b"BZh",                     "bzip2"),
    (0, b"\xfd7zXZ\x00",            "xz"),
    (0, b"7z\xbc\xaf\x27\x1c",      "7z"),
    (0, b"Rar!\x1a\x07",            "rar"),
    (0, b"SQLite format 3\x00",     "sqlite"),
    (0, b"OggS",                    "ogg"),
    (0, b"fLaC",                    "flac"),
    (0, b"ID3",                     "mp3"),
    (0, b"\xff\xfb",                "mp3"),
    (0, b"\xff\xf3",                "mp3"),
    (0, b"L\x00\x00\x00\x01\x14\x02\x00", "lnk"),
    (257, b"ustar",                 "tar"),
]

_FTYP = {
    b"heic": "heic", b"heix": "heic", b"hevc": "heic",
    b"mif1": "heic", b"msf1": "heic", b"heim": "heic",
    b"isom": "mp4", b"iso2": "mp4", b"mp41": "mp4", b"mp42": "mp4",
    b"avc1": "mp4", b"qt  ": "mp4", b"M4V ": "mp4", b"M4A ": "mp3",
}


def _bmp(head: bytes) -> bool:
    """'BM' 두 글자는 평범한 글에도 나온다. 헤더 길이까지 맞아야 BMP 다."""
    if head[:2] != b"BM" or len(head) < 18:
        return False
    size = struct.unpack_from("<I", head, 2)[0]
    hdr = struct.unpack_from("<I", head, 14)[0]
    return size >= 26 and hdr in (12, 40, 52, 56, 64, 108, 124)


def _exe(head: bytes) -> bool:
    """'MZ' 도 마찬가지. PE 헤더가 제자리에 있어야 실행파일이다."""
    if head[:2] != b"MZ" or len(head) < 0x40:
        return False
    off = struct.unpack_from("<I", head, 0x3C)[0]
    return 0x40 <= off <= len(head) - 4 and head[off:off + 2] == b"PE"


def _magic(head: bytes) -> str | None:
    for off, sig, fmt in _MAGIC:
        if len(head) >= off + len(sig) and head[off:off + len(sig)] == sig:
            return fmt
    if head[:4] == b"RIFF" and len(head) >= 12:
        if head[8:12] == b"WEBP":
            return "webp"
        if head[8:12] == b"WAVE":
            return "wav"
    if len(head) >= 12 and head[4:8] == b"ftyp":
        return _FTYP.get(head[8:12], "mp4")
    if _bmp(head):
        return "bmp"
    if _exe(head):
        return "exe"
    return None


# ─────────────────────────────────────────────────────────────
# 2단계 — 껍데기 열어보기 (ZIP / OLE)
# ─────────────────────────────────────────────────────────────
_ZIP_MARKERS = [
    ("word/document.xml",     "docx"),
    ("ppt/presentation.xml",  "pptx"),
    ("xl/workbook.xml",       "xlsx"),
    ("xl/workbook.bin",       "xlsb"),
    ("Contents/content.hpf",  "hwpx"),
    ("Contents/section0.xml", "hwpx"),
]

_MIMETYPE = {
    "application/epub+zip": "epub",
    "application/hwp+zip": "hwpx",
    "application/vnd.oasis.opendocument.text": "odt",
    "application/vnd.oasis.opendocument.spreadsheet": "ods",
    "application/vnd.oasis.opendocument.presentation": "odp",
}


def _sniff_zip(src) -> tuple[str, str]:
    """ZIP 껍데기를 열어 진짜 형식을 본다. src 는 경로 또는 바이트."""
    try:
        zf = zipfile.ZipFile(long_path(src) if isinstance(src, str) else io.BytesIO(src))
    except Exception:
        return "zip", "내용(ZIP 표지는 있으나 목록을 읽지 못함)"
    try:
        names = set(zf.namelist())
        if "mimetype" in names:
            try:
                mt = zf.read("mimetype").decode("ascii", "ignore").strip()
                if mt in _MIMETYPE:
                    return _MIMETYPE[mt], f"내용(ZIP 안 mimetype={mt})"
            except Exception:
                pass
        for marker, fmt in _ZIP_MARKERS:
            if marker in names:
                return fmt, f"내용(ZIP 안에 {marker})"
        if any(n.startswith("Contents/") for n in names) and \
           any("version.xml" in n for n in names):
            return "hwpx", "내용(ZIP 안 Contents/ + version.xml)"
        if "[Content_Types].xml" in names:
            for pre, fmt in (("word/", "docx"), ("ppt/", "pptx"), ("xl/", "xlsx")):
                if any(n.startswith(pre) for n in names):
                    return fmt, f"내용(ZIP 안 {pre} 폴더)"
        return "zip", "내용(그냥 ZIP)"
    finally:
        zf.close()


# OLE 저장소 이름 → fmt. 위에서부터 먼저 맞는 것.
_OLE_MARKERS = [
    ("EncryptedPackage",      "encrypted_office"),   # 암호 건 docx/xlsx/pptx
    ("FileHeader",            "hwp"),
    ("HwpSummaryInformation", "hwp"),
    ("WordDocument",          "doc"),
    ("PowerPoint Document",   "ppt"),
    ("Workbook",              "xls"),
    ("Book",                  "xls"),
]


def _sniff_ole(path: str | None, data: bytes) -> tuple[str, str]:
    try:
        import olefile
    except ImportError:
        for name, fmt in _OLE_MARKERS:
            if name.encode("utf-16-le") in data:
                return fmt, f"내용(OLE 안에 {name} · 간이판별)"
        return "unknown", "내용(OLE 인데 olefile 이 없어 확인 불가)"
    try:
        ole = olefile.OleFileIO(long_path(path) if path else io.BytesIO(data))
        try:
            entries = {e[-1].lstrip("\x05\x06") for e in
                       ole.listdir(streams=True, storages=True)}
            entries |= {e[0] for e in ole.listdir(streams=True, storages=True)}
        finally:
            ole.close()
    except Exception as e:
        return "unknown", f"내용(OLE 여는 중 오류: {e})"
    for name, fmt in _OLE_MARKERS:
        if name in entries:
            return fmt, f"내용(OLE 안에 {name})"
    return "unknown", "내용(OLE 인데 아는 형식이 아님)"


# ─────────────────────────────────────────────────────────────
# 3단계 — 글자파일인지, 무슨 글자파일인지
# ─────────────────────────────────────────────────────────────
_BOM = [
    (b"\xef\xbb\xbf", "utf-8-sig"),
    (b"\xff\xfe\x00\x00", "utf-32-le"),
    (b"\x00\x00\xfe\xff", "utf-32-be"),
    (b"\xff\xfe", "utf-16-le"),
    (b"\xfe\xff", "utf-16-be"),
]


def _mostly_readable(s: str) -> bool:
    if not s:
        return False
    good = sum(1 for ch in s if ch.isprintable() or ch in "\r\n\t")
    return good / len(s) >= 0.95


def detect_encoding(data: bytes) -> str | None:
    """글자파일이면 인코딩 이름을, 아니면 None 을 돌려준다."""
    for bom, enc in _BOM:
        if data.startswith(bom):
            return enc

    sample = data[:8192]
    if b"\x00" in sample:
        # BOM 없는 UTF-16 — 글자마다 NUL 이 끼어 있다
        even = sample[:len(sample) // 2 * 2]
        for enc in ("utf-16-le", "utf-16-be"):
            try:
                if _mostly_readable(even.decode(enc)):
                    return enc
            except UnicodeDecodeError:
                continue
        return None

    for enc in ("utf-8", "cp949"):
        try:
            text = sample.decode(enc)
        except UnicodeDecodeError as e:
            if e.start < len(sample) - 3:          # 끝에서 잘린 것만 눈감아 준다
                continue
            text = sample[:e.start].decode(enc)
        if _mostly_readable(text):
            return enc
    return None


_SRT_TIME = re.compile(r"\d\d:\d\d:\d\d[,.]\d{3}\s*-->")
_HTML_TAG = re.compile(r"<(?:p|div|br|table|span|a|body|head|h[1-6]|ul|ol|li)\b", re.I)


def _sniff_text(text: str, ext_fmt: str) -> tuple[str, str]:
    """글자파일의 세부 종류 → (fmt, 근거)"""
    head = text.lstrip("\ufeff").lstrip()[:4000]
    low = head.lower()

    # ① 우연히 나올 수 없는 구조 — 확장자를 이긴다
    if low.startswith("<?xml"):
        if "<hwpml" in low[:2000]:
            return "hml", "내용(XML 안 <HWPML> — 한글 문서)"
        if "<svg" in low[:2000]:
            return "svg", "내용(XML 안 <svg>)"
        if "<html" in low[:2000]:
            return "html", "내용(XHTML)"
        return "xml", "내용(<?xml 로 시작)"
    if low.startswith("<!doctype html") or low.startswith("<html"):
        return "html", "내용(HTML 태그로 시작)"
    if low.startswith("<svg"):
        return "svg", "내용(<svg> 로 시작)"
    if head[:1] in ("{", "[") and ext_fmt != "code":
        try:
            json.loads(text.lstrip("\ufeff"))
            return "json", "내용(JSON 으로 해석됨)"
        except Exception:
            pass
    if _SRT_TIME.search(head[:600]):
        return "srt", "내용(자막 시간표)"

    # ② 글자파일끼리는 확장자를 존중한다
    if ext_fmt in TEXT_FMTS:
        return ext_fmt, "확장자+내용(글자파일 확인)"

    # ③ 확장자가 없거나 모를 때만 내용으로 짐작한다
    if _HTML_TAG.search(head):
        return "html", "내용(HTML 태그)"
    lines = [l for l in text.splitlines()[:30] if l.strip()]
    if len(lines) >= 2:
        for sep, fmt in (("\t", "tsv"), (",", "csv")):
            counts = [l.count(sep) for l in lines]
            if counts[0] >= 1 and len(set(counts)) == 1:
                return fmt, f"내용(모든 줄의 구분자 개수가 같음)"
    md_marks = sum(1 for pat in (r"^#{1,6}\s", r"^[-*+]\s", r"^\d+\.\s", r"^>\s",
                                 r"^```", r"\[[^\]]+\]\([^)]+\)", r"\*\*[^*]+\*\*",
                                 r"^\|.+\|$")
                   if re.search(pat, text[:4000], re.M))
    if md_marks >= 2:
        return "md", f"내용(마크다운 표시 {md_marks}종)"
    return "txt", "내용(그냥 글자)"


def _sniff_gshortcut(text: str) -> str | None:
    """구글 문서 바로가기(.gdoc 등)는 작은 JSON 이다"""
    try:
        info = json.loads(text)
    except Exception:
        return None
    if not isinstance(info, dict):
        return None
    url = str(info.get("url", ""))
    if "docs.google.com/document" in url:
        return "gdoc"
    if "docs.google.com/spreadsheets" in url:
        return "gsheet"
    if "docs.google.com/presentation" in url:
        return "gslides"
    if "doc_id" in info or "resource_id" in info:
        return "gdoc"
    return None


# ─────────────────────────────────────────────────────────────
# 본체
# ─────────────────────────────────────────────────────────────
def sniff_bytes(data: bytes, name: str = "", path: str | None = None) -> FileKind:
    """바이트로 정체를 알아낸다.

    name  파일 이름 (확장자 대조용)
    path  실제 경로. ZIP·OLE 은 목록이 파일 뒤쪽에 있어 경로로 여는 편이 정확하다.
    """
    ext = os.path.splitext(name)[1].lower() if name else ""
    ext_fmt = EXT_HINT.get(ext, "")

    if not data:
        return _kind("empty", "내용(크기 0)", ext=ext, ext_fmt=ext_fmt,
                     note="빈 파일입니다")

    head = data[:512]
    fmt, why, enc = _magic(head), "", ""
    if fmt:
        why = "내용(매직 넘버)"

    if head.startswith(ZIP_MAGIC) or head.startswith(ZIP_EMPTY):
        fmt, why = _sniff_zip(path if path else data)
    elif head.startswith(OLE_MAGIC):
        fmt, why = _sniff_ole(path, data)

    if not fmt:
        enc = detect_encoding(data) or ""
        if enc:
            text = data[:65536].decode(enc, "replace")
            g = _sniff_gshortcut(data.decode(enc, "replace")) if len(data) < 4096 else None
            if g:
                fmt, why = g, "내용(구글 문서 바로가기 JSON)"
            else:
                fmt, why = _sniff_text(text, ext_fmt)
        else:
            fmt, why = "unknown", "내용(아는 형식이 아닌 이진 파일)"

    k = _kind(fmt, why, ext=ext, ext_fmt=ext_fmt, encoding=enc)

    if ext_fmt and ext_fmt != fmt and not any(ext_fmt in g and fmt in g for g in _SAME):
        k.mismatch = True
        e_label = LABELS.get(ext_fmt, ("?", ""))[0]
        k.note = (f"이름은 {ext}({e_label})인데 속은 {k.label}입니다. "
                  f"내용을 기준으로 변환합니다.")
    elif not ext:
        k.note = "확장자가 없어 내용만으로 판별했습니다."
    return k


def sniff(path: str, read_bytes: int = 1 << 20) -> FileKind:
    """파일 하나의 정체를 알아낸다. 앞 1MB 만 읽는다."""
    name = os.path.basename(path)
    try:
        with open(long_path(path), "rb") as f:
            data = f.read(read_bytes)
    except OSError as e:
        k = _kind("unknown", f"읽을 수 없음: {e}",
                  ext=os.path.splitext(name)[1].lower())
        k.note = str(e)
        return k
    return sniff_bytes(data, name=name, path=path)


def sniff_folder(folder: str, limit: int | None = None) -> list[tuple[str, FileKind]]:
    """폴더 아래 모든 파일의 정체를 알아낸다. 변환은 하지 않는다."""
    from .fsutil import walk
    out = []
    for dp, dns, fns in walk(folder):
        dns[:] = sorted(d for d in dns if not d.startswith(".") and
                        d not in ("__pycache__", "node_modules", "_변환결과"))
        for fn in sorted(fns):
            if is_junk_file(fn):
                continue
            p = os.path.join(dp, fn)
            out.append((p, sniff(p)))
            if limit and len(out) >= limit:
                return out
    return out
