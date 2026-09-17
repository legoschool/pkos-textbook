# -*- coding: utf-8 -*-
"""
그림 다루기 — 사진 자체, 그리고 문서 속에 박힌 그림
=====================================================

1. **사진 파일을 자료로 만든다**
   사진에는 찍은 날짜·기기(EXIF)가 들어 있다. 그것만 꺼내도 시간 축에 꽂을 수
   있는 자료가 된다. 글자 인식(OCR)은 ocr.py 가 맡는다.

2. **문서 속 그림을 꺼낸다**
   발표자료·논문은 그림이 본문만큼 중요하다. 읽기 함수가 그림을 만나면
   ReadContext.save_image() 로 넘기고, 그 자리에 그림 링크를 적는다.
   (예전에는 그림을 전부 문서 맨 끝에 모았다 — 어느 슬라이드 그림인지 알 수 없었다)

3. **웹에서 보이는 형식으로 바꾼다**
   TIFF·BMP·WMF·EMF·JPEG2000 은 마크다운 뷰어와 브라우저에서 보이지 않는다.
   Pillow 로 PNG 로 바꾸고, 바꿀 수 없으면 버린다(보이지 않는 파일만 쌓이므로).

EXIF·크기 읽기는 바깥 라이브러리 없이 처리한다.

PKOS(개인지식운영체계) 프로젝트
"""

from __future__ import annotations

import io
import os
import re
import struct
import zipfile

from .fsutil import long_path

WEB_EXT = {"png", "jpg", "jpeg", "gif", "webp", "svg"}


# ─────────────────────────────────────────────────────────────
# 크기 읽기 — 헤더만 보고 가로·세로를 알아낸다
# ─────────────────────────────────────────────────────────────
def image_size(data: bytes) -> tuple[int, int] | None:
    """(가로, 세로). 알 수 없으면 None."""
    try:
        if data[:8] == b"\x89PNG\r\n\x1a\n":
            return struct.unpack_from(">II", data, 16)
        if data[:3] == b"GIF":
            return struct.unpack_from("<HH", data, 6)
        if data[:2] == b"BM":
            w, h = struct.unpack_from("<ii", data, 18)
            return abs(w), abs(h)
        if data[:4] == b"RIFF" and data[8:12] == b"WEBP":
            if data[12:16] == b"VP8X":
                return (int.from_bytes(data[24:27], "little") + 1,
                        int.from_bytes(data[27:30], "little") + 1)
            if data[12:16] == b"VP8 ":
                return (struct.unpack_from("<H", data, 26)[0] & 0x3FFF,
                        struct.unpack_from("<H", data, 28)[0] & 0x3FFF)
        if data[:3] == b"\xff\xd8\xff":
            i, n = 2, len(data)
            while i + 9 < n:
                if data[i] != 0xFF:
                    i += 1
                    continue
                m = data[i + 1]
                if m in (0xD8, 0xD9) or 0xD0 <= m <= 0xD7 or m == 0xFF:
                    i += 1 if m == 0xFF else 2
                    continue
                seg = struct.unpack_from(">H", data, i + 2)[0]
                if 0xC0 <= m <= 0xCF and m not in (0xC4, 0xC8, 0xCC):
                    h, w = struct.unpack_from(">HH", data, i + 5)
                    return w, h
                i += 2 + seg
    except Exception:
        pass
    return None


def sniff_image_ext(data: bytes) -> str:
    """바이트로 그림 형식을 알아낸다 (확장자가 틀린 그림이 흔하다)."""
    if data[:8] == b"\x89PNG\r\n\x1a\n":
        return "png"
    if data[:3] == b"\xff\xd8\xff":
        return "jpg"
    if data[:3] == b"GIF":
        return "gif"
    if data[:4] == b"RIFF" and data[8:12] == b"WEBP":
        return "webp"
    if data[:2] == b"BM":
        return "bmp"
    if data[:4] in (b"II*\x00", b"MM\x00*"):
        return "tiff"
    if data[:4] == b"\xd7\xcd\xc6\x9a" or data[:4] == b"\x01\x00\x09\x00":
        return "wmf"
    if data[40:44] == b" EMF":
        return "emf"
    if data[:12] == b"\x00\x00\x00\x0cjP  \r\n\x87\n" or data[:4] == b"\xff\x4f\xff\x51":
        return "jp2"
    head = data[:200].lstrip().lower()
    if head.startswith(b"<?xml") or head.startswith(b"<svg"):
        return "svg"
    return ""


def web_ready(data: bytes, ext: str = "") -> tuple[bytes | None, str]:
    """브라우저에서 보이는 형식으로. 바꿀 수 없으면 (None, '')."""
    real = sniff_image_ext(data) or (ext or "").lower().lstrip(".")
    if real == "jpeg":
        real = "jpg"
    if real in WEB_EXT:
        return data, real
    try:
        from PIL import Image
        img = Image.open(io.BytesIO(data))
        img.load()
        if img.mode not in ("RGB", "RGBA", "L", "LA"):
            img = img.convert("RGBA" if "A" in img.getbands() else "RGB")
        buf = io.BytesIO()
        img.save(buf, "PNG")
        return buf.getvalue(), "png"
    except Exception:
        return None, ""


# ─────────────────────────────────────────────────────────────
# EXIF — 사진에 박힌 찍은 날짜·기기
# ─────────────────────────────────────────────────────────────
_EXIF_TAGS = {
    0x010F: "제조사", 0x0110: "기기", 0x0112: "방향",
    0x9003: "찍은날짜", 0x9004: "만든날짜", 0x0132: "수정날짜",
    0x010E: "설명", 0x013B: "촬영자", 0x8298: "저작권",
}
_GPS_IFD = 0x8825
_EXIF_IFD = 0x8769


def _read_ifd(data: bytes, base: int, off: int, endian: str, out: dict, depth: int = 0):
    if depth > 3 or base + off + 2 > len(data):
        return
    try:
        count = struct.unpack_from(endian + "H", data, base + off)[0]
    except struct.error:
        return
    for i in range(min(count, 300)):
        p = base + off + 2 + i * 12
        if p + 12 > len(data):
            return
        tag, typ, num = struct.unpack_from(endian + "HHI", data, p)
        size = {1: 1, 2: 1, 3: 2, 4: 4, 5: 8, 7: 1, 9: 4, 10: 8}.get(typ, 1)
        total = size * num
        val_off = p + 8
        if total > 4:
            val_off = base + struct.unpack_from(endian + "I", data, p + 8)[0]
        if val_off < 0 or val_off + total > len(data):
            continue
        if tag in (_EXIF_IFD, _GPS_IFD) and typ in (4, 13):
            if tag == _EXIF_IFD:
                sub = struct.unpack_from(endian + "I", data, p + 8)[0]
                _read_ifd(data, base, sub, endian, out, depth + 1)
            continue
        if tag not in _EXIF_TAGS:
            continue
        name = _EXIF_TAGS[tag]
        if typ == 2:
            s = data[val_off:val_off + total].split(b"\x00")[0]
            try:
                v = s.decode("utf-8").strip()
            except UnicodeDecodeError:
                v = s.decode("cp949", "ignore").strip()
            if v:
                out[name] = v
        elif typ == 3:
            out[name] = struct.unpack_from(endian + "H", data, val_off)[0]


def read_exif(data: bytes) -> dict:
    """JPEG 의 APP1 구역에서 EXIF 를 꺼낸다. 없으면 빈 사전."""
    out: dict = {}
    if data[:3] != b"\xff\xd8\xff":
        return out
    i, n = 2, len(data)
    while i + 4 < n:
        if data[i] != 0xFF:
            i += 1
            continue
        marker = data[i + 1]
        if marker == 0xDA:
            break
        seg = struct.unpack_from(">H", data, i + 2)[0]
        if marker == 0xE1 and data[i + 4:i + 10] == b"Exif\x00\x00":
            tiff = i + 10
            if tiff + 8 > n:
                break
            endian = "<" if data[tiff:tiff + 2] == b"II" else ">"
            try:
                first = struct.unpack_from(endian + "I", data, tiff + 4)[0]
                _read_ifd(data, tiff, first, endian, out)
            except Exception:
                pass
            break
        i += 2 + seg
    return out


def exif_date(exif: dict) -> str:
    raw = exif.get("찍은날짜") or exif.get("만든날짜") or exif.get("수정날짜") or ""
    m = re.match(r"(\d{4})[:\-](\d{2})[:\-](\d{2})", str(raw))
    if not m or m.group(1) in ("0000", "1970"):
        return ""
    return f"{m.group(1)}-{m.group(2)}-{m.group(3)}"


def describe_image(path: str) -> dict:
    """사진에서 꺼낼 수 있는 사실들을 모은다."""
    try:
        with open(long_path(path), "rb") as f:
            data = f.read(512 * 1024)
        size = os.path.getsize(long_path(path))
    except OSError:
        return {}
    info: dict = {"용량": size}
    wh = image_size(data)
    if not wh:
        try:
            from PIL import Image
            with Image.open(long_path(path)) as im:
                wh = im.size
        except Exception:
            wh = None
    if wh:
        info["가로"], info["세로"] = wh
    ex = read_exif(data)
    if ex:
        d = exif_date(ex)
        if d:
            info["찍은날짜"] = d
        for k in ("제조사", "기기", "설명", "촬영자"):
            if ex.get(k):
                info[k] = ex[k]
    return info


# ─────────────────────────────────────────────────────────────
# 문서 속 그림 모으기 — 저장은 ReadContext 가 한다
# ─────────────────────────────────────────────────────────────
_MEDIA_DIRS = ("word/media/", "ppt/media/", "xl/media/", "BinData/",
               "Contents/BinData/", "Pictures/")


def _natural(s: str):
    return [int(t) if t.isdigit() else t.lower() for t in re.split(r"(\d+)", s)]


def zip_images(path: str, prefixes=_MEDIA_DIRS):
    """ZIP 계열 문서 안의 그림을 (이름, 바이트)로 차례대로 내놓는다."""
    try:
        zf = zipfile.ZipFile(long_path(path))
    except Exception:
        return
    try:
        names = sorted((n for n in zf.namelist()
                        if n.startswith(prefixes) and not n.endswith("/")), key=_natural)
        for n in names:
            try:
                yield os.path.basename(n), zf.read(n)
            except Exception:
                continue
    finally:
        zf.close()


def hwp_images(path: str):
    """구형 한글(.hwp) 의 BinData 저장소에서 그림을 (이름, 바이트)로 내놓는다."""
    try:
        import olefile
        import zlib
    except ImportError:
        return
    try:
        ole = olefile.OleFileIO(long_path(path))
    except Exception:
        return
    try:
        entries = sorted((e for e in ole.listdir(streams=True)
                          if e and e[0] == "BinData"), key=lambda e: _natural(e[-1]))
        for entry in entries:
            try:
                raw = ole.openstream(entry).read()
            except Exception:
                continue
            data = None
            for attempt in (lambda b: zlib.decompress(b, -15), zlib.decompress, lambda b: b):
                try:
                    data = attempt(raw)
                    if sniff_image_ext(data):
                        break
                except Exception:
                    data = None
            if data and sniff_image_ext(data):
                yield entry[-1], data
    finally:
        ole.close()
