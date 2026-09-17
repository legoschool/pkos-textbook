# -*- coding: utf-8 -*-
"""
파일 입출력과 이름 짓기 — 한 곳에서만 처리한다
================================================

흩어진 자료를 다루다 보면 같은 문제가 곳곳에서 되풀이된다. 이 파일은 그것을
한 번에 막는 자리다. 다른 모듈은 파일을 직접 열지 말고 여기를 거친다.

1. **긴 경로 (260자)**
   윈도우는 기본 설정에서 경로가 260자를 넘으면 파일을 열지도 만들지도 못한다.
   구글 드라이브 안의 자료는 경로가 길어 실제로 걸린다. 게다가 os.walk 는
   그런 폴더를 **오류 없이 조용히 건너뛴다** — 자료가 사라진 줄도 모르게 된다.
   경로 앞에 \\\\?\\ 를 붙이면 이 제한을 넘을 수 있다.

2. **마크다운 링크의 공백·괄호**
   ![그림](../assets/수업 자료 (1)/001.png) 는 대부분의 뷰어에서 깨진다.
   공백과 괄호를 %20·%28·%29 로 바꿔 적어야 한다. 한글은 그대로 둔다.

3. **머리말(YAML) 문자열**
   제목에 \\ 나 " 가 들어가면 머리말이 깨진다. JSON 문자열 규칙으로 적으면
   YAML 도 그대로 읽는다.

PKOS(개인지식운영체계) 프로젝트
"""

from __future__ import annotations

import io
import os
import re
import json
import shutil
import hashlib

_WIN = os.name == "nt"
_LIMIT = 240            # 260 에서 파일 이름이 붙을 여유를 뺀 값


# ─────────────────────────────────────────────────────────────
# 1. 긴 경로
# ─────────────────────────────────────────────────────────────
def long_path(p: str, always: bool = False) -> str:
    """윈도우에서 긴 경로를 열 수 있는 모양(\\\\?\\C:\\…)으로 바꾼다.
    짧은 경로는 그대로 돌려준다(always=True 면 항상 바꾼다)."""
    if not _WIN or not p or p.startswith("\\\\?\\"):
        return p
    ap = os.path.abspath(p)
    if not always and len(ap) < _LIMIT:
        return p
    if ap.startswith("\\\\"):                      # \\서버\공유폴더
        return "\\\\?\\UNC\\" + ap[2:]
    return "\\\\?\\" + ap


def plain_path(p: str) -> str:
    """\\\\?\\ 를 떼어 사람이 읽는 경로로 되돌린다."""
    if p.startswith("\\\\?\\UNC\\"):
        return "\\\\" + p[8:]
    if p.startswith("\\\\?\\"):
        return p[4:]
    return p


def exists(p: str) -> bool:
    return os.path.exists(long_path(p))


def isdir(p: str) -> bool:
    return os.path.isdir(long_path(p))


def getsize(p: str) -> int:
    try:
        return os.path.getsize(long_path(p))
    except OSError:
        return 0


def ensure_dir(p: str):
    if p:
        os.makedirs(long_path(p), exist_ok=True)


def write_text(path: str, text: str):
    ensure_dir(os.path.dirname(path))
    with io.open(long_path(path), "w", encoding="utf-8", newline="\n") as f:
        f.write(text)


def write_bytes(path: str, data: bytes):
    ensure_dir(os.path.dirname(path))
    with open(long_path(path), "wb") as f:
        f.write(data)


def read_text(path: str) -> str:
    with io.open(long_path(path), encoding="utf-8") as f:
        return f.read()


def copy_file(src: str, dst: str):
    ensure_dir(os.path.dirname(dst))
    shutil.copy2(long_path(src), long_path(dst))


def remove(p: str):
    try:
        os.remove(long_path(p))
    except OSError:
        pass


def remove_tree(p: str):
    shutil.rmtree(long_path(p), ignore_errors=True)


def remove_empty_dirs(top: str, stop: str):
    """top 부터 위로 올라가며 빈 폴더를 지운다. stop 폴더는 지우지 않는다."""
    top, stop = os.path.abspath(top), os.path.abspath(stop)
    while top.startswith(stop + os.sep):
        try:
            os.rmdir(long_path(top))
        except OSError:
            return
        top = os.path.dirname(top)


def walk(root: str):
    """os.walk 와 같되, 긴 경로 폴더를 건너뛰지 않는다.
    돌려주는 경로는 \\\\?\\ 를 뗀 보통 경로다."""
    for dp, dns, fns in os.walk(long_path(root, always=True)):
        yield plain_path(dp), dns, fns


def sha256_file(path: str, limit: int = 64 * 1024 * 1024) -> str:
    """앞 64MB 의 SHA-256. 크기를 함께 섞어, 앞부분만 같은 큰 파일을 구별한다."""
    h = hashlib.sha256()
    try:
        size = os.path.getsize(long_path(path))
        h.update(str(size).encode())
        with open(long_path(path), "rb") as f:
            read = 0
            while read < limit:
                b = f.read(1 << 20)
                if not b:
                    break
                h.update(b)
                read += len(b)
    except OSError:
        return ""
    return h.hexdigest()


def short_hash(text: str, n: int = 8) -> str:
    return hashlib.sha1(text.encode("utf-8")).hexdigest()[:n]


# ─────────────────────────────────────────────────────────────
# 2. 이름
# ─────────────────────────────────────────────────────────────
_INVALID = re.compile(r'[\\/:*?"<>|\x00-\x1f]')
_RESERVED = {"con", "prn", "aux", "nul", *(f"com{i}" for i in range(1, 10)),
             *(f"lpt{i}" for i in range(1, 10))}


def safe_name(name: str, maxlen: int = 80) -> str:
    """파일·폴더 이름으로 쓸 수 있게 다듬는다."""
    s = _INVALID.sub("_", name or "").strip()
    s = re.sub(r"\s{2,}", " ", s)
    s = s[:maxlen].rstrip(" .")
    if not s:
        s = "무제"
    if s.split(".")[0].lower() in _RESERVED:      # 윈도우 예약어 (CON, NUL …)
        s = "_" + s
    return s


# ─────────────────────────────────────────────────────────────
# 3. 마크다운에 안전하게 적기
# ─────────────────────────────────────────────────────────────
_URL_ESCAPE = set(' %()<>[]#?"`') | {chr(c) for c in range(32)}


def md_url(path: str) -> str:
    """링크 주소 칸에 넣을 경로. 공백·괄호 같은 것만 %XX 로 바꾸고 한글은 둔다."""
    p = path.replace("\\", "/")
    return "".join(f"%{ord(ch):02X}" if ch in _URL_ESCAPE else ch for ch in p)


def md_label(text: str) -> str:
    """링크 글자 칸([…])에 넣을 글. 대괄호를 막는다."""
    s = re.sub(r"\s+", " ", text or "").strip()
    return s.replace("\\", "\\\\").replace("[", "\\[").replace("]", "\\]")


def md_link(label: str, path: str, image: bool = False) -> str:
    return f"{'!' if image else ''}[{md_label(label)}]({md_url(path)})"


def yaml_str(value) -> str:
    """머리말(YAML) 값. JSON 문자열로 적으면 \\·\"·줄바꿈 모두 안전하다."""
    s = re.sub(r"[\r\n]+", " ", str(value if value is not None else ""))
    return json.dumps(s, ensure_ascii=False)
