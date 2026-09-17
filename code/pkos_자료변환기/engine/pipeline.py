# -*- coding: utf-8 -*-
"""
변환 파이프라인 — 폴더를 통째로 넣으면 마크다운이 나온다
===========================================================

    폴더 훑기 → 정체 판별 → (압축이면 풀어서 다시) → 바뀌었나? → 중복인가?
            → 읽기(그림은 제자리에) → 개인정보 가리기 → .md 쓰기 → 목차

**여러 번 돌리는 도구다.** 지식 체계는 자료가 계속 늘어난다. 그래서 변환 기록
(_index.json)을 남기고, 다음에 돌릴 때 그것을 이어받는다.

    · 크기·수정시각이 같으면 건너뛴다. 다르면 SHA-256 으로 내용까지 본다.
    · 내용이 바뀐 파일만 다시 변환한다. (처음 판은 한 번 만든 md 를 영원히 두었다)
    · 목차는 이전 기록과 합쳐서 쓴다. (처음 판은 새 파일 1개를 넣고 돌리면
      목차가 24개 → 2개로 줄었다 — 그번에 변환한 것만 적었기 때문이다)

결과 폴더
    _변환결과/
    ├── INDEX.md              사람이 읽는 목차
    ├── _index.json           변환 기록 (AI·웹이 읽는 목록이기도 하다)
    ├── _오류.md               못 읽은 것과 이유   · _오류_상세.log
    ├── _중복.md               같은 내용이던 파일들
    ├── _개인정보_보고서.md
    ├── assets/<결과 경로>/    문서마다 따로 두는 그림 폴더
    ├── 블로그/                블로그 백업 PDF → 글 단위 (백업끼리 겹치는 글은 한 번만)
    └── (원본 폴더 구조)/….md

PKOS(개인지식운영체계) 프로젝트
"""

from __future__ import annotations

import io
import os
import re
import json
import time
import shutil
import zipfile
import tempfile
import traceback
import collections
from dataclasses import dataclass, field

from . import fsutil as fs
from .fsutil import long_path, yaml_str, md_url, md_label, md_link, safe_name
from .sniff import sniff, is_junk_file, FileKind
from .readers import ReadContext, read_by_kind, guess_title, strip_leading_h1
from .readers_legacy import ReadResult, sanitize
from .privacy import Policy, PrivacyReport, Hit
from .privacy_guard import AcademicPrivacyFilter
from . import ocr as _ocr

TOOL = "pkos_자료변환기"
MANIFEST_VERSION = 2
# 읽는 방식이 바뀌면 올린다. 기록의 판이 다르면 원본이 그대로여도 다시 변환한다
# (그러지 않으면 도구를 고쳐도 이미 만든 결과에는 반영되지 않는다).
ENGINE_VERSION = "2.0"
_SKIP_DIRS = {"__pycache__", "node_modules", "_변환결과", "venv", ".venv"}


# ─────────────────────────────────────────────────────────────
@dataclass
class Settings:
    src: str                                    # 넣을 폴더 (또는 파일)
    out: str = ""                               # 결과 폴더 (비우면 원본 옆 _변환결과)

    max_mb: float = 300.0                       # 이보다 크면 건너뛴다
    min_chars: int = 10                         # 이보다 짧으면 '내용 없음' 표시
    keep_tree: bool = True                      # 원본 폴더 구조 유지
    skip_existing: bool = True                  # False 면 모두 다시 만든다 (--다시)

    open_archives: bool = True                  # .zip 을 풀어서 안까지 변환
    archive_depth: int = 2                      # 압축 안의 압축은 몇 겹까지

    extract_images: bool = True                 # 문서 속 그림을 제자리에 꺼내 두기
    min_image_bytes: int = 8000
    max_images_per_doc: int = 300
    copy_photos: bool = True                    # 사진 파일 자체도 assets 로 복사

    dedup: bool = True                          # 내용이 같으면 한 번만
    blog_split: bool = True                     # 블로그 백업 PDF 는 글 단위로

    ocr_engine: str = "없음"                     # 없음 / tesseract / claude

    privacy_on: bool = True
    policy: Policy = field(default_factory=Policy)
    report_original: bool = True

    verbose: bool = True

    def resolved_out(self) -> str:
        if self.out:
            return os.path.abspath(self.out)
        src = os.path.abspath(self.src)
        base = src if os.path.isdir(src) else os.path.dirname(src)
        return os.path.join(base, "_변환결과")


@dataclass
class Item:
    """변환 대상 하나. 압축 안의 파일도 이것으로 나타낸다."""
    path: str                       # 실제로 읽을 경로 (압축 안의 것은 풀어 둔 임시 파일)
    rel: str                        # 원본 기준 경로 ('/' 로 잇는다. 압축 안은 a.zip!속/파일)
    kind: FileKind | None = None
    size: int = 0
    mtime: float = 0.0
    from_archive: str = ""


# ─────────────────────────────────────────────────────────────
# 네이버 블로그 백업 PDF 알아보기
# ─────────────────────────────────────────────────────────────
def looks_like_blog_backup(path: str, probe_pages: int = 40) -> bool:
    """블로그 분리기(blogpdf)가 글을 가르는 표시 — '날짜시각' 한 줄 바로 다음에
    'blog.naver.com/아이디/글번호' 한 줄 — 가 앞쪽에 한 번이라도 있으면 백업이다.

    처음 판은 '앞 12쪽 중 2쪽 이상'을 요구했다. 실제 백업은 첫 글이 여러 쪽에
    걸치는 일이 흔해서 백업을 보통 PDF 로 잘못 넘겼다. 두 줄이 연달아 정확히 그
    모양인 경우는 일반 문서에서 우연히 나오지 않고, 분리기와 같은 규칙을 써야
    '알아봤는데 나누지 못함'이 생기지 않는다.
    """
    try:
        import pymupdf
        from .blogpdf import DATE_RE, URL_RE
    except ImportError:
        return False
    try:
        doc = pymupdf.open(long_path(path), filetype="pdf")
    except Exception:
        return False
    try:
        if doc.needs_pass:
            return False
        for i in range(min(probe_pages, doc.page_count)):
            lines = [l.strip() for l in doc[i].get_text().split("\n")]
            for a, b in zip(lines, lines[1:]):
                if DATE_RE.match(a) and URL_RE.match(b):
                    return True
        return False
    except Exception:
        return False
    finally:
        doc.close()


def _zip_member_name(info: zipfile.ZipInfo) -> str:
    """ZIP 안 파일 이름. 한국 윈도우에서 만든 압축은 이름을 cp949 로 적고 UTF-8
    표시를 켜지 않는다 — 파이썬은 그것을 cp437 로 읽어 '╟╤▒█' 처럼 깨뜨린다."""
    name = info.filename
    if info.flag_bits & 0x800:
        return name
    try:
        raw = name.encode("cp437")
    except UnicodeEncodeError:
        return name
    for enc in ("utf-8", "cp949"):
        try:
            return raw.decode(enc)
        except UnicodeDecodeError:
            continue
    return name


# ═════════════════════════════════════════════════════════════
class Converter:
    def __init__(self, s: Settings):
        self.s = s
        self.src = os.path.abspath(s.src)
        self.out = s.resolved_out()
        self.manifest: dict[str, dict] = {}       # 이전 실행의 기록
        self.prev_errors: dict[str, dict] = {}
        self.prev_dups: dict[str, dict] = {}
        self.prev_archives: dict[str, dict] = {}
        self.records: dict[str, dict] = {}        # 이번 결과
        self.errors: dict[str, dict] = {}
        self.dups: dict[str, dict] = {}
        self.archives: dict[str, dict] = {}
        self.kept_archives: set[str] = set()      # 바뀌지 않아 풀지 않은 압축
        self.vanished: dict[str, dict] = {}
        self.attempted: set[str] = set()
        self.current: set[str] = set()
        self.used_md: dict[str, str] = {}
        self.seen_hash: dict[str, str] = {}
        self.stats = collections.Counter()
        self.sniffed = 0                          # 이번에 실제로 열어 본 파일 수
        self.junk = 0
        self.debug: list[str] = []
        self._tmp: str | None = None
        self._zip_no = 0
        self.privacy = AcademicPrivacyFilter(s.policy) if s.privacy_on else None
        self.report = PrivacyReport(show_original=s.report_original)
        try:
            _ocr.set_engine(s.ocr_engine)
        except ValueError as e:
            self.log(f"⚠ {e} — 글자 인식을 끕니다")
            _ocr.set_engine("없음")

    def log(self, *a):
        if self.s.verbose:
            print(*a, flush=True)

    # ─────────────────────────────────────────────────────────
    # 변환 기록
    # ─────────────────────────────────────────────────────────
    def load_manifest(self):
        p = os.path.join(self.out, "_index.json")
        if not fs.exists(p):
            return
        try:
            data = json.loads(fs.read_text(p))
        except Exception:
            self.log("⚠ 이전 변환 기록(_index.json)을 읽지 못해 처음부터 만듭니다")
            return
        docs = data if isinstance(data, list) else data.get("문서", []) + data.get("원본없음", [])
        for rec in docs:
            if isinstance(rec, dict) and rec.get("src") and rec.get("md"):
                self.manifest[rec["src"]] = rec
                self.used_md[rec["md"].lower()] = rec["src"]
        if isinstance(data, dict):
            self.prev_errors = {e["file"]: e for e in data.get("오류", []) if "file" in e}
            self.prev_dups = {d["file"]: d for d in data.get("중복", []) if "file" in d}
            self.prev_archives = {a["src"]: a for a in data.get("압축", []) if "src" in a}

    @staticmethod
    def _unchanged(rec: dict | None, it: Item) -> bool:
        """원본 크기·수정시각과 도구 판이 모두 같은가. 사진은 글자 인식 설정까지 같아야 한다
        (나중에 글자 인식을 켜면 전에 변환한 사진도 다시 읽어야 하므로)."""
        if not rec or rec.get("size") != it.size or rec.get("engine") != ENGINE_VERSION:
            return False
        if abs(float(rec.get("mtime", -1)) - it.mtime) >= 2:
            return False
        if rec.get("group") == "그림" and rec.get("ocr", "없음") != _ocr.current_engine():
            return False
        return True

    def _stamp(self, it: Item) -> dict:
        """기록마다 붙이는 공통 값 — 다음 실행에서 '바뀌었나'와 '무슨 형식인가'를 알아내는 데 쓴다."""
        k = it.kind
        return {"size": it.size, "mtime": it.mtime, "engine": ENGINE_VERSION,
                "ocr": _ocr.current_engine(), "format": k.fmt, "label": k.label,
                "group": k.group, "why": k.why, "ext": k.ext, "encoding": k.encoding,
                "ext_mismatch": k.mismatch}

    @staticmethod
    def _kind_from(rec: dict) -> FileKind | None:
        if not rec.get("format"):
            return None
        return FileKind(fmt=rec["format"], label=rec.get("label", ""), group=rec.get("group", "기타"),
                        why=rec.get("why", ""), ext=rec.get("ext", ""),
                        mismatch=bool(rec.get("ext_mismatch")), encoding=rec.get("encoding", ""))

    def _in_kept_archive(self, rel: str) -> bool:
        return "!" in rel and any(rel.startswith(a + "!") for a in self.kept_archives)

    def _present(self, rel: str) -> bool:
        return rel in self.current or self._in_kept_archive(rel)

    # ─────────────────────────────────────────────────────────
    # 대상 모으기 — 확장자로 거르지 않고 전부 열어 본다
    # ─────────────────────────────────────────────────────────
    def _tmpdir(self) -> str:
        if self._tmp is None:
            self._tmp = tempfile.mkdtemp(prefix="pkos_")
        return self._tmp

    def cleanup(self):
        if self._tmp:
            shutil.rmtree(self._tmp, ignore_errors=True)
            self._tmp = None

    def _make_item(self, path: str, rel: str, from_archive: str = "",
                   mtime: float | None = None) -> Item:
        try:
            st = os.stat(long_path(path))
            size, mt = st.st_size, st.st_mtime
        except OSError:
            size, mt = 0, 0.0
        it = Item(path=path, rel=rel, size=size, mtime=mt if mtime is None else mtime,
                  from_archive=from_archive)
        # 바뀌지 않은 파일은 지난번 판별을 그대로 쓴다 (파일을 열지 않는다)
        for prev in (self.manifest.get(rel), self.prev_errors.get(rel),
                     self.prev_dups.get(rel), self.prev_archives.get(rel)):
            if prev and prev.get("format") and self._unchanged(prev, it):
                it.kind = self._kind_from(prev)
                break
        if it.kind is None:
            it.kind = sniff(path)
            self.sniffed += 1
        return it

    def collect(self) -> list[Item]:
        items: list[Item] = []
        self.junk = 0
        if os.path.isfile(long_path(self.src)):
            items.append(self._make_item(self.src, os.path.basename(self.src)))
        else:
            out_abs = os.path.normcase(self.out)
            for dp, dns, fns in fs.walk(self.src):
                keep = []
                for d in sorted(dns):
                    full = os.path.normcase(os.path.join(dp, d))
                    if d.startswith(".") or d in _SKIP_DIRS or full == out_abs or \
                            full.startswith(out_abs + os.sep):
                        continue
                    keep.append(d)
                dns[:] = keep
                for fn in sorted(fns):
                    if is_junk_file(fn):
                        self.junk += 1
                        continue
                    p = os.path.join(dp, fn)
                    rel = os.path.relpath(p, self.src).replace("\\", "/")
                    items.append(self._make_item(p, rel))
        if self.s.open_archives:
            items += self._archive_members(items, 0)
        return items

    def _archive_members(self, items: list[Item], depth: int) -> list[Item]:
        """items 안의 .zip 을 풀어 **새로 나온 파일만** 돌려준다 (압축 속 압축은 depth 겹까지)."""
        added: list[Item] = []
        for it in items:
            if not it.kind or it.kind.fmt != "zip":
                continue
            if self.s.skip_existing and self._unchanged(self.prev_archives.get(it.rel), it):
                self.kept_archives.add(it.rel)          # 바뀌지 않았다 — 풀지 않고 지난 결과를 쓴다
                continue
            try:
                zf = zipfile.ZipFile(long_path(it.path))
            except Exception as e:
                self._error(it.rel, f"압축을 열지 못했습니다: {e}", "압축", it)
                continue
            self._zip_no += 1
            base = os.path.join(self._tmpdir(), f"z{self._zip_no}")
            with zf:
                for info in zf.infolist():
                    if info.is_dir():
                        continue
                    name = _zip_member_name(info)
                    parts = [safe_name(p, 60) for p in name.replace("\\", "/").split("/")
                             if p not in ("", ".", "..")]
                    if not parts or name.startswith("__MACOSX/") or is_junk_file(parts[-1]):
                        continue
                    rel = f"{it.rel}!{'/'.join(parts)}"
                    if info.flag_bits & 0x1:
                        self._error(rel, "암호가 걸린 압축이라 풀지 못했습니다", "압축")
                        continue
                    if info.file_size > self.s.max_mb * 1024 * 1024:
                        self._error(rel, f"너무 큼 ({info.file_size / 1048576:.0f}MB)", "압축")
                        continue
                    dst = os.path.join(base, *parts)
                    try:
                        fs.ensure_dir(os.path.dirname(dst))
                        with zf.open(info) as fin, open(long_path(dst), "wb") as fout:
                            shutil.copyfileobj(fin, fout)
                    except Exception as e:
                        self._error(rel, f"압축을 푸는 중 오류: {e}", "압축")
                        continue
                    try:
                        mt = time.mktime(info.date_time + (0, 0, -1))
                    except Exception:
                        mt = it.mtime
                    added.append(self._make_item(dst, rel, from_archive=it.rel, mtime=mt))
        if added and depth + 1 < self.s.archive_depth:
            added += self._archive_members(added, depth + 1)
        return added

    # ─────────────────────────────────────────────────────────
    # 훑어보기 — 변환 없이 무엇이 있는지만
    # ─────────────────────────────────────────────────────────
    def scan(self) -> dict:
        self.load_manifest()
        try:
            items = self.collect()
        finally:
            self.cleanup()
        by_label, by_group = collections.Counter(), collections.Counter()
        mismatched, total, new = [], 0, 0
        for it in items:
            k = it.kind
            by_label[k.label] += 1
            by_group[k.group] += 1
            total += it.size
            if k.mismatch:
                mismatched.append((it.rel, k.ext, k.label))
            if not self._unchanged(self.manifest.get(it.rel), it):
                new += 1

        self.log(f"대상 : {self.src}")
        self.log(f"파일 : {len(items):,}개 · {total / 1048576:,.0f} MB"
                 + (f"  (잡동사니 {self.junk}개 제외)" if self.junk else ""))
        if self.manifest:
            self.log(f"       이전에 변환한 기록 {len(self.manifest):,}개 · "
                     f"새로 만들거나 바뀐 것 {new:,}개")
        if self.kept_archives:
            self.log(f"       압축 {len(self.kept_archives)}개는 지난번과 같아 풀지 않았습니다 "
                     f"(안의 파일은 위 숫자에 없음)")
        self.log("\n  갈래별")
        for g, n in by_group.most_common():
            self.log(f"    {g:8} {n:6,}개")
        self.log("\n  형식별")
        for l, n in by_label.most_common(20):
            self.log(f"    {l:16} {n:6,}개")
        if mismatched:
            self.log(f"\n  ⚠ 이름과 속이 다른 파일 {len(mismatched)}개 — 내용 기준으로 변환합니다")
            for rel, ext, label in mismatched[:10]:
                self.log(f"    {rel}  ({ext or '확장자 없음'} → {label})")
            if len(mismatched) > 10:
                self.log(f"    … 외 {len(mismatched) - 10}개")
        self.log(f"\n  저장 위치 : {self.out}")
        return {"files": len(items), "mb": round(total / 1048576), "new": new,
                "by_label": dict(by_label), "by_group": dict(by_group),
                "mismatched": len(mismatched), "junk": self.junk}

    # ─────────────────────────────────────────────────────────
    # 실행
    # ─────────────────────────────────────────────────────────
    def run(self, limit: int | None = None) -> dict:
        t0 = time.time()
        fs.ensure_dir(self.out)
        self.load_manifest()
        counts = collections.Counter()
        try:
            items = self.collect()
            self.current = {it.rel for it in items}
            # 바뀌지 않은 이전 기록의 내용 지문을 먼저 올려 둔다 (사본 판정용).
            # 맛보기(limit)로 일부만 돌려도 판정이 흔들리지 않게 자르기 전에 한다.
            by_rel = {it.rel: it for it in items}
            for src, rec in self.manifest.items():
                if rec.get("sha256") and (self._in_kept_archive(src) or
                                          (src in by_rel and self._unchanged(rec, by_rel[src]))):
                    self.seen_hash.setdefault(rec["sha256"], src)
            if limit:
                items = items[:limit]

            total = len(items)
            self.log(f"파일 {total:,}개를 살펴봅니다."
                     + (f"  (글자 인식: {_ocr.current_engine()})" if _ocr.is_on() else ""))
            for i, it in enumerate(items, 1):
                if it.size > 20 * 1048576:
                    self.log(f"  … 큰 파일 처리 중: {it.rel} ({it.size / 1048576:.0f}MB)")
                try:
                    r = self._one(it)
                except Exception as e:
                    self._error(it.rel, f"처리 중 오류 ({type(e).__name__}): {e}",
                                it.kind.label if it.kind else "", it)
                    self.debug.append(f"[{it.rel}]\n{traceback.format_exc()}")
                    r = "실패"
                counts[r] += 1
                if i % 25 == 0 or i == total:
                    self.log(f"  {i:,}/{total:,}  변환 {counts['변환']} · 그대로 {counts['그대로']}"
                             f" · 사본 {counts['사본']} · 실패 {counts['실패']}")
            full_scan = limit is None and os.path.isdir(long_path(self.src))
            self.write_outputs(items, full_scan)
        finally:
            self.cleanup()

        secs = int(time.time() - t0)
        self.log(f"\n완료 — 이번에 변환 {counts['변환']}개 · 그대로 둠 {counts['그대로']}개 · "
                 f"사본 {counts['사본']}개 · 실패 {counts['실패']}개 · {secs // 60}분 {secs % 60}초")
        self.log(f"       전체 기록: 문서 {len(self.records):,}개 · 사본 {len(self.dups):,}개 · "
                 f"못 읽음 {len(self.errors):,}개"
                 + (f" · 풀지 않은 압축 {len(self.kept_archives)}개(지난번과 같음)"
                    if self.kept_archives else ""))
        if self.stats:
            self.log("\n  형식별 (이번에 변환한 것)")
            for k, n in self.stats.most_common(20):
                self.log(f"    {k:18} {n:5}개")
        if self.errors:
            self.log(f"\n  못 읽은 파일 {len(self.errors)}개 → _오류.md")
        if self.vanished:
            self.log(f"  원본이 사라진 기록 {len(self.vanished)}개 → 목차 맨 아래")
        if self.privacy:
            self.log("\n" + "─" * 46)
            self.log(self.report.summary())
            if self.report.rows and self.s.report_original:
                self.log("  ⚠ 보고서에는 가리기 전 원본이 들어 있습니다. 공유하지 마세요.")
        self.log(f"\n목차 : {os.path.join(self.out, 'INDEX.md')}")
        return {"done": counts["변환"], "unchanged": counts["그대로"], "duplicates": counts["사본"],
                "failed": counts["실패"], "errors": len(self.errors), "records": len(self.records),
                "sniffed": self.sniffed, "kept_archives": len(self.kept_archives),
                "stats": dict(self.stats)}

    def _error(self, rel: str, msg: str, kind: str = "", it: Item | None = None):
        entry = {"file": rel, "error": msg, "kind": kind}
        if it is not None and it.kind is not None:
            entry.update(self._stamp(it))
        self.errors[rel] = entry

    # ── 결과 경로
    def _md_path(self, it: Item, prev: dict | None) -> tuple[str, str]:
        """→ (절대경로, 결과폴더 기준 '/' 경로). 이전 기록이 있으면 그 경로를 그대로 쓴다."""
        if prev and prev.get("md"):
            rel = prev["md"]
        else:
            parts = it.rel.replace("!", "/").split("/")
            stem, ext = os.path.splitext(parts[-1])
            base = f"{stem}{ext.replace('.', '_')}"
            dirs = [safe_name(p, 60) for p in parts[:-1]] if self.s.keep_tree else []
            rel = "/".join(dirs + [safe_name(base, 80) + ".md"])
            owner = self.used_md.get(rel.lower())
            if owner is not None and owner != it.rel:
                # 'a.b' 와 'a_b' 처럼 결과 이름이 겹친다 — 원본 경로로 짧은 꼬리를 붙인다
                rel = "/".join(dirs + [safe_name(base, 70) + f"_{fs.short_hash(it.rel)}.md"])
        self.used_md[rel.lower()] = it.rel
        return os.path.join(self.out, *rel.split("/")), rel

    # ── 파일 하나
    def _one(self, it: Item) -> str:
        k = it.kind
        self.attempted.add(it.rel)
        if k.fmt == "zip" and self.s.open_archives:
            self.archives[it.rel] = {"src": it.rel, **self._stamp(it)}
            return "압축"                                  # 안의 파일이 따로 대상에 들어 있다
        if it.size > self.s.max_mb * 1048576:
            self._error(it.rel, f"너무 큼 ({it.size / 1048576:.0f}MB · 한도 {self.s.max_mb:.0f}MB)",
                        k.label, it)
            return "실패"

        prev = self.manifest.get(it.rel)
        sha = ""
        if self.s.skip_existing:
            # ① 지난번에 변환했고 바뀌지 않았다
            if prev and fs.exists(os.path.join(self.out, *prev["md"].split("/"))):
                if self._unchanged(prev, it):
                    self.used_md[prev["md"].lower()] = it.rel
                    self.records[it.rel] = prev
                    return "그대로"
                sha = fs.sha256_file(it.path)
                if sha and prev.get("sha256") == sha and prev.get("engine") == ENGINE_VERSION \
                        and prev.get("group") != "그림":
                    prev.update(size=it.size, mtime=it.mtime)    # 시각만 바뀌고 내용은 같다
                    self.records[it.rel] = prev
                    return "그대로"
            # ② 지난번에 못 읽었고 바뀌지 않았다 — 글자 인식 설정이 같고 '설치 필요'가 아니면 다시 열지 않는다
            perr = self.prev_errors.get(it.rel)
            if perr and self._unchanged(perr, it) and perr.get("ocr") == _ocr.current_engine() \
                    and "설치 필요" not in perr.get("error", ""):
                self.errors[it.rel] = perr
                return "실패"
            # ③ 지난번에 사본이었고, 원본이 아직 있다
            pdup = self.prev_dups.get(it.rel)
            if pdup and self._unchanged(pdup, it) and self._present(pdup.get("same_as", "")):
                self.dups[it.rel] = pdup
                return "사본"

        if self.s.dedup:
            sha = sha or fs.sha256_file(it.path)
            first = self.seen_hash.get(sha)
            if sha and first and first != it.rel:
                self.dups[it.rel] = {"file": it.rel, "same_as": first, "sha256": sha,
                                     **self._stamp(it)}
                return "사본"

        if k.fmt == "pdf" and self.s.blog_split and looks_like_blog_backup(it.path):
            return self._blog(it, sha, prev)

        md_abs, md_rel = self._md_path(it, prev)
        # 그림 폴더는 결과 파일 경로를 따라간다 → 문서마다 따로이고 겹칠 수 없다
        asset_rel = "assets/" + md_rel[:-3]
        asset_abs = os.path.join(self.out, *asset_rel.split("/"))
        fs.remove_tree(asset_abs)                            # 다시 변환할 때 옛 그림을 지운다
        want_assets = self.s.extract_images or (k.is_image and self.s.copy_photos)
        ctx = ReadContext(
            asset_dir=asset_abs if want_assets else "",
            asset_url=os.path.relpath(asset_abs, os.path.dirname(md_abs)).replace("\\", "/"),
            min_bytes=self.s.min_image_bytes, limit=self.s.max_images_per_doc)

        res = read_by_kind(it.path, k, ctx)
        if not res.ok:
            fs.remove_tree(asset_abs)
            fs.remove_empty_dirs(os.path.dirname(asset_abs), self.out)
            self._error(it.rel, res.error, k.label, it)
            self.stats[f"실패:{k.label}"] += 1
            return "실패"

        body = sanitize(res.text)
        stem = os.path.splitext(it.rel.replace("!", "/").split("/")[-1])[0]
        title, title_from = guess_title(body, stem, str(res.meta.get("제목후보", "") or ""))
        body = strip_leading_h1(body, title)

        n_priv = 0
        if self.privacy:
            body, hits = self.privacy.mask_markdown(body)
            names = {h.original.replace(" ", ""): h.masked for h in hits if h.kind == "이름"}
            for nm, masked in names.items():                 # 본문에서 확인된 이름은 제목에서도
                title = title.replace(nm, masked)
            path_hits = [Hit("파일이름에 이름", nm, "(파일 이름은 그대로)", 0, 0,
                             f"경로: {it.rel}", "확실") for nm in names if nm in it.rel]
            if hits or path_hits:
                self.report.add(it.rel, hits + path_hits)
                n_priv = len(hits)

        note = ""
        if len(body.strip()) < self.s.min_chars:
            note = ("> ⚠ 글자를 거의 찾지 못했습니다. 그림 위주이거나 스캔한 문서일 수 있습니다.\n\n")
            self.stats["내용거의없음"] += 1

        fs.write_text(md_abs, self._front_matter(it, res, title, title_from, sha,
                                                 len(body), len(ctx.saved)) + note + body + "\n")
        if not ctx.saved:
            fs.remove_empty_dirs(asset_abs, self.out)

        self.records[it.rel] = {
            "src": it.rel, "md": md_rel, "title": title, "chars": len(body),
            "images": len(ctx.saved), "assets": asset_rel if ctx.saved else "",
            "date": self._date(it, res), "sha256": sha, "archive": it.from_archive,
            "privacy": n_priv, "tags": res.meta.get("tags", []), **self._stamp(it),
        }
        if sha:
            self.seen_hash.setdefault(sha, it.rel)
        self.stats[k.label] += 1
        return "변환"

    @staticmethod
    def _date(it: Item, res: ReadResult) -> str:
        if res.meta and res.meta.get("찍은날짜"):
            return str(res.meta["찍은날짜"])
        return time.strftime("%Y-%m-%d", time.localtime(it.mtime)) if it.mtime else ""

    def _front_matter(self, it: Item, res: ReadResult, title: str, title_from: str,
                      sha: str, chars: int, n_images: int) -> str:
        k = it.kind
        L = ["---", f"title: {yaml_str(title)}"]
        date = self._date(it, res)
        if date:
            L.append(f"date: {date}")
        L += [f"source: {yaml_str(it.rel)}",
              f"format: {yaml_str(k.fmt)}",
              f"format_label: {yaml_str(k.label)}",
              f"group: {yaml_str(k.group)}",
              f"detected: {yaml_str(k.why)}"]
        if k.ext:
            L.append(f"ext: {yaml_str(k.ext)}")
        if k.mismatch:
            L.append("ext_mismatch: true")
        if title_from != "파일이름":
            L.append(f"title_from: {yaml_str(title_from)}")
        if it.from_archive:
            L.append(f"archive: {yaml_str(it.from_archive)}")
        if sha:
            L.append(f"sha256: {yaml_str(sha[:16])}")
        L.append(f"chars: {chars}")
        if n_images:
            L.append(f"images: {n_images}")
        skip = {"인코딩", "제목후보", "원래머리말", "tags", "찍은날짜", "그림"}
        extra = [f"{kk} {vv}" for kk, vv in (res.meta or {}).items()
                 if kk not in skip and vv not in (None, "", [], {})]
        if extra:
            L.append(f"info: {yaml_str(' · '.join(map(str, extra)))}")
        L.append(f"tags: {json.dumps(res.meta.get('tags', []) or [], ensure_ascii=False)}")
        front = res.meta.get("원래머리말") if res.meta else None
        if front:
            L.append("original_front_matter: |")
            L += ["  " + ln for ln in str(front).splitlines()]
        rel_shown = it.rel.replace("`", "'")
        note = f"*원본 `{rel_shown}`"
        if k.mismatch:
            note += f" · ⚠ 이름은 {k.ext} 인데 실제로는 {k.label}"
        if date:
            note += f" · {date}"
        L += ["---", "", f"# {title}", "", note + "*", "", ""]
        return "\n".join(L)

    # ─────────────────────────────────────────────────────────
    # 블로그 백업 PDF → 글 단위
    # ─────────────────────────────────────────────────────────
    def _blog(self, it: Item, sha: str, prev: dict | None) -> str:
        from .blogpdf import Converter as BlogConverter, Settings as BlogSettings
        out_dir = os.path.join(self.out, "블로그")
        index_path = os.path.join(out_dir, "_index.json")
        if prev is not None or not self.s.skip_existing:
            # 전에 변환한 백업을 다시 변환한다 (원본이 바뀌었거나 도구가 바뀌었거나 --다시).
            # 블로그 분리기는 이미 아는 글을 늘 건너뛰므로, 이 백업에서 나온 글을 먼저 뺀다.
            # 다른 백업에서 온 글은 그대로 두어 백업끼리의 겹침 제거는 유지한다.
            self._drop_blog_posts(out_dir, os.path.basename(it.path))
        before = set()
        if fs.exists(index_path):
            try:
                before = {e["file"] for e in json.loads(fs.read_text(index_path))}
            except Exception:
                before = set()

        self.log(f"  📘 블로그 백업으로 보입니다 → 글 단위로 쪼갭니다 : {it.rel}")
        bc = BlogConverter(BlogSettings(
            pdf_dir=long_path(os.path.dirname(it.path), always=True),
            out_dir=long_path(out_dir, always=True),
            extract_images=self.s.extract_images,
            min_image_bytes=self.s.min_image_bytes,
            skip_existing=True, verbose=False))
        r = bc.run(pdf_names=[os.path.basename(it.path)])
        added = int(r.get("added", 0))
        overlap = int((r.get("stats") or {}).get("중복건너뜀", 0))
        if not r.get("total"):
            self.log("     … 글 경계를 찾지 못해 보통 PDF 로 변환합니다")
            saved = self.s.blog_split
            self.s.blog_split = False
            try:
                return self._one_retry(it)
            finally:
                self.s.blog_split = saved

        new_files = [e for e in bc.index if e["file"] not in before]
        for e in new_files:
            self._fix_blog_post(it, out_dir, e)
        self._write_blog_index(out_dir, bc.index)

        self.records[it.rel] = {
            "src": it.rel, "md": "블로그/INDEX.md",
            "title": f"{os.path.splitext(os.path.basename(it.rel))[0]} (블로그 백업)",
            "chars": 0, "images": 0, "assets": "", "date": "", "sha256": sha,
            "archive": it.from_archive, "privacy": 0, "tags": [], **self._stamp(it),
            "blog": True, "posts_added": added, "posts_overlap": overlap,
        }
        if sha:
            self.seen_hash.setdefault(sha, it.rel)
        self.stats["블로그 백업"] += 1
        self.log(f"     … 새 글 {added}편 · 다른 백업과 겹쳐 건너뜀 {overlap}편")
        return "변환"

    @staticmethod
    def _drop_blog_posts(out_dir: str, pdf_name: str):
        index_path = os.path.join(out_dir, "_index.json")
        if not fs.exists(index_path):
            return
        try:
            index = json.loads(fs.read_text(index_path))
        except Exception:
            return
        keep = []
        for e in index:
            if e.get("src") != pdf_name:
                keep.append(e)
                continue
            fs.remove(os.path.join(out_dir, e["file"]))
            fs.remove_tree(os.path.join(out_dir, "images", os.path.splitext(e["file"])[0]))
        if len(keep) != len(index):
            fs.write_text(index_path, json.dumps(keep, ensure_ascii=False, indent=1))

    def _one_retry(self, it: Item) -> str:
        self.attempted.discard(it.rel)
        return self._one(it)

    def _fix_blog_post(self, it: Item, out_dir: str, e: dict):
        """블로그 글 파일 보정 — 머리말 이스케이프, 그림 링크, 개인정보."""
        p = os.path.join(out_dir, e["file"])
        try:
            text = fs.read_text(p)
        except OSError:
            return
        m = re.match(r"\A---\n.*?\n---\n", text, re.S)
        body = text[m.end():] if m else text
        # 블로그 분리기는 그림을 한 줄에 하나씩 ![](images/<제목>/img_001.jpeg) 로 적는다.
        # 제목에 괄호가 있으면 첫 ')' 에서 끊기므로 줄 끝의 ')' 까지 통째로 잡는다.
        body = re.sub(r"^!\[\]\((.+)\)[ \t]*$", lambda x: f"![]({md_url(x.group(1))})", body,
                      flags=re.M)
        if self.privacy:
            body, hits = self.privacy.mask_markdown(body)
            if hits:
                self.report.add(f"{it.rel} › {e['file']}", hits)
        front = ["---", f"title: {yaml_str(e.get('title', ''))}",
                 f"date: {e.get('date', '')}", f"time: {yaml_str(e.get('time', ''))}",
                 f"source: {yaml_str(it.rel)}", f"category: {yaml_str(e.get('category', ''))}",
                 f"url: {yaml_str(e.get('url', ''))}", 'format: "blogpost"',
                 f"images: {e.get('images', 0)}", "tags: []", "---", ""]
        fs.write_text(p, "\n".join(front) + body)

    @staticmethod
    def _write_blog_index(out_dir: str, index: list[dict]):
        by_year = collections.defaultdict(list)
        for e in index:
            by_year[str(e.get("date", ""))[:4] or "날짜없음"].append(e)
        L = ["# 📚 블로그 기록 목차", "", f"전체 **{len(index):,}편** · 자동 생성", "",
             "| 연도 | 편수 |", "|------|-----:|"]
        for y in sorted(by_year, reverse=True):
            L.append(f"| {y} | {len(by_year[y])} |")
        for y in sorted(by_year, reverse=True):
            L += ["", f"## {y}년 ({len(by_year[y])}편)", ""]
            for e in sorted(by_year[y], key=lambda x: (x.get("date", ""), x.get("time", "")),
                            reverse=True):
                cat = f" · {e['category']}" if e.get("category") else ""
                L.append(f"- {e.get('date', '')} · [{md_label(e.get('title', ''))}]"
                         f"({md_url(e['file'])}){cat}")
        fs.write_text(os.path.join(out_dir, "INDEX.md"), "\n".join(L) + "\n")

    # ─────────────────────────────────────────────────────────
    # 목차·기록·오류·중복·개인정보 보고서
    # ─────────────────────────────────────────────────────────
    def write_outputs(self, items: list[Item], full_scan: bool):
        # 이번에 다루지 않은 지난 기록을 이어받는다.
        #   맛보기(limit) 로 일부만 돌렸거나, 바뀌지 않아 풀지 않은 압축 안의 것들이다.
        #   전체를 훑었는데 원본이 없으면 '원본없음' 으로 옮긴다 (변환본은 지우지 않는다).
        for src, rec in self.manifest.items():
            if src in self.records or src in self.attempted:
                continue
            if full_scan and not self._present(src):
                self.vanished[src] = rec
            else:
                self.records[src] = rec
        for prev, now in ((self.prev_errors, self.errors), (self.prev_dups, self.dups)):
            for f, e in prev.items():
                if f not in self.attempted and f not in now and \
                        (not full_scan or self._in_kept_archive(f)):
                    now[f] = e
        for a, e in self.prev_archives.items():
            if a not in self.archives and (not full_scan or a in self.current):
                self.archives[a] = e

        recs = sorted(self.records.values(), key=lambda r: r["src"])
        fs.write_text(os.path.join(self.out, "_index.json"), json.dumps({
            "도구": TOOL, "판": MANIFEST_VERSION, "엔진": ENGINE_VERSION,
            "갱신": time.strftime("%Y-%m-%d %H:%M"), "원본": self.src, "문서": recs,
            "오류": sorted(self.errors.values(), key=lambda e: e["file"]),
            "중복": sorted(self.dups.values(), key=lambda d: d["file"]),
            "압축": sorted(self.archives.values(), key=lambda a: a["src"]),
            "원본없음": sorted(self.vanished.values(), key=lambda r: r["src"]),
        }, ensure_ascii=False, indent=1))

        self._write_index(recs)
        self._write_table("_오류.md", bool(self.errors), self._errors_md)
        self._write_table("_중복.md", bool(self.dups), self._dups_md)
        if self.debug:
            fs.write_text(os.path.join(self.out, "_오류_상세.log"), "\n\n".join(self.debug))
        self._write_privacy_report()

    def _write_index(self, recs: list[dict]):
        by_group = collections.defaultdict(list)
        for r in recs:
            by_group[r.get("group", "기타")].append(r)
        n_img = sum(int(r.get("images", 0) or 0) for r in recs)
        n_mis = sum(1 for r in recs if r.get("ext_mismatch"))
        L = ["# 📂 변환 목차", "",
             f"전체 **{len(recs):,}개** · {time.strftime('%Y-%m-%d %H:%M')} 갱신", ""]
        if n_img:
            L.append(f"- 문서에서 꺼낸 그림 {n_img:,}장 → `assets/`")
        if n_mis:
            L.append(f"- 이름과 속이 달랐던 파일 {n_mis}개 (내용 기준으로 변환)")
        if self.dups:
            L.append(f"- 같은 내용이라 건너뛴 사본 {len(self.dups)}개 → {md_link('_중복.md', '_중복.md')}")
        if self.errors:
            L.append(f"- 못 읽은 파일 {len(self.errors)}개 → {md_link('_오류.md', '_오류.md')}")
        L += ["", "| 갈래 | 개수 |", "|------|-----:|"]
        order = sorted(by_group, key=lambda g: -len(by_group[g]))
        for g in order:
            L.append(f"| {g} | {len(by_group[g]):,} |")
        for g in order:
            L += ["", f"## {g} ({len(by_group[g]):,}개)", ""]
            for r in by_group[g]:
                bits = []
                if r.get("blog"):
                    bits.append(f"블로그 글 새로 {r.get('posts_added', 0)}편")
                elif r.get("chars"):
                    bits.append(f"{int(r['chars']):,}자")
                if r.get("images"):
                    bits.append(f"그림 {r['images']}장")
                if r.get("ext_mismatch"):
                    bits.append("⚠이름불일치")
                tail = (" · " + " · ".join(bits)) if bits else ""
                src = str(r["src"]).replace("`", "'")
                L.append(f"- [{md_label(r.get('title', ''))}]({md_url(r['md'])}){tail} · `{src}`")
        if self.vanished:
            L += ["", f"## 원본이 사라진 문서 ({len(self.vanished)}개)", "",
                  "원본 폴더에서 없어졌지만 변환본은 남겨 두었습니다.", ""]
            for r in sorted(self.vanished.values(), key=lambda x: x["src"]):
                L.append(f"- [{md_label(r.get('title', ''))}]({md_url(r['md'])}) · `{r['src']}`")
        fs.write_text(os.path.join(self.out, "INDEX.md"), "\n".join(L) + "\n")

    def _write_table(self, name: str, has: bool, builder):
        p = os.path.join(self.out, name)
        if has:
            fs.write_text(p, builder())
        else:
            fs.remove(p)                                     # 지난번 것이 남아 헷갈리지 않게

    @staticmethod
    def _cell(s) -> str:
        return re.sub(r"\s+", " ", str(s)).replace("|", "\\|").strip()

    def _errors_md(self) -> str:
        L = ["# ⚠ 변환하지 못한 파일", "", f"{len(self.errors)}개", "",
             "| 파일 | 종류 | 이유와 해결 방법 |", "|---|---|---|"]
        for e in sorted(self.errors.values(), key=lambda x: x["file"]):
            L.append(f"| `{self._cell(e['file'])}` | {self._cell(e.get('kind', ''))} | "
                     f"{self._cell(e['error'])} |")
        L += ["", "---", "", "### 자주 나오는 경우", "",
              "**글자가 없는 PDF·사진** — 스캔하거나 찍은 문서입니다. `--글자인식 tesseract` "
              "를 붙여 다시 돌리거나, 원본 문서 파일이 있으면 그것을 넣으세요.", "",
              "**암호가 걸린 파일** — 만든 프로그램에서 암호를 풀고 저장한 뒤 넣으세요.", "",
              "**구글 문서·시트·슬라이드** — 드라이브에서 .docx/.xlsx/.pptx 로 내려받아 넣으세요.",
              "", "**.7z / .rar** — 직접 푼 뒤 폴더를 넣으세요. `.zip` 은 자동으로 풉니다."]
        if self.debug:
            L += ["", "*프로그램 오류의 자세한 기록은 `_오류_상세.log` 에 있습니다.*"]
        return "\n".join(L) + "\n"

    def _dups_md(self) -> str:
        L = ["# 🔁 같은 내용이라 건너뛴 파일", "",
             f"{len(self.dups)}개. 내용(SHA-256)이 완전히 같아 한 번만 변환했습니다.", "",
             "| 건너뛴 파일 | 이미 변환한 것 |", "|---|---|"]
        for d in sorted(self.dups.values(), key=lambda x: x["file"]):
            L.append(f"| `{self._cell(d['file'])}` | `{self._cell(d['same_as'])}` |")
        return "\n".join(L) + "\n"

    def _write_privacy_report(self):
        rep_md = os.path.join(self.out, "_개인정보_보고서.md")
        rep_json = os.path.join(self.out, "_개인정보.json")
        if self.privacy is None:
            return
        # 이번에 다루지 않은 파일의 지난 기록은 이어받는다
        if fs.exists(rep_json):
            try:
                old = json.loads(fs.read_text(rep_json)).get("rows", [])
            except Exception:
                old = []
            for row in old:
                src = str(row.get("file", "")).split(" › ")[0]
                if src in self.attempted or src in self.vanished:
                    continue
                row.setdefault("original", "(이전 실행에서 숨김)")
                self.report.rows.append(row)
                self.report.counter[row.get("kind", "?")] += 1
        if self.report.rows:
            written = self.report.write(long_path(self.out, always=True))
            if not written:
                fs.remove(rep_md)
        else:
            fs.remove(rep_md)
            fs.remove(rep_json)
