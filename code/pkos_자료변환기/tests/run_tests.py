# -*- coding: utf-8 -*-
"""
자가진단 — 변환기가 제대로 도는지 스스로 검사한다
===================================================

    python pkos.py --자가진단
    python tests/run_tests.py            (같은 것)
    python tests/run_tests.py --남기기    (시험 폴더를 지우지 않고 위치를 알려 준다)

시험 자료를 임시 폴더에 새로 만들고(make_fixtures.py), 실제 변환을 돌린 뒤 결과를
하나하나 판정한다. 개발하며 실제로 겪은 결함마다 시험이 하나씩 있다.

다섯 번 돌린다.
    1회  처음 변환
    2회  아무것도 안 바꾸고 다시     → 모두 '그대로', 목차 개수 유지
    3회  파일 하나 추가 + 하나 수정   → 그 둘만 다시 변환, 목차는 합쳐짐
    4회  --맛보기 3                  → 목차가 줄지 않음
    5회  원본 하나 삭제               → '원본이 사라진 문서'로 옮겨짐, 변환본은 남음
"""

from __future__ import annotations

import io
import os
import re
import sys
import json
import time
import shutil
import hashlib
import tempfile
import urllib.parse

HERE = os.path.dirname(os.path.abspath(__file__))
ROOT = os.path.dirname(HERE)
sys.path.insert(0, ROOT)
sys.path.insert(0, HERE)

try:
    sys.stdout.reconfigure(encoding="utf-8")
except Exception:
    pass

from engine import Converter, Settings                    # noqa: E402
from engine.sniff import sniff                            # noqa: E402
from engine.fsutil import long_path                       # noqa: E402
from engine.privacy_guard import AcademicPrivacyFilter    # noqa: E402
import make_fixtures                                      # noqa: E402


class Report:
    def __init__(self):
        self.rows: list[tuple[str, str, str, str]] = []   # (구역, 결과, 이름, 설명)
        self.section = ""

    def part(self, name: str):
        self.section = name

    def check(self, name: str, ok, detail: str = ""):
        self.rows.append((self.section, "통과" if ok else "실패", name, "" if ok else detail))

    def skip(self, name: str, why: str):
        self.rows.append((self.section, "건너뜀", name, why))

    def show(self) -> int:
        cur = None
        for sec, res, name, detail in self.rows:
            if sec != cur:
                print(f"\n[{sec}]")
                cur = sec
            mark = {"통과": "✔", "실패": "✘", "건너뜀": "·"}[res]
            print(f"  {mark} {name}" + (f"\n      → {detail}" if detail else ""))
        n = {r: sum(1 for x in self.rows if x[1] == r) for r in ("통과", "실패", "건너뜀")}
        print(f"\n통과 {n['통과']} · 실패 {n['실패']} · 건너뜀 {n['건너뜀']}")
        return 1 if n["실패"] else 0


def read(path: str) -> str:
    with io.open(long_path(path), encoding="utf-8") as f:
        return f.read()


def run(src: str, out: str, limit=None, **kw) -> dict:
    s = Settings(src=src, out=out, verbose=False, **kw)
    return Converter(s).run(limit=limit)


def load_index(out: str) -> dict:
    return json.loads(read(os.path.join(out, "_index.json")))


def main(argv=None) -> int:
    argv = sys.argv[1:] if argv is None else argv
    keep = "--남기기" in argv or "--keep" in argv
    rep = Report()
    tmp = tempfile.mkdtemp(prefix="pkos_자가진단_")
    src, out = os.path.join(tmp, "자료"), os.path.join(tmp, "결과")
    print(f"시험 자료를 만듭니다 … {tmp}")
    make_fixtures.build(src)
    t0 = time.time()
    try:
        _run_all(rep, src, out)
    finally:
        code = rep.show()
        print(f"({time.time() - t0:.1f}초)")
        if keep:
            print(f"\n시험 폴더를 남겼습니다: {tmp}")
        else:
            shutil.rmtree(long_path(tmp, always=True), ignore_errors=True)
    return code


def _run_all(rep: Report, src: str, out: str):
    # ═════════════════════════════════════════════════════════
    rep.part("형식 판별 — 이름이 아니라 내용으로")
    cases = [
        ("사실은워드.pdf", "docx", True), ("확장자없는발표", "pptx", False),
        ("글/쉼표메모.txt", "txt", False), ("글/명단.tsv", "tsv", False),
        ("글/유니코드메모.txt", "txt", False), ("글/cp949메모.txt", "txt", False),
        ("글/빈파일.txt", "empty", False), ("글/웹페이지.html", "html", False),
        ("한글/한글문서.hwpx", "hwpx", False), ("압축묶음.zip", "zip", False),
        ("서식/서식파일.dotx", "docx", False), ("서식/엑셀인데이름이틀림.dat", "xlsx", False),
        ("옛형식/암호워드.docx", "encrypted_office", True), ("옛형식/옛엑셀.xls", "xls", False),
    ]
    for rel, fmt, mismatch in cases:
        p = os.path.join(src, *rel.split("/"))
        if not os.path.exists(long_path(p)):
            rep.skip(f"{rel} → {fmt}", "시험 파일 없음")
            continue
        k = sniff(p)
        rep.check(f"{rel} → {fmt}" + (" (이름불일치 표시)" if mismatch else ""),
                  k.fmt == fmt and k.mismatch == mismatch, f"{k.fmt}, 불일치={k.mismatch}, {k.why}")

    # ═════════════════════════════════════════════════════════
    rep.part("1회 — 처음 변환")
    r1 = run(src, out)
    rep.check("끝까지 돎 (예외로 멈추지 않음)", True)
    idx = load_index(out)
    docs = {d["src"]: d for d in idx["문서"]}
    errs = {e["file"]: e for e in idx["오류"]}
    dups = {d["file"]: d["same_as"] for d in idx["중복"]}

    def md_of(rel: str) -> str:
        d = docs.get(rel)
        return read(os.path.join(out, *d["md"].split("/"))) if d else ""

    rep.check("파일 하나가 오류여도 나머지를 변환", r1["done"] >= 25, f"변환 {r1['done']}개")
    for name, needle in (("메모(.txt)", "깊은곳의_메모"), ("워드", "깊은곳의_워드"), ("PDF", "깊은곳의_PDF")):
        hit = [s for s in docs if needle in s]
        err = [e["error"] for f, e in errs.items() if needle in f]
        rep.check(f"긴 경로(260자↑)의 {name}도 변환", hit, err[0] if err else "목록에 없음")
    junk = [s for s in list(docs) + list(errs) if re.search(r"desktop\.ini|Thumbs\.db|\.lnk$", s)]
    rep.check("윈도우 잡동사니는 대상에서 뺌", not junk, str(junk))
    rep.check("사용자 폴더 'assets' 도 변환", "assets/내자료.txt" in docs)

    # ── 이름·형식
    d = docs.get("사실은워드.pdf", {})
    rep.check("이름이 .pdf 인 워드를 워드로 변환", d.get("format") == "docx" and d.get("ext_mismatch"))
    rep.check("확장자 없는 PPT 를 변환", docs.get("확장자없는발표", {}).get("format") == "pptx")
    rep.check("서식파일(.dotx)을 변환", "서식/서식파일.dotx" in docs, errs.get("서식/서식파일.dotx", {}).get("error", ""))
    rep.check("쇼파일(.ppsx)을 변환", "서식/쇼파일.ppsx" in docs, errs.get("서식/쇼파일.ppsx", {}).get("error", ""))
    rep.check("이름이 .xlsx 가 아닌 엑셀을 변환", "서식/엑셀인데이름이틀림.dat" in docs,
              errs.get("서식/엑셀인데이름이틀림.dat", {}).get("error", ""))
    rep.check("매크로 발표(.pptm)를 변환", "그룹 안 글상자 1" in md_of("서식/매크로발표.pptm"),
              errs.get("서식/매크로발표.pptm", {}).get("error", ""))
    smi, vtt = md_of("글/강의자막.smi"), md_of("글/영상자막.vtt")
    rep.check("SAMI 자막(.smi, cp949)에서 대사만", "첫 번째 자막입니다" in smi and "두 번째 자막입니다" in smi
              and "SYNC" not in smi and "&nbsp;" not in smi)
    rep.check("WebVTT 자막(.vtt)에서 대사만", "웹 자막 한 줄입니다" in vtt and "-->" not in vtt)

    # ── 워드
    w = md_of("워드문서.docx")
    a, b, c = w.find("표 앞 문단"), w.find("| 항목"), w.find("표 뒤 문단")
    rep.check("워드 표가 본문 제자리에", 0 <= a < b < c, f"위치 {a}, {b}, {c}")
    rep.check("워드 글상자 글을 읽음", "글상자 안의 글" in w)
    rep.check("워드 머리글을 읽음", "머리글에만 있는 글" in w)
    img = re.search(r"!\[[^\]]*\]\(([^)]+)\)", w)
    rep.check("워드 그림이 제자리에 (둘째 절 뒤, 그림 뒤 문단 앞)",
              img and w.find("## 둘째 절") < img.start() < w.find("그림 뒤 문단"))
    rep.check("워드 제목 스타일 → # 제목", re.search(r"^## 둘째 절$", w, re.M))

    # ── PPT
    p = md_of("발표자료.pptx")
    rep.check("PPT 그룹 도형 안의 글을 읽음", "그룹 안 글상자 1" in p and "그룹 안 글상자 2" in p)
    rep.check("PPT 슬라이드 제목이 구획 제목", re.search(r"^## 2\. 그룹 도형 슬라이드$", p, re.M))
    rep.check("PPT 표를 읽음", "| 이름 | 점수 |" in p)
    rep.check("PPT 발표자 노트를 읽음", "발표자 노트 내용입니다" in p)
    s3 = p.find("## 3.")
    rep.check("PPT 그림이 해당 슬라이드 안에", s3 >= 0 and re.search(r"!\[", p[s3:]))

    # ── 엑셀·표
    x = md_of("성적표.xlsx")
    rep.check("엑셀 날짜를 날짜로 (00:00:00 꼬리 없음)", "2026-03-02" in x and "00:00:00" not in x)
    rep.check("엑셀 정수를 정수로", "| 1 | 국어 | 95 |" in x)
    rep.check("TSV 를 탭으로 나눈 표로", "| 번호 | 이름 | 반 |" in md_of("글/명단.tsv"))
    rep.check("cp949 CSV", "| 국어 | 90 |" in md_of("글/점수.csv"))
    rep.check("쉼표 섞인 .txt 는 표로 바꾸지 않음", "|---" not in md_of("글/쉼표메모.txt"))

    # ── 글자파일
    rep.check("cp949 텍스트", "옛 윈도우 메모장" in md_of("글/cp949메모.txt"))
    rep.check("UTF-16 텍스트", "UTF-16 으로 저장한" in md_of("글/유니코드메모.txt"))
    n = md_of("글/노트.md")
    head, _, body = n.partition("\n---\n")
    rep.check("원래 머리말(YAML)을 본문에 남기지 않음", "title: 옵시디언" not in body)
    rep.check("원래 머리말의 tags 를 이어받음", '"수업"' in head and '"파이썬"' in head)
    h = md_of("글/웹페이지.html")
    rep.check("웹: div 로만 된 글", "div 로만 된 첫 글" in h and "span 안의 글" in h)
    rep.check("웹: 표를 표로", "| 가 | 나 |" in h)
    rep.check("웹: <pre> 코드", "print('코드')" in h)

    # ── 한글 hwpx
    k = md_of("한글/한글문서.hwpx")
    rep.check("한글: 서식이 바뀐 낱말에 공백이 끼지 않음", "안녕하세요" in k, "‘안 녕하세요’ 처럼 끊김")
    rep.check("한글: 탭 뒤의 글을 읽음", "뒤의 글" in k)
    rep.check("한글: 표 속 표가 바깥 칸을 덮어쓰지 않음", "| 가 | 나 |" in k and "속1" in k)
    rep.check("한글: 구역 11개를 순서대로 (9 다음 10)", 0 <= k.find("구역 9") < k.find("구역 10"))
    rep.check("한글: 그림을 제자리에", re.search(r"!\[[^\]]*\]\([^)]+\.png\)", k))

    # ── 오픈도큐먼트
    od = md_of("한글/공문.odt")
    rep.check("ODT: 문단 속 틀 안의 표를 읽음", "| 항목 | 내용 |" in od and "합친 칸" in od,
              "한글에서 내보낸 ODT 는 표가 틀 속 글상자에 있다")
    rep.check("ODT: 공백 요소(text:s)를 공백으로", "A   B" in od)
    rep.check("ODT: 그림과 제목", re.search(r"!\[", od) and re.search(r"^# 공문 제목$", od, re.M))
    rep.check("ODT: 표 칸 안의 그림도 꺼냄 (그림 2장)", docs.get("한글/공문.odt", {}).get("images") == 2,
              f"그림 {docs.get('한글/공문.odt', {}).get('images')}장")
    rep.check("ODT: 칸 속 글상자 글을 한 번만 적음", od.count("칸 속 글상자 글") == 1,
              f"{od.count('칸 속 글상자 글')}번")

    # ── ZIP
    zips = [s for s in docs if s.startswith("압축묶음.zip!")]
    rep.check("한국 윈도우 압축(cp949 이름)의 파일 이름이 깨지지 않음",
              any(s.endswith("압축속_한글문서.txt") for s in zips), str(zips))
    rep.check("압축 속 하위 폴더까지 변환", any(s.endswith("하위/압축속_메모.md") for s in zips))

    # ── PDF
    f = md_of("글자PDF.pdf")
    rep.check("PDF 쪽 표시와 본문", "<!-- 쪽 2 -->" in f and "PDF 둘째 쪽 본문" in f)
    rep.check("PDF 그림이 그 쪽에", re.search(r"<!-- 쪽 2 -->.*!\[", f, re.S))
    rep.check("암호 PDF 는 '암호'라고 안내", "암호" in errs.get("암호PDF.pdf", {}).get("error", ""))
    rep.check("스캔 PDF 는 '글자가 없다'고 안내", "글자가 없는" in errs.get("스캔PDF.pdf", {}).get("error", ""))
    rep.check("빈 파일은 '빈 파일'이라고 안내", "빈 파일" in errs.get("글/빈파일.txt", {}).get("error", ""))
    rep.check("글자가 짧아도 그림이 없으면 정상 PDF 로 변환 (스캔본으로 거절하지 않음)",
              any("깊은곳의_PDF" in s for s in docs))

    # ── 블로그 백업
    b1, b2 = docs.get("블로그백업/1_2.pdf", {}), docs.get("블로그백업/2_3.pdf", {})
    rep.check("블로그 백업을 알아봄 (첫 글이 여러 쪽이어도)", b1.get("blog") and b2.get("blog"),
              f"{b1.get('format')}, {b2.get('format')}")
    blog_index = os.path.join(out, "블로그", "_index.json")
    posts = json.loads(read(blog_index)) if os.path.exists(long_path(blog_index)) else []
    rep.check("겹치는 백업 사이의 같은 글은 한 번만 (글 3편)", len(posts) == 3, f"{len(posts)}편")
    rep.check("겹친 글 수를 기록", b2.get("posts_overlap") == 1, str(b2.get("posts_overlap")))
    first_post = next((p for p in posts if p.get("url", "").endswith("220000000001")), None)
    fp = read(os.path.join(out, "블로그", first_post["file"])) if first_post else ""
    rep.check("긴 첫 글의 다음 쪽 내용까지 한 글로", "69번째 줄" in fp)
    rep.check("블로그 글 제목의 따옴표가 머리말을 깨지 않음", 'title: "드론 수업 \\"첫날\\" 기록"' in fp)
    rep.check("쪽 바닥글(N · 블로그이름)을 본문에서 뺌", "시험블로그" not in fp)
    second = next((p for p in posts if p.get("url", "").endswith("220000000002")), None)
    sp = read(os.path.join(out, "블로그", second["file"])) if second else ""
    rep.check("블로그 글에도 개인정보 가리기", sp and "010-9876-5432" not in sp)

    # ── 겹침
    ra, rb = docs.get("같은제목/보고서.docx", {}), docs.get("같은제목/보고서.pptx", {})
    rep.check("제목이 같은 두 문서의 그림 폴더가 따로",
              ra.get("assets") and rb.get("assets") and ra["assets"] != rb["assets"])
    pa, pb = docs.get("사진/가/IMG_0001.jpg", {}), docs.get("사진/나/IMG_0001.jpg", {})
    rep.check("같은 이름 사진 두 장이 따로 보관됨", pa.get("assets") and pb.get("assets")
              and pa["assets"] != pb["assets"])
    rep.check("사진 제목은 찍은 날짜로", pa.get("title") == "사진 2019-07-21", pa.get("title", ""))
    sv = md_of("사진/도식.svg")
    rep.check("SVG 그림을 보관하고 안의 글자를 꺼냄", "수업 흐름도" in sv and re.search(r"!\[", sv),
              errs.get("사진/도식.svg", {}).get("error", ""))
    rep.check("'a.b' 와 'a_b' 의 결과가 겹치지 않음", "글/a.b" in docs and "글/a_b" in docs
              and docs["글/a.b"]["md"] != docs["글/a_b"]["md"])
    rep.check("내용이 같은 사본은 한 번만", "사본/성적표 - 복사본.xlsx" in dups)
    t = docs.get("글/문서1.txt", {}).get("title", "")
    rep.check("이름이 쓸모없으면 본문 첫 줄이 제목", t == 'C:\\Users\\"따옴표" 경로 설명', t)

    # ── 옛 형식 (미리 만든 시험 파일이 있을 때)
    if "옛형식/옛엑셀.xls" in docs or "옛형식/옛엑셀.xls" in errs:
        o = md_of("옛형식/옛엑셀.xls")
        rep.check("옛 엑셀: 정수를 정수로 (3500000.0 아님)", "3500000" in o and "3500000.0" not in o)
        rep.check("옛 엑셀: 날짜를 날짜로", "2026-03-02" in o)
        rep.check("옛 엑셀: 숨긴 시트 표시", "(숨긴 시트)" in o)
    else:
        rep.skip("옛 엑셀", "tests/fixtures/옛엑셀.xls 없음")
    if os.path.exists(os.path.join(src, "옛형식", "암호워드.docx")):
        rep.check("암호 걸린 오피스 문서는 '암호'라고 안내",
                  "암호" in errs.get("옛형식/암호워드.docx", {}).get("error", ""))

    # ── 머리말·링크
    bad_yaml, bad_links = [], []
    md_files = [os.path.join(dp, fn) for dp, _, fns in os.walk(long_path(out, always=True))
                for fn in fns if fn.endswith(".md") and not fn.startswith("_")]
    for path in md_files:
        text = read(path)
        if text.startswith("---\n"):
            for ln in text.split("\n---\n", 1)[0].splitlines()[1:]:
                m = re.match(r'^[A-Za-z_]+:\s*(".*")\s*$', ln)
                if m:
                    try:
                        json.loads(m.group(1))
                    except Exception:
                        bad_yaml.append(f"{os.path.basename(path)}: {ln[:50]}")
        no_code = re.sub(r"^(`{3,}).*?^\1", "", text, flags=re.S | re.M)   # 코드 블록 속 글은 링크가 아니다
        for dest in re.findall(r"\]\(([^)\s]+)\)", no_code):
            if dest.startswith(("http://", "https://", "#")):
                continue
            target = os.path.normpath(os.path.join(os.path.dirname(path),
                                                   urllib.parse.unquote(dest)))
            if not os.path.exists(target):
                bad_links.append(f"{os.path.basename(path)} → {dest}")
    rep.check("모든 머리말 문자열이 올바름 (\\ 와 \" 포함)", not bad_yaml, "; ".join(bad_yaml[:3]))
    rep.check("모든 링크(그림·목차)가 실제 파일을 가리킴", not bad_links, "; ".join(bad_links[:3]))

    # ── 개인정보
    pm = md_of("개인정보/함정.txt")
    rep.check("주소 근처의 진짜 카드번호를 가림", "4111-1111-1111-1111" not in pm)
    rep.check("'계좌: 신한 110-…' 을 가림", "110-123-456789" not in pm)
    rep.check("'계좌번호(경남은행) 207-…' 을 가림", "207-0012-3456-78" not in pm)
    rep.check("DOI 는 가리지 않음", "1365480216659733" in pm)
    rep.check("'이름 설정' 의 '설정'을 이름으로 가리지 않음", "설정" in pm and "설*" not in pm)
    rep.check("진짜 이름·전화는 가림", "홍**" in pm and "010-1234-5678" not in pm)
    rep.check("본문에서 확인된 이름은 제목에서도 가림",
              docs.get("개인정보/홍길동 상담기록.txt", {}).get("title", "").startswith("홍**"))
    photo_doc = md_of("개인정보/홍길동 사진자료.docx")
    rep.check("가리기가 그림 링크를 깨지 않음", re.search(r"!\[[^\]]*\]\([^)]*홍길동[^)]*\)", photo_doc))
    rep.check("개인정보 보고서를 만듦", os.path.exists(long_path(os.path.join(out, "_개인정보_보고서.md"))))

    f2 = AcademicPrivacyFilter()
    # 라벨 줄에서 이름이 확정된 뒤, 같은 낱말이 다른 곳에 어떤 꼬리로 나오는지 본다
    for text, should_mask, label in (
            ("담당자 홍길동\n담당자를 홍길동으로 변경합니다.", True, "'홍길동으로' 는 이름을 놓치지 않음"),
            ("이름: 이상수\n이상수에 따르면", True, "'이상수에' 는 이름을 놓치지 않음"),
            ("작성자 설계\n수업을 설계하는 방법", False, "'설계하는' 이 있으면 이름이 아님"),
    ):
        hits = [hh for hh in f2.find(text) if hh.kind == "이름"]
        rep.check(f"이름 판정: {label}", bool(hits) == should_mask, str([hh.original for hh in hits]))

    # 실제 자료(대학원 폴더 215개)에서 나온 오탐과, 그것을 고치다 놓치면 안 되는 것
    from engine.privacy_guard import _luhn_ok
    assert not _luhn_ok("1234567890123456"), "시험용 번호가 검증식을 통과하면 안 된다"
    for text, kind, should_mask, label in (
            ("연도별 통계 2016 2017 2018 2019 2020 합계", "카드번호", False, "표 속 연도 나열"),
            ("문서확인번호:1234567890123456", "카드번호", False, "검증식을 통과 못 한 문서 번호"),
            ("카드 1234-5678-9012-3456 로 결제", "카드번호", True, "'카드' 표시가 있으면 검증식과 상관없이"),
            ("https://doi.org/10. 1177/1365480216659733", "카드번호", False, "공백으로 끊긴 DOI"),
            ("Review, 11. doi: 10.1023/ A:  1022193728205. Petrich", "주민등록번호", False,
             "Springer 형식 DOI (10.1023/A:…)"),
            ("https://doi.org/10.1007/A:\n1022193728205 Gao, H.", "주민등록번호", False,
             "줄이 끊긴 Springer DOI"),
            ("주민등록번호: 900101-1234567", "주민등록번호", True, "진짜 주민등록번호는 그대로 가림"),
            ("보기 https://example.com/v?id=3&date=2019-03-15&p=2", "생년월일", False, "주소 속 날짜"),
            ("담당: 이상수 교수\n이상수(2024)에 따르면", "이름", False, "인용된 저자"),
            ("학생 김민지 결석, 김민지(2023년 입학)", "이름", True, "연도로 끝나지 않는 괄호는 인용 아님"),
    ):
        hits = [hh for hh in f2.find(text) if hh.kind == kind]
        rep.check(f"실제 오탐 대응: {label}", bool(hits) == should_mask, str([hh.original for hh in hits]))

    # ═════════════════════════════════════════════════════════
    rep.part("2회 — 아무것도 바꾸지 않고 다시")
    r2 = run(src, out)
    idx2 = load_index(out)
    rep.check("다시 변환한 것 없음", r2["done"] == 0, f"변환 {r2['done']}개")
    rep.check("목차 개수 유지", len(idx2["문서"]) == len(idx["문서"]),
              f"{len(idx['문서'])} → {len(idx2['문서'])}")
    rep.check("바뀌지 않은 파일은 열어 보지도 않음 (판별 재사용)", r2["sniffed"] == 0,
              f"{r2['sniffed']}개를 다시 열었음")
    rep.check("바뀌지 않은 압축은 다시 풀지 않음", r2["kept_archives"] >= 1)
    rep.check("실패·사본 기록도 이어받음", len(idx2["오류"]) == len(idx["오류"]) and
              len(idx2["중복"]) == len(idx["중복"]),
              f"오류 {len(idx['오류'])}→{len(idx2['오류'])} · 사본 {len(idx['중복'])}→{len(idx2['중복'])}")
    rep.check("압축 안의 변환 결과도 이어받음", sum(1 for dd in idx2["문서"] if "!" in dd["src"]) ==
              sum(1 for dd in idx["문서"] if "!" in dd["src"]))

    r2b = run(src, out, ocr_engine="tesseract")
    idx2b = load_index(out)
    scan_err = {e["file"]: e for e in idx2b["오류"]}.get("스캔PDF.pdf", {}).get("error", "")
    rep.check("글자 인식을 켜면 전에 실패한 스캔 PDF 를 다시 시도", "글자 인식" in scan_err, scan_err)
    run(src, out)                     # 글자 인식을 끈 원래 상태로 되돌린다

    # ═════════════════════════════════════════════════════════
    rep.part("3회 — 새 파일 추가 + 기존 파일 수정")
    with io.open(os.path.join(src, "새로추가.txt"), "w", encoding="utf-8") as fh:
        fh.write("나중에 추가한 메모입니다.\n")
    target = os.path.join(src, "글", "cp949메모.txt")
    with io.open(target, "a", encoding="cp949") as fh:
        fh.write("고친 줄입니다.\n")
    os.utime(target, (time.time() + 5, time.time() + 5))
    r3 = run(src, out)
    idx3 = load_index(out)
    docs3 = {dd["src"]: dd for dd in idx3["문서"]}
    rep.check("추가·수정한 두 파일만 다시 변환", r3["done"] == 2, f"변환 {r3['done']}개")
    rep.check("목차에 새 파일이 더해짐 (기존 목록 유지)",
              len(idx3["문서"]) == len(idx["문서"]) + 1, f"{len(idx['문서'])} → {len(idx3['문서'])}")
    rep.check("수정한 내용이 반영됨",
              "고친 줄입니다" in read(os.path.join(out, *docs3["글/cp949메모.txt"]["md"].split("/"))))

    # 도구를 고쳐 판 번호가 달라진 기록은 원본이 그대로여도 다시 만든다
    manifest_path = os.path.join(out, "_index.json")
    data = load_index(out)
    for dd in data["문서"]:
        if dd["src"] in ("성적표.xlsx", "블로그백업/1_2.pdf"):
            dd["engine"] = "1.0"
    with io.open(long_path(manifest_path), "w", encoding="utf-8") as fh:
        json.dump(data, fh, ensure_ascii=False)
    r3b = run(src, out)
    rep.check("도구 판이 바뀐 기록은 다시 변환 (고친 내용이 반영됨)", r3b["done"] == 2,
              f"변환 {r3b['done']}개")
    b1b = {dd["src"]: dd for dd in load_index(out)["문서"]}.get("블로그백업/1_2.pdf", {})
    posts_b = json.loads(read(os.path.join(out, "블로그", "_index.json")))
    rep.check("다시 변환한 블로그 백업의 글을 새로 만듦 (다른 백업 글은 유지, 중복 없음)",
              b1b.get("posts_added") == 2 and len(posts_b) == 3,
              f"새로 {b1b.get('posts_added')}편 · 전체 {len(posts_b)}편")

    # ═════════════════════════════════════════════════════════
    rep.part("4회 — 맛보기(앞 3개만)")
    run(src, out, limit=3)
    rep.check("목차가 줄지 않음", len(load_index(out)["문서"]) == len(idx3["문서"]))

    # ═════════════════════════════════════════════════════════
    rep.part("5회 — 원본 하나를 지움")
    gone = docs3["새로추가.txt"]
    os.remove(os.path.join(src, "새로추가.txt"))
    run(src, out)
    idx5 = load_index(out)
    rep.check("원본이 사라진 문서로 옮겨짐", any(v["src"] == "새로추가.txt" for v in idx5["원본없음"])
              and all(v["src"] != "새로추가.txt" for v in idx5["문서"]))
    rep.check("변환본은 지우지 않음", os.path.exists(long_path(os.path.join(out, *gone["md"].split("/")))))


if __name__ == "__main__":
    sys.exit(main())
