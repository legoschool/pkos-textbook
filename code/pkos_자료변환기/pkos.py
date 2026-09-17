# -*- coding: utf-8 -*-
"""
자료 변환기 — 폴더를 넣으면 마크다운이 나온다
===============================================

    python 변환.py "D:\\흩어진자료"               변환
    python 변환.py "D:\\흩어진자료" --훑기          무엇이 있는지만 보기
    python 변환.py "이상한파일" --확인              이 파일이 뭔지만 알아보기
    python 변환.py "D:\\자료" --맛보기 20           앞 20개만 시험 변환
    python 변환.py --점검                          필요한 라이브러리 확인

파일 이름과 확장자는 믿지 않는다. 내용을 보고 무엇인지 판단한다.

PKOS(개인지식운영체계) 프로젝트
"""

from __future__ import annotations

import os
import sys
import argparse

# 윈도우 명령창에서 한글이 깨지지 않게
try:
    sys.stdout.reconfigure(encoding="utf-8")
    sys.stderr.reconfigure(encoding="utf-8")
except Exception:
    pass

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

from engine import fsutil                             # noqa: E402
from engine.sniff import sniff, sniff_folder          # noqa: E402
from engine.pipeline import Converter, Settings       # noqa: E402
from engine.privacy import Policy                     # noqa: E402
from engine import ocr as _ocr                        # noqa: E402


# ─────────────────────────────────────────────────────────────
필요한것 = [
    ("pymupdf",       "PDF",                      "pymupdf"),
    ("docx",          "워드 .docx",                "python-docx"),
    ("pptx",          "파워포인트 .pptx",           "python-pptx"),
    ("openpyxl",      "엑셀 .xlsx",                "openpyxl"),
    ("olefile",       "한글 .hwp · 옛 .doc/.ppt",  "olefile"),
    ("xlrd",          "옛 엑셀 .xls",              "xlrd"),
    ("bs4",           "웹문서 .html",              "beautifulsoup4"),
    ("PIL",           "그림 변환 · 자가진단",       "pillow"),
]


def 설치() -> int:
    """requirements.txt 의 라이브러리를 깐다.

    이 일을 배치 파일이 아니라 여기서 하는 이유 — 윈도우 배치 파일(.cmd)에
    한글을 넣으면 실행 시점 콘솔 코드페이지(949/65001)와 파일 인코딩이
    어긋나는 순간 파서가 엉켜 엉뚱한 줄을 명령으로 실행해 버린다. 그래서
    .cmd 는 순수 영문으로 두고, 한글 화면은 인코딩을 확실히 다루는
    파이썬 쪽에서 맡는다.
    """
    import subprocess
    here = os.path.dirname(os.path.abspath(__file__))
    req = os.path.join(here, "requirements.txt")
    if not os.path.exists(req):
        print("requirements.txt 를 찾지 못했습니다.")
        return 1

    # pip 는 콘솔에 바로 쓰고 이쪽 print 는 버퍼를 거치므로, 흘려보내지
    # 않으면 안내문이 pip 출력 뒤에 나온다.
    print("필요한 라이브러리를 설치합니다. 처음이면 몇 분 걸립니다.\n", flush=True)
    subprocess.run([sys.executable, "-m", "pip", "install",
                    "--disable-pip-version-check", "--quiet", "--upgrade", "pip"])
    r = subprocess.run([sys.executable, "-m", "pip", "install",
                        "--disable-pip-version-check", "-r", req])
    if r.returncode != 0:
        print("\n설치에 실패했습니다. 인터넷 연결을 확인하고 다시 해보세요.")
        return 1
    print("\n설치가 끝났습니다. 확인합니다.\n")
    return 점검()


def 마법사(s: Settings) -> int:
    """폴더를 끌어다 놓았을 때 — 먼저 훑어 보여주고, 물어본 뒤 변환한다."""
    print()
    print("  [1/2] 무엇이 들어 있는지 먼저 봅니다")
    print("  " + "─" * 46)
    Converter(s).scan()

    print()
    try:
        답 = input("  이대로 변환할까요? (y/n) ").strip().lower()
    except (EOFError, KeyboardInterrupt):
        답 = "n"
    if 답 not in ("y", "yes", "ㅛ"):
        print("\n  변환하지 않았습니다.")
        return 0

    print()
    print("  [2/2] 변환합니다")
    print("  " + "─" * 46)
    Converter(s).run()          # scan 이 임시폴더를 지웠으므로 새로 만든다
    return 0


def 점검() -> int:
    print("필요한 라이브러리\n")
    빠진것 = []
    for mod, 쓰임, 설치이름 in 필요한것:
        try:
            __import__(mod)
            print(f"  ✔ {쓰임:26} {설치이름}")
        except ImportError:
            print(f"  ✘ {쓰임:26} {설치이름}   ← 없음")
            빠진것.append(설치이름)

    print("\n글자 인식(OCR) 엔진 — 켜야 쓰인다 (--글자인식)")
    for name, (됨, 설명) in _ocr.available().items():
        print(f"  {'✔' if 됨 else '·'} {name:10} {설명}")

    if 빠진것:
        print("\n한 줄로 설치")
        print(f"  pip install {' '.join(빠진것)}")
        print("\n※ 한글 .hwpx 와 텍스트·JSON·자막은 라이브러리 없이도 읽힙니다.")
        return 1
    print("\n다 갖춰졌습니다.")
    return 0


def 확인(경로: str) -> int:
    """파일(또는 폴더)의 정체만 알아본다. 변환하지 않는다."""
    if fsutil.isdir(경로):
        rows = sniff_folder(경로)
        print(f"{경로}\n")
        print(f"{'파일':<46} {'판별':<18} {'근거'}")
        print("─" * 100)
        어긋남 = 0
        for p, k in rows:
            rel = os.path.relpath(p, 경로)
            표시 = (rel[:43] + "…") if len(rel) > 44 else rel
            mark = " ⚠" if k.mismatch else ""
            print(f"{표시:<46} {k.label + mark:<18} {k.why}")
            어긋남 += 1 if k.mismatch else 0
        print("─" * 100)
        print(f"{len(rows)}개 · 이름과 속이 다른 파일 {어긋남}개")
        return 0

    if not fsutil.exists(경로):
        print(f"그런 파일이 없습니다: {경로}")
        return 2

    k = sniff(경로)
    크기 = fsutil.getsize(경로)
    print(f"파일      {경로}")
    print(f"크기      {크기:,} 바이트 ({크기 / 1024 / 1024:.1f} MB)")
    print(f"판별      {k.label}  ({k.fmt})")
    print(f"갈래      {k.group}")
    print(f"근거      {k.why}")
    if k.ext:
        print(f"확장자    {k.ext}" + (f"  → {k.ext_fmt} 라고 주장" if k.ext_fmt else ""))
    if k.encoding:
        print(f"인코딩    {k.encoding}")
    if k.mismatch:
        print(f"\n⚠ {k.note}")
    elif k.note:
        print(f"\n{k.note}")
    return 0


# ─────────────────────────────────────────────────────────────
def main(argv=None) -> int:
    ap = argparse.ArgumentParser(
        prog="pkos.py",
        description="흩어진 자료를 형식·이름 상관없이 마크다운으로 바꿉니다.",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog=__doc__.split("PKOS")[0].split("=\n")[-1])

    ap.add_argument("대상", nargs="?", help="변환할 폴더 또는 파일")
    ap.add_argument("-o", "--결과", "--out", dest="out", default="",
                    help="결과 폴더 (기본: 원본 옆 _변환결과)")

    ap.add_argument("--훑기", "--scan", action="store_true",
                    help="변환하지 않고 무엇이 있는지만 본다")
    ap.add_argument("--확인", "--check", action="store_true",
                    help="파일의 정체만 알아본다")
    ap.add_argument("--점검", "--doctor", action="store_true",
                    help="필요한 라이브러리가 깔려 있는지 본다")
    ap.add_argument("--설치", "--setup", action="store_true",
                    help="필요한 라이브러리를 설치한다 (설치.cmd 가 부른다)")
    ap.add_argument("--마법사", "--wizard", action="store_true",
                    help="훑어 보여주고 물어본 뒤 변환한다 (변환하기.cmd 가 부른다)")
    ap.add_argument("--자가진단", "--selftest", action="store_true",
                    help="시험 자료를 만들어 변환기가 제대로 도는지 스스로 검사한다")
    ap.add_argument("--맛보기", "--limit", type=int, default=0, metavar="N",
                    help="앞 N개만 변환해 본다")

    ap.add_argument("--글자인식", "--ocr", default="없음",
                    choices=["없음", "tesseract", "claude"],
                    help="사진·스캔본에서 글자를 읽는 방법 (기본: 없음)")

    ap.add_argument("--개인정보끄기", action="store_true",
                    help="개인정보 가리기를 끈다 (기본은 켜짐)")
    ap.add_argument("--보고서원본숨김", action="store_true",
                    help="개인정보 보고서에 가리기 전 값을 남기지 않는다")

    ap.add_argument("--그림안꺼냄", action="store_true",
                    help="문서 속 그림을 꺼내지 않는다")
    ap.add_argument("--압축안열기", action="store_true",
                    help=".zip 을 풀지 않는다")
    ap.add_argument("--중복허용", action="store_true",
                    help="내용이 같은 파일도 각각 변환한다")
    ap.add_argument("--블로그분리끄기", action="store_true",
                    help="블로그 백업 PDF 도 보통 PDF 로 다룬다")

    ap.add_argument("--다시", "--force", action="store_true",
                    help="이미 만들어 둔 md 도 다시 만든다")
    ap.add_argument("--최대MB", type=float, default=300.0, metavar="MB",
                    help="이보다 큰 파일은 건너뛴다 (기본 300)")
    ap.add_argument("--조용히", "--quiet", action="store_true")

    a = ap.parse_args(argv)

    if a.점검:
        return 점검()

    if a.설치:
        return 설치()

    if a.자가진단:
        from tests.run_tests import main as 자가진단
        return 자가진단()

    if not a.대상:
        ap.print_help()
        print("\n예)  python pkos.py \"D:\\흩어진자료\" --훑기")
        return 2

    대상 = os.path.abspath(a.대상)
    if not fsutil.exists(대상):
        print(f"그런 경로가 없습니다: {대상}")
        return 2

    if a.확인:
        return 확인(대상)

    s = Settings(
        src=대상,
        out=os.path.abspath(a.out) if a.out else "",
        skip_existing=not a.다시,
        max_mb=a.최대MB,
        extract_images=not a.그림안꺼냄,
        open_archives=not a.압축안열기,
        dedup=not a.중복허용,
        blog_split=not a.블로그분리끄기,
        ocr_engine=a.글자인식,
        privacy_on=not a.개인정보끄기,
        policy=Policy(),
        report_original=not a.보고서원본숨김,
        verbose=not a.조용히,
    )

    if a.마법사:
        return 마법사(s)

    c = Converter(s)
    if a.훑기:
        c.scan()
        return 0

    r = c.run(limit=a.맛보기 or None)
    return 1 if r["failed"] and not r["done"] and not r["unchanged"] else 0


if __name__ == "__main__":
    try:
        sys.exit(main())
    except KeyboardInterrupt:
        print("\n중단했습니다. 다시 실행하면 이어서 진행합니다.")
        sys.exit(130)
