# -*- coding: utf-8 -*-
"""
pkos_자료변환기 엔진
=====================

흩어진 자료를 형식·이름 상관없이 마크다운으로 바꾸는 엔진 묶음.

    sniff.py           내용으로 형식 판별 (이 도구의 핵심)
    readers.py         판별된 형식으로 읽기 · 그림을 제자리에
    oldoffice.py       .doc / .ppt / .xls 읽기 (명세대로)
    images.py          사진 EXIF, 그림 형식 바꾸기
    ocr.py             글자 인식 (기본은 꺼짐)
    privacy_guard.py   개인정보 필터 보정 (문헌 번호·계좌·흔한 낱말)
    pipeline.py        폴더 훑기 → 변환 → 기록·목차
    fsutil.py          파일 입출력 · 긴 경로 · 링크·머리말에 안전하게 적기
    mdutil.py          마크다운 표 · 셀 값

기존 pkems_pdf_변환 에서 가져와 **손대지 않은** 것
    readers_legacy.py  한글 .hwp 이진 해석 등 (pkems_readers.py)
    privacy.py         개인정보 탐지·가리기  (pkems_privacy.py)
    blogpdf.py         블로그 백업 PDF 글 분리 (pkems_converter.py)

PKOS(개인지식운영체계) 프로젝트
"""

from .sniff import sniff, sniff_bytes, sniff_folder, FileKind
from .readers import read_by_kind, guess_title, ReadContext
from .pipeline import Converter, Settings, ENGINE_VERSION as __version__
from .privacy import Policy
from .fsutil import safe_name

__all__ = [
    "sniff", "sniff_bytes", "sniff_folder", "FileKind",
    "read_by_kind", "guess_title", "ReadContext",
    "Converter", "Settings", "Policy", "safe_name",
]
