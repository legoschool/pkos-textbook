# -*- coding: utf-8 -*-
"""
글자 인식(OCR) 자리 — 기본은 꺼져 있고, 엔진을 골라 켠다
=========================================================

사진과 스캔 PDF 는 글자가 '그림'이라 그냥 읽을 수 없다. 기본값은 `없음` 이다.

    python pkos.py "폴더" --글자인식 tesseract
    python pkos.py "폴더" --글자인식 claude

켜면 두 곳에서 쓰인다.
    · 사진 파일(.jpg .png …)
    · 글자가 하나도 없는 PDF — 쪽마다 그림으로 그려서 읽는다

엔진을 새로 붙이려면 함수 하나를 만들어 ENGINES 에 넣는다.

    def 내엔진(data: bytes, ext: str, lang: str) -> OcrResult
        성공: OcrResult(True, text=…)   실패: OcrResult(False, error="이유")
        예외를 밖으로 던지지 않는다 (일괄 변환이 멈추면 안 되므로)

PKOS(개인지식운영체계) 프로젝트
"""

from __future__ import annotations

import io
import os
from dataclasses import dataclass

from .fsutil import long_path


@dataclass
class OcrResult:
    ok: bool
    text: str = ""
    engine: str = "없음"
    error: str = ""


# ─────────────────────────────────────────────────────────────
# 1. 없음 — 기본값
# ─────────────────────────────────────────────────────────────
def _engine_none(data: bytes, ext: str, lang: str) -> OcrResult:
    return OcrResult(False, engine="없음", error="글자 인식이 꺼져 있습니다.")


# ─────────────────────────────────────────────────────────────
# 2. Tesseract — 무료 · 내 PC 안에서만 처리
# ─────────────────────────────────────────────────────────────
#   (1) 프로그램 : winget install UB-Mannheim.TesseractOCR   (설치 때 Korean 선택)
#   (2) 파이썬   : pip install pytesseract pillow
def _engine_tesseract(data: bytes, ext: str, lang: str) -> OcrResult:
    try:
        import pytesseract
        from PIL import Image
    except ImportError as e:
        return OcrResult(False, engine="tesseract",
                         error=f"설치 필요: pip install pytesseract pillow ({e})")
    try:
        img = Image.open(io.BytesIO(data))
        text = (pytesseract.image_to_string(img, lang=lang or "kor+eng") or "").strip()
        if not text:
            return OcrResult(False, engine="tesseract", error="글자를 찾지 못했습니다")
        return OcrResult(True, text=text, engine="tesseract")
    except pytesseract.TesseractNotFoundError:
        return OcrResult(False, engine="tesseract",
                         error="Tesseract 프로그램이 없습니다: winget install UB-Mannheim.TesseractOCR")
    except Exception as e:
        return OcrResult(False, engine="tesseract", error=str(e))


# ─────────────────────────────────────────────────────────────
# 3. Claude — 표·손글씨까지 가장 잘 읽지만, 이미지가 외부로 전송되고 비용이 든다
# ─────────────────────────────────────────────────────────────
#   pip install anthropic
#   인증: 환경변수 ANTHROPIC_API_KEY, 또는 `ant auth login` 으로 저장한 프로필
#   ⚠ 개인정보가 담긴 사진·문서에는 쓰지 말 것 (내용이 밖으로 나간다)
_CLAUDE_MODEL = "claude-opus-5"
_CLAUDE_MAX_EDGE = 1568        # 이보다 큰 그림은 API 가 어차피 줄여서 본다
_CLAUDE_PROMPT = (
    "이 이미지에 있는 글을 마크다운으로 옮겨 적어라. "
    "표는 마크다운 표로, 제목은 # 으로 나타낸다. "
    "설명·인사말 없이 옮긴 내용만 출력한다. 글자가 없으면 아무것도 출력하지 않는다.")


def _prepare_image(data: bytes) -> tuple[bytes, str]:
    """긴 변을 1568px 로 줄이고 PNG/JPEG 로 맞춘다 → (바이트, media_type)"""
    from PIL import Image
    img = Image.open(io.BytesIO(data))
    img.load()
    if max(img.size) > _CLAUDE_MAX_EDGE:
        img.thumbnail((_CLAUDE_MAX_EDGE, _CLAUDE_MAX_EDGE))
    buf = io.BytesIO()
    if img.mode in ("RGBA", "LA", "P"):
        img.save(buf, "PNG")
        return buf.getvalue(), "image/png"
    img.convert("RGB").save(buf, "JPEG", quality=90)
    return buf.getvalue(), "image/jpeg"


def _engine_claude(data: bytes, ext: str, lang: str) -> OcrResult:
    try:
        import base64
        import anthropic
    except ImportError as e:
        return OcrResult(False, engine="claude", error=f"설치 필요: pip install anthropic ({e})")
    try:
        payload, media_type = _prepare_image(data)
    except Exception as e:
        return OcrResult(False, engine="claude", error=f"그림을 열지 못했습니다: {e}")

    try:
        client = anthropic.Anthropic()
        response = client.beta.messages.create(
            model=_CLAUDE_MODEL,
            max_tokens=16000,
            # 안전 분류기가 거절하면 서버가 알맞은 모델로 다시 시도한다
            betas=["server-side-fallback-2026-07-01"],
            fallbacks="default",
            messages=[{"role": "user", "content": [
                {"type": "image", "source": {"type": "base64", "media_type": media_type,
                                             "data": base64.standard_b64encode(payload).decode()}},
                {"type": "text", "text": _CLAUDE_PROMPT},
            ]}],
        )
    except anthropic.AuthenticationError:
        return OcrResult(False, engine="claude",
                         error="인증 실패: ANTHROPIC_API_KEY 를 확인하거나 `ant auth login` 을 하세요")
    except anthropic.RateLimitError:
        return OcrResult(False, engine="claude", error="요청이 너무 많습니다. 잠시 뒤 다시 실행하세요")
    except anthropic.APIStatusError as e:
        return OcrResult(False, engine="claude", error=f"API 오류 {e.status_code}: {e.message}")
    except anthropic.APIConnectionError:
        return OcrResult(False, engine="claude", error="인터넷 연결을 확인하세요")
    except Exception as e:                             # 인증 정보가 아예 없을 때 등
        return OcrResult(False, engine="claude", error=str(e))

    if response.stop_reason == "refusal":
        return OcrResult(False, engine="claude", error="모델이 이 이미지 처리를 거절했습니다")
    text = "".join(b.text for b in response.content if b.type == "text").strip()
    if not text:
        return OcrResult(False, engine="claude", error="글자를 찾지 못했습니다")
    return OcrResult(True, text=text, engine="claude")


# ─────────────────────────────────────────────────────────────
# 등록표
# ─────────────────────────────────────────────────────────────
ENGINES = {
    "없음": _engine_none,
    "tesseract": _engine_tesseract,
    "claude": _engine_claude,
}

_current = "없음"
_lang = "kor+eng"


def set_engine(name: str, lang: str = "kor+eng"):
    global _current, _lang
    if name not in ENGINES:
        raise ValueError(f"모르는 엔진: {name} (쓸 수 있는 것: {', '.join(ENGINES)})")
    _current, _lang = name, lang


def current_engine() -> str:
    return _current


def is_on() -> bool:
    return _current != "없음"


def available() -> dict[str, tuple[bool, str]]:
    """엔진별로 (지금 쓸 수 있는가, 설명)."""
    out = {"없음": (True, "기본값")}
    try:
        import pytesseract
        try:
            v = pytesseract.get_tesseract_version()
            out["tesseract"] = (True, f"Tesseract {v}")
        except Exception:
            out["tesseract"] = (False, "pytesseract 는 있으나 Tesseract 프로그램이 없음")
    except ImportError:
        out["tesseract"] = (False, "pip install pytesseract pillow")
    try:
        import anthropic  # noqa: F401
        out["claude"] = (True, "설치됨 · 인증은 실제로 쓸 때 확인")
    except ImportError:
        out["claude"] = (False, "pip install anthropic")
    return out


def run_ocr_bytes(data: bytes, ext: str = "") -> OcrResult:
    fn = ENGINES.get(_current, _engine_none)
    try:
        return fn(data, ext, _lang)
    except Exception as e:
        return OcrResult(False, engine=_current, error=f"예기치 못한 오류: {e}")


def run_ocr(path: str) -> OcrResult:
    """파일 하나의 글자를 읽는다. 꺼져 있으면 파일을 읽지도 않는다."""
    if not is_on():
        return _engine_none(b"", "", _lang)
    try:
        with open(long_path(path), "rb") as f:
            data = f.read()
    except OSError as e:
        return OcrResult(False, engine=_current, error=str(e))
    return run_ocr_bytes(data, os.path.splitext(path)[1])
