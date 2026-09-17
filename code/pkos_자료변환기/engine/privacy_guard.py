# -*- coding: utf-8 -*-
"""
개인정보 필터 보정 — 대학원·수업 자료에서 드러난 문제를 덧씌워 고친다
======================================================================

privacy.py 는 학교 업무 문서로 다듬은 것이라 그 자료에는 잘 맞는다. 논문과 강의
자료를 넣으니 새 문제가 나왔다. privacy.py 는 pkems_pdf_변환 에서 그대로 가져온
파일이라 **건드리지 않고**, 여기서 find() 를 덧씌운다.

1. 문헌 번호를 카드번호로 오인 (오탐)
       https://doi.org/10.1177/1365480216659733   → ****-****-****-****
   숫자가 **들어 있는 덩어리**가 주소·DOI 이거나, 바로 앞에 ISBN·DOI 표시가
   붙어 있을 때만 거른다. (처음 판은 '앞뒤 70자 안에 주소가 있으면' 걸렀는데,
   그러면 '결제 안내 https://… 카드번호 4111-…' 의 진짜 카드번호까지 놓쳤다)

2. 계좌번호를 놓침 (누락 · 더 위험)
       환급 계좌는 국민은행 123456-78-901234
       계좌: 신한 110-123-456789
       계좌번호(경남은행) 207-0012-3456-78
   원본 규칙은 '계좌' 표시 바로 뒤의 숫자만 인정한다. 사이에 은행 이름 하나가
   끼는 것까지 허용한다. 날짜 모양은 원본처럼 거른다.

3. 흔한 낱말을 사람 이름으로 오인 (오탐)
   원본은 '이름·작성자·담당자' 같은 표시 뒤 2~3글자를 이름으로 확정하고, **문서
   전체에서 같은 낱말을 모두** 가린다. 실제 수업자료 낱말 2,747개를 시험하니
   445개가 그렇게 오인됐다(모형·설계·지식·이해·원리…).
       파일 이름 설정 방법 … 환경 설정에서     → '설*' 이 문서 곳곳에

   낱말 목록으로는 끝이 없어서 **한국어 문법을 근거로** 거른다. 같은 낱말이 문서
   어딘가에서 사람 이름에는 붙지 않는 꼬리(-하다 -적 -화 -된 에서 …)를 달고
   나오면 사람 이름이 아니다. 근거가 없으면 원본대로 가린다(놓치는 쪽이 위험하다).
   꼬리 없이 쓰이는 기술 용어(마우스·노트북…)만 목록으로 막는다.

실제 자료 215개(대학원 폴더 전체)로 돌려 보고 더 고친 것

4. 표 속 숫자 나열을 카드번호로 오인      '2016 2017 2018 2019'
   앞뒤로 숫자가 더 이어지면 표로 본다. '카드' 표시가 가까이 없으면 카드번호
   검증식(Luhn)을 통과해야 가린다 — 아무 16자리 문서 번호는 90% 가 걸러진다.
5. 주소(URL) 속 날짜를 생년월일로 오인     '…?id=3&date=2019-03-15&…'
6. 인용 속 저자 이름을 가려 인용이 망가짐   '이상수(2024)' → '이*수(2024)'
   문서에 '이름(연도)' '(이름, 연도)' 꼴로 나오는 이름은 공개된 저자로 보고
   가리지 않는다. '홍길동(2023년 입학)' 은 해당하지 않는다.

전화번호·이메일은 건드리지 않는다. 주민등록번호는 DOI·ISBN 의 일부일 때만 거른다.

PKOS(개인지식운영체계) 프로젝트
"""

from __future__ import annotations

import re

from .privacy import PrivacyFilter, Hit, mask_account, _looks_like_date


# ─────────────────────────────────────────────────────────────
# 1. 문헌 번호
# ─────────────────────────────────────────────────────────────
_REF_IN_TOKEN = re.compile(r"doi|https?://|www\.|arxiv|pmid|pmc\d|isbn|issn|10\.\d{4,9}/", re.I)
_REF_LABEL = re.compile(r"(?:doi|isbn|issn|arxiv|pmid|pmcid)\s*[:：]?\s*$", re.I)
# Springer 형식 DOI '10.1023/A:1022193728205' 가 PDF 에서 '10.1023/ A:  1022…' 로 끊겨 나온다
_DOI_PREFIX = re.compile(r"10\.\s?\d{4,9}\s?/\s?(?:[A-Za-z]{1,3}\s?:)?\s*$")


def _token(text: str, start: int, end: int) -> str:
    s = start
    while s > 0 and not text[s - 1].isspace():
        s -= 1
    e = end
    while e < len(text) and not text[e].isspace():
        e += 1
    return text[s:e]


def is_reference_number(text: str, start: int, end: int) -> bool:
    """[start, end) 의 숫자가 주소·DOI·ISBN 의 일부인가"""
    if _REF_IN_TOKEN.search(_token(text, start, end)):
        return True
    if start > 0 and text[start - 1] == "/":
        return True             # PDF 에서 'doi.org/10. 1177/1365…' 처럼 끊겨 나온 주소의 뒷부분
    # PDF 는 긴 DOI 를 줄에서 자주 끊는다 ('…/10.1007/A:' ↵ '1022193728205').
    # 표시가 번호 바로 앞(공백·줄바꿈만 사이)에 있어야 하는 조건은 정규식의 $ 가 지킨다.
    before = text[max(0, start - 30):start].replace("\n", " ")
    return bool(_REF_LABEL.search(before) or _DOI_PREFIX.search(before))


# ─────────────────────────────────────────────────────────────
# 카드번호 — 표 속 숫자 나열과 문서 번호를 거른다
# ─────────────────────────────────────────────────────────────
_CARD_LABEL = re.compile(r"(?:카드|card|결제|신용|체크)", re.I)


def _luhn_ok(digits: str) -> bool:
    """카드번호 검증식. 실제 카드번호는 모두 통과하고, 아무 16자리는 10% 만 통과한다."""
    total = 0
    for i, ch in enumerate(reversed(digits)):
        d = int(ch)
        if i % 2:
            d = d * 2 - 9 if d > 4 else d * 2
        total += d
    return total % 10 == 0


def is_not_card(text: str, start: int, end: int) -> bool:
    val = text[start:end]
    groups = re.findall(r"\d+", val)
    # ① '2016 2017 2018 2019' 처럼 공백으로 이어진 숫자 나열 — 앞뒤로 숫자가 더 붙어 있다
    if " " in val and (re.search(r"\d\s+$", text[max(0, start - 6):start])
                       or re.match(r"\s+\d", text[end:end + 6])):
        return True
    if len(groups) == 4 and all(1900 <= int(g) <= 2099 for g in groups):
        return True
    # ② 가까이 '카드' 표시가 있으면 검증식과 상관없이 가린다 (오타 난 진짜 번호일 수 있다)
    if _CARD_LABEL.search(text[max(0, start - 20):start]):
        return False
    # ③ 표시가 없으면 검증식을 통과해야 카드번호로 본다
    return not _luhn_ok("".join(groups))


# ─────────────────────────────────────────────────────────────
# 날짜 — 주소(URL) 속 날짜는 생년월일이 아니다
# ─────────────────────────────────────────────────────────────
def is_url_date(text: str, start: int, end: int) -> bool:
    tok = _token(text, start, end)
    if re.search(r"[=&?]|://|www\.", tok):
        return True
    return start > 0 and text[start - 1] == "/" and text[end:end + 1] == "/"


# ─────────────────────────────────────────────────────────────
# 인용 속 저자 — 논문 저자 이름은 공개 정보다
# ─────────────────────────────────────────────────────────────
def is_cited_author(name: str, text: str) -> bool:
    """문서 어딘가에 '이상수(2024)' '(이상수, 2024)' '이상수 외(2024)' 꼴로 나오는가.
    '홍길동(2023년 입학)' 은 해당하지 않는다 — 괄호 안이 연도 하나로 끝나야 한다."""
    n = re.escape(name)
    year = r"(?:19|20)\d{2}[a-z]?"
    return bool(re.search(rf"{n}\s*\(\s*{year}\s*[),]"
                          rf"|{n}\s*(?:,|외|등|et al\.?)\s*{year}\s*\)"
                          rf"|{n}\s*(?:외|등|et al\.?)\s*\(\s*{year}", text))


# ─────────────────────────────────────────────────────────────
# 2. 계좌번호 — 은행 이름이 끼어드는 실제 표현
# ─────────────────────────────────────────────────────────────
_BANK_SHORT = (r"국민|신한|우리|하나|농협|기업|산업|수협|신협|새마을|우체국|카카오|토스|"
               r"케이|씨티|SC제일|제일|부산|경남|대구|광주|전북|제주|IBK|KB|NH|BNK|DGB")
_BANK = (rf"(?:[가-힣]{{2,6}}(?:은행|저축은행|상호금융|증권|생명|화재)|[가-힣]{{2,4}}뱅크|"
         rf"{_BANK_SHORT})")
_ACCOUNT_WIDE = re.compile(
    r"(?:계좌\s*(?:번호)?|입금\s*계좌|송금\s*계좌|환불\s*계좌|예금\s*주?|account)"
    r"[은는이가을를]?\s*[:：]?\s*"
    rf"(?:[(\[]?\s*{_BANK}\s*[)\]]?\s*[:：]?\s*)?"
    r"(\d[\d-]{7,20}\d)", re.IGNORECASE)


# ─────────────────────────────────────────────────────────────
# 3. 사람 이름이 아닌 낱말
# ─────────────────────────────────────────────────────────────
# 사람 이름 뒤에는 오지 않는 꼬리. 사람 이름에도 붙는 것은 모두 뺐다
#   에게·한테·께      사람에게 쓰는 조사
#   에                '이상수에 따르면'
#   으로·로           '담당자를 홍길동으로 변경' (공문에 흔하다)
#   하고·하면         '철수하고 영희', '이순신하면'
_NONPERSON_TAIL = (r"(?:에서|로써|"
                   r"하[는다며여였지기게]|했|한다|할|한(?!테)|해[서야도]|"
                   r"적(?:인|으로|이)?|화(?:된|하|를|가|는)?|성(?:이|을|은|의|과)?|"
                   r"된|되[어는고지며]|될|시키|스러|롭)")

# 꼬리 없이 쓰이는 기술·교육 용어 (시험으로 걸리는 것을 확인한 것만)
NOT_NAMES = {
    "마우스", "노트북", "모니터", "키보드", "스피커", "프린터", "유튜브", "인터넷",
    "파이썬", "브라우저", "서버실", "디스크", "메모리", "라디오", "위젯", "위젯들",
    "변수명", "함수명", "파일명", "폴더명", "모듈명", "문자열", "정수형", "실수형",
    "배열형", "조건문", "반복문", "조건식", "주석문", "오류값", "반환값", "최소값",
    "최대값", "초기값", "기본값", "구조체", "매개체", "저장소", "실행문", "선언문",
    "연산식", "라벨", "박스", "원본", "모둠", "모둠원", "기입란", "소집단",
    "강의실", "강의안", "지도자", "신청자", "제출물", "채점표", "문항수", "정답률",
    "오답률", "방향성", "심화반", "백업본", "안내문", "전환점", "최적화", "남은것",
    "참가자", "발표자", "학습자", "교수자", "대상자", "연구자", "평가지", "정답지",
    "배부물", "유인물", "설계자", "제공자", "진단자", "설명자", "전달자", "조정자",
}


def looks_like_common_word(word: str, text: str) -> bool:
    """이 낱말이 문서 안에서 사람 이름이 아닌 쓰임새로 나오는가"""
    w = word.replace(" ", "")
    if w in NOT_NAMES:
        return True
    return bool(re.search(rf"(?<![가-힣]){re.escape(w)}{_NONPERSON_TAIL}", text))


# ─────────────────────────────────────────────────────────────
# 마크다운에서 가리면 안 되는 자리
# ─────────────────────────────────────────────────────────────
# 링크 주소 칸 · 맨 주소 · HTML 주석(쪽 표시). 여기 든 숫자·낱말을 가리면 링크가 깨진다.
_PROTECT = re.compile(r"\]\([^)\n]*\)|https?://[^\s)>\]]+|<!--.*?-->", re.S)


class AcademicPrivacyFilter(PrivacyFilter):
    """대학원·수업 자료에 맞춘 개인정보 필터."""

    def __init__(self, policy=None):
        super().__init__(policy)
        self.skipped: list[Hit] = []           # 걸러낸 것 (확인용)

    def find(self, text: str) -> list[Hit]:
        hits = super().find(text)
        not_person: dict[str, bool] = {}
        keep = []
        for h in hits:
            if self._false_positive(h, text, not_person):
                self.skipped.append(h)
            else:
                keep.append(h)
        keep += self._more_accounts(text, keep)
        keep.sort(key=lambda h: h.start)        # mask() 는 앞에서부터 이어 붙인다
        return keep

    @staticmethod
    def _false_positive(h: Hit, text: str, not_person: dict) -> bool:
        if h.kind in ("카드번호", "계좌번호", "주민등록번호") and \
                is_reference_number(text, h.start, h.end):
            return True
        if h.kind == "카드번호" and is_not_card(text, h.start, h.end):
            return True
        if h.kind == "생년월일" and is_url_date(text, h.start, h.end):
            return True
        if h.kind == "이름":
            w = h.original.replace(" ", "")
            if w not in not_person:
                not_person[w] = looks_like_common_word(w, text) or is_cited_author(w, text)
            return not_person[w]
        return False

    def _more_accounts(self, text: str, existing: list[Hit]) -> list[Hit]:
        taken = [(h.start, h.end) for h in existing]
        found = []
        for m in _ACCOUNT_WIDE.finditer(text):
            s, e = m.span(1)
            if any(s < be and e > bs for bs, be in taken):
                continue
            val = m.group(1)
            if _looks_like_date(val) or len(re.sub(r"\D", "", val)) < 9:
                continue
            if is_reference_number(text, s, e):
                continue
            found.append(Hit("계좌번호", val, mask_account(val, self.p.계좌번호), s, e,
                             text[max(0, s - 18):e + 18].replace("\n", " ").strip(), "보통"))
            taken.append((s, e))
        return found

    def mask_markdown(self, text: str) -> tuple[str, list[Hit]]:
        """마크다운 본문을 가린다. 링크 주소·맨 주소·주석 안은 건드리지 않는다."""
        protected = [(m.start(), m.end()) for m in _PROTECT.finditer(text)]
        hits = [h for h in self.find(text)
                if not any(h.start < e and h.end > s for s, e in protected)]
        keep = [h for h in hits if self.p.mode_for(h.kind) != "그대로"]
        out, last = [], 0
        for h in keep:
            if h.start < last:
                continue
            out.append(text[last:h.start])
            out.append(h.masked)
            last = h.end
        out.append(text[last:])
        return "".join(out), keep
