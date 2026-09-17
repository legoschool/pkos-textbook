# 자료변환기 — 흩어진 자료를 마크다운으로

한글·워드·PPT·엑셀·PDF·사진·압축파일이 뒤섞인 폴더를 넣으면, AI와 웹이 읽기 좋은 **마크다운(.md) 묶음과 목차**가
나오는 파이썬 도구입니다. 설치부터 변환·결과 읽기·코드로 직접 쓰기까지 따라 하는 **실습 교재**를 함께 공개합니다.

## 바로 가기

| 무엇 | 링크 |
|---|---|
| **실습 교재** — 15블록 · 실제 실행 화면 11장 | https://legoschool.github.io/pkos-textbook/ |
| **코드 받기** — `pkos_자료변환기.zip` (26개 파일, 159KB) | https://legoschool.github.io/pkos-textbook/downloads/pkos-converter.zip |
| **실습 자료** — `흩어진자료_예시.zip` (17개, 1.1MB) | https://legoschool.github.io/pkos-textbook/downloads/sample-data.zip |
| **소개 PPT** — 왜 만들었나 · 어떻게 쓰나 · 앞으로 (5쪽) | https://legoschool.github.io/pkos-textbook/downloads/pkos-intro.pptx |
| **모든 파일 한 번에** — 판 번호별 내려받기 | https://github.com/legoschool/pkos-textbook/releases |
| **코드 둘러보기** | [code/pkos_자료변환기](code/pkos_자료변환기) |

## 5분 만에 시작 (Windows)

1. 코드 zip 과 실습 자료 zip 을 받아 `C:\pkos_실습` 에 풉니다.
2. `pkos_자료변환기\설치.cmd` 를 더블클릭합니다. 파이썬이 없으면 설치를 물어보고, 라이브러리 8개를 깝니다.
3. 변환할 폴더를 `변환하기.cmd` 위로 끌어다 놓고 `y` 를 누릅니다.
4. 원본 옆 `_변환결과\INDEX.md` 에서 목차와 변환된 문서를 봅니다.

명령창에서는:

```bash
python pkos.py "C:\pkos_실습\흩어진자료" --훑기
python pkos.py "C:\pkos_실습\흩어진자료"
python pkos.py --자가진단
```

자가진단은 시험 자료를 만들어 126가지를 판정합니다. `통과 126 · 실패 0` 이면 준비 끝입니다.

## 무엇을 읽나

한글(`.hwp` `.hwpx`) · 워드 · 파워포인트 · 엑셀(옛 `.doc` `.ppt` `.xls` 포함) · 오픈도큐먼트 · PDF · 웹문서 · 자막 ·
사진 · `.zip`. **확장자는 믿지 않고 파일 내용으로 형식을 판별**합니다. 여러 번 돌리면 바뀐 파일만 새로 만들고,
이름·전화번호·계좌번호 같은 개인정보는 기본으로 가립니다. 자세한 것은 [code/pkos_자료변환기/README.md](code/pkos_자료변환기/README.md).

## 구글 사이트에 넣기

구글 사이트 편집 → **삽입** → **삽입(Embed)** → **URL로** → `https://legoschool.github.io/pkos-textbook/` → **전체 페이지** 또는 **삽입**.
사이트 안에서 zip 받기가 막히면 교재 오른쪽 위 **새 창 ↗** 에서 받습니다.

## 저장소 구성

```
index.html                 실습 교재 (GitHub Pages 첫 화면)
img/                       교재의 실행 화면 사진
downloads/
  pkos-converter.zip       도구 코드
  sample-data.zip          실습 자료 (지어낸 내용)
  pkos-intro.pptx          소개 발표자료
code/pkos_자료변환기/       도구 코드 원본 (zip 과 같은 내용)
LICENSE                    MIT
```

## 라이선스

MIT — 자유롭게 쓰고, 고치고, 나눌 수 있습니다. 실습 자료의 이름·전화번호·번호들은 모두 지어낸 것입니다.
개인정보 가리기는 완벽하지 않으니 공개 전에는 사람이 직접 확인하세요.

만든 이: 레고학교
