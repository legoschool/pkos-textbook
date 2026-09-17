# -*- coding: utf-8 -*-
"""
변환.py — 한글 이름으로 부르고 싶을 때 쓰는 문 (알맹이는 pkos.py 에 있다)

    python 변환.py "D:\\흩어진자료" --훑기
    python pkos.py "D:\\흩어진자료" --훑기      ← 똑같이 동작한다

두 이름을 함께 두는 이유 — 윈도우 배치 파일(.cmd)에서 한글 파일 이름을
부르면 콘솔 코드페이지에 따라 '파일을 찾을 수 없다'가 난다. 그래서
설치.cmd·변환하기.cmd 는 영문 이름인 pkos.py 를 부르고, 사람이 직접 칠 때는
읽기 쉬운 이 이름을 쓴다.
"""

import sys
import os

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

from pkos import main  # noqa: E402

if __name__ == "__main__":
    try:
        sys.exit(main())
    except KeyboardInterrupt:
        print("\n중단했습니다. 다시 실행하면 이어서 진행합니다.")
        sys.exit(130)
