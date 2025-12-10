# core/pdf_analyzer.py
"""
완전 최소 디버그 버전.

- PDF 내용도, 템플릿도 전부 무시하고
- 새 엑셀 파일을 하나 만들어서
  A1에 "HELLO", B2에 "WORLD"를 써서 반환한다.

이래도 다운로드한 엑셀에 HELLO / WORLD가 안 보이면,
app.py가 이 함수를 아예 안 쓰고 있다는 뜻이다.
"""

from __future__ import annotations
from typing import Tuple
import io

import pandas as pd
from openpyxl import Workbook


def analyze_pdf_and_template(
    pdf_bytes: bytes,
    template_bytes: bytes,
) -> Tuple[pd.DataFrame, bytes]:
    # 1) 요약 DataFrame: 그냥 테스트용 한 줄
    summary = pd.DataFrame({"메시지": ["디버그용 요약 (새 워크북 테스트)"]})

    # 2) 완전히 새 워크북 생성 (템플릿은 아예 안 씀)
    wb = Workbook()
    ws = wb.active
    ws.title = "TEST_SHEET"

    ws["A1"] = "HELLO"
    ws["B2"] = "WORLD"
    ws["C3"] = "이 파일이 보이면 pdf_analyzer는 정상 작동 중"

    # 3) 메모리에 저장해서 bytes 반환
    out = io.BytesIO()
    wb.save(out)
    out.seek(0)
    result_bytes = out.getvalue()

    return summary, result_bytes