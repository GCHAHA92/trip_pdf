# core/pdf_analyzer.py
"""
디버그용 pdf_analyzer:
- PDF 내용은 무시하고
- 템플릿 엑셀의 C5, H5에 TEST 값을 써 넣어
  실제로 엑셀에 쓰기가 되는지 확인하는 용도
"""

from __future__ import annotations
from typing import Tuple
import io

import pandas as pd
from openpyxl import load_workbook


def analyze_pdf_and_template(
    pdf_bytes: bytes,
    template_bytes: bytes,
) -> Tuple[pd.DataFrame, bytes]:
    """
    PDF 내용은 무시하고,
    템플릿 엑셀의 C5 셀에 '테스트', H5 셀에 12345를 써 넣은 뒤
    수정된 엑셀 파일을 bytes로 반환한다.
    """

    # 1) 요약 DataFrame은 그냥 빈 DataFrame 리턴 (UI 깨지지 않게 컬럼 하나 넣어줌)
    summary = pd.DataFrame({"메시지": ["디버그용 요약 (PDF 내용은 무시함)"]})

    # 2) 템플릿 엑셀 로드
    wb = load_workbook(io.BytesIO(template_bytes))
    ws = wb.active

    # 3) 테스트로 C5, H5에 값 써 넣기
    ws["C5"] = "테스트"
    ws["H5"] = 12345

    # 4) 메모리 상에 저장해서 bytes로 반환
    out = io.BytesIO()
    wb.save(out)
    out.seek(0)
    result_bytes = out.getvalue()

    return summary, result_bytes