# core/pdf_analyzer.py
"""
PDF 파서 + 계산 규칙을 이용해
성명별 집계를 만들고, 지급조서 템플릿(현재 구조 고정)에 결과를 채워넣는 모듈.

템플릿 전제 (현재 사용 중인 양식 기준):
- C4: '성명' 헤더
- C5부터 실제 데이터 입력 (사람 이름)
- A열: 연번, C열: 성명, D열: 4시간 미만, E열: 4시간 이상
- F열: 관용차량/차량사용횟수, H열: PDF 기준 지급액(원)
- L열: 규칙에 따른 계산결과(올바른 지급액)를 표시하는 용도

추가 규칙:
- 계산 결과 사람 수가 21명 이하라면:
    - C5 ~ C(5 + n - 1)에 사람들 이름 채우고
    - 합계는 항상 H26 셀에 둔다.
- 계산 결과 사람 수가 21명을 초과하면:
    - 원래 합계가 있던 26행 위에 필요한 행 수만큼 삽입해서
    - "마지막 사람 바로 아래 행"의 H열에 합계를 둔다.
"""

from __future__ import annotations
from typing import Tuple
import io

import pandas as pd
from openpyxl import load_workbook
from openpyxl.styles import Font

from .pdf_parser import parse_pdf_to_rows
from .rules import compute_allowance_by_person


# ===== 템플릿 고정 위치 =====
HEADER_ROW = 4       # C4 = 성명 헤더
DATA_START_ROW = 5   # C5부터 데이터
COL_NO = 1           # A열 연번
COL_NAME = 3         # C열 성명
COL_UNDER4 = 4       # D열 4시간 미만
COL_OVER4 = 5        # E열 4시간 이상
COL_CAR = 6          # F열 차량사용횟수
COL_PDF_AMT = 8      # H열 PDF 기준 지급액(원)
COL_CALC_AMT = 12    # L열: 계산결과(규칙상 올바른 금액) 표시용


def summarize_pdf_by_person(df_rows: pd.DataFrame) -> pd.DataFrame:
    """
    행단위 df_rows (parse_pdf_to_rows 결과)를 받아
    PDF 기준 성명별 집계 DataFrame 생성.

    반환:
        index: 성명
        columns:
          - 총지급액_숫자
          - 4시간미만
          - 4시간이상
          - 차량사용횟수
    """
    df = df_rows.copy()

    # 분 단위 출장 시간 기준으로 4시간 미만/이상 플래그
    df["is_under4"] = (df["minutes"] > 0) & (df["minutes"] < 240)
    df["is_over4"] = df["minutes"] >= 240
    df["car_cnt"] = df["car_used"].astype(int)

    grouped = df.groupby("성명").agg(
        총지급액_숫자=("amount_pdf", "sum"),
        _4시간미만=("is_under4", "sum"),
        _4시간이상=("is_over4", "sum"),
        차량사용횟수=("car_cnt", "sum"),
    )

    grouped = grouped.rename(columns={"_4시간미만": "4시간미만", "_4시간이상": "4시간이상"})
    for c in ["총지급액_숫자", "4시간미만", "4시간이상", "차량사용횟수"]:
        grouped[c] = grouped[c].fillna(0).astype(int)

    return grouped


def fill_template_with_summary(template_bytes: bytes, summary: pd.DataFrame) -> bytes:
    """
    지급조서 템플릿(엑셀 바이너리)에 summary 정보를 채워넣고,
    결과 엑셀을 bytes로 반환.

    템플릿 전제:
      - C4: '성명' 헤더
      - C5부터 데이터 입력
      - 기본 21명(= C5~C25)까지 들어가는 양식이고,
        이 경우 합계는 항상 H26 셀에 둔다.
      - 21명을 초과하면 H26 위에 줄을 추가로 삽입해서,
        마지막 사람 바로 아래 H열에 합계를 둔다.
    """
    # summary: index = 성명
    summary = summary.copy()
    summary = summary.fillna(0)

    # 타입 보정
    for col in ["총지급액_숫자", "4시간미만", "4시간이상", "차량사용횟수", "계산_총지급액", "차이"]:
        if col in summary.columns:
            summary[col] = summary[col].astype(int)

    wb = load_workbook(io.BytesIO(template_bytes))
    ws = wb.active

    red_bold_font = Font(color="FF0000", bold=True)

    # 성명 목록 정렬(보기 좋게)
    people = list(summary.sort_index().iterrows())
    n_people = len(people)

    # 사람 없으면 템플릿만 그대로 반환
    if n_people == 0:
        out_empty = io.BytesIO()
        wb.save(out_empty)
        out_empty.seek(0)
        return out_empty.getvalue()

    # ────────── 합계 위치 설계 ──────────
    # DATA_START_ROW = 5 (C5부터 이름)
    # 기본 설계: 최대 21명 -> 데이터는 5~25, 합계는 26행(H26)
    MAX_ROWS_WITHOUT_INSERT = 21
    BASE_SUM_ROW = DATA_START_ROW + MAX_ROWS_WITHOUT_INSERT  # 5 + 21 = 26

    # 인원이 21명 초과면, BASE_SUM_ROW(26행) 위에 줄 삽입
    if n_people > MAX_ROWS_WITHOUT_INSERT:
        extra_rows = n_people - MAX_ROWS_WITHOUT_INSERT
        ws.insert_rows(BASE_SUM_ROW, amount=extra_rows)
        # 이제 합계 행은 26 + extra_rows로 내려감
        total_row = BASE_SUM_ROW + extra_rows
    else:
        # 21명 이하이면 합계는 항상 고정 H26에 둔다.
        total_row = BASE_SUM_ROW

    # ────────── 데이터 채우기 ──────────
    current_row = DATA_START_ROW
    for idx, (name, row) in enumerate(people, start=1):
        name_str = str(name).strip()
        if not name_str:
            continue

        under4 = int(row.get("4시간미만", 0) or 0)
        over4 = int(row.get("4시간이상", 0) or 0)
        car_cnt = int(row.get("차량사용횟수", 0) or 0)
        amount_pdf = int(row.get("총지급액_숫자", 0) or 0)
        amount_calc = int(row.get("계산_총지급액", 0) or 0)

        # 연번(A열)
        ws.cell(row=current_row, column=COL_NO).value = idx

        # 성명(C열)
        ws.cell(row=current_row, column=COL_NAME).value = name_str

        # 4시간 미만/이상/차량
        ws.cell(row=current_row, column=COL_UNDER4).value = "" if under4 == 0 else under4
        ws.cell(row=current_row, column=COL_OVER4).value = "" if over4 == 0 else over4
        ws.cell(row=current_row, column=COL_CAR).value = "" if car_cnt == 0 else car_cnt

        # PDF 금액(H열)
        cell_h = ws.cell(row=current_row, column=COL_PDF_AMT)
        cell_h.value = "" if amount_pdf == 0 else amount_pdf

        # 계산 금액(L열) - PDF와 다를 때만
        cell_l = ws.cell(row=current_row, column=COL_CALC_AMT)
        if amount_pdf != amount_calc and amount_calc != 0:
            cell_l.value = amount_calc
            cell_h.font = red_bold_font
        else:
            cell_l.value = ""

        current_row += 1

    last_data_row = current_row - 1  # 마지막 사람 행

    # ────────── 합계 수식(H열) 설정 ──────────
    # 합계는 total_row 행의 H열:
    #   =SUM(H5:H{last_data_row})
    if last_data_row >= DATA_START_ROW:
        first_cell = ws.cell(row=DATA_START_ROW, column=COL_PDF_AMT).coordinate
        last_cell = ws.cell(row=last_data_row, column=COL_PDF_AMT).coordinate
        ws.cell(row=total_row, column=COL_PDF_AMT).value = f"=SUM({first_cell}:{last_cell})"

    out = io.BytesIO()
    wb.save(out)
    out.seek(0)
    return out.getvalue()

def analyze_pdf_and_template(
    pdf_bytes: bytes,
    template_bytes: bytes,
) -> Tuple[pd.DataFrame, bytes]:
    """
    PDF + 템플릿을 받아서:
      - 성명별 요약 DataFrame(summary)
      - 템플릿에 결과 채운 엑셀 bytes
    를 반환.
    """
    # 1) PDF 파싱 (행 단위 데이터)
    df_rows = parse_pdf_to_rows(pdf_bytes)

    # 2) PDF 기준 집계
    pdf_summary = summarize_pdf_by_person(df_rows)

    # 3) 규칙에 따른 실제 지급액 계산
    calc_totals = compute_allowance_by_person(df_rows)

    # 4) summary 결합
    summary = pdf_summary.copy()
    summary["계산_총지급액"] = calc_totals
    summary["계산_총지급액"] = summary["계산_총지급액"].fillna(0).astype(int)
    summary["차이"] = summary["계산_총지급액"] - summary["총지급액_숫자"]

    # 5) 템플릿 채우기
    result_bytes = fill_template_with_summary(template_bytes, summary)

    return summary, result_bytes