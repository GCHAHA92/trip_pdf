# core/pdf_analyzer.py
"""
PDF 파서 + 계산 규칙을 이용해
성명별 집계를 만들고, 지급조서 템플릿 엑셀에 결과를 채워넣는 모듈.
"""

from __future__ import annotations
from typing import Tuple
import io

import pandas as pd
from openpyxl import load_workbook
from openpyxl.styles import Font

from .pdf_parser import parse_pdf_to_rows
from .rules import compute_allowance_by_person


# 템플릿에서 데이터가 시작하는 기본 행 (헤더 바로 아래)
# 예: 4행에 "연번 / 직급 / 성명 / …" 헤더가 있고, 5행부터 데이터면 START_ROW=5
START_ROW_DEFAULT = 5


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


def _find_header_and_name_col(ws) -> tuple[int, int]:
    """
    워크시트에서 '성명'이 들어있는 헤더 셀 위치(행, 열)를 찾는다.
    - 대략 1~20행 안에서만 검색
    반환: (header_row, name_col_idx)
    """
    for row in range(1, 21):
        for col in range(1, 40):
            v = ws.cell(row=row, column=col).value
            if v is None:
                continue
            if str(v).strip() == "성명":
                return row, col
    # 못 찾으면 기본값(4행 C열) 가정
    return 4, 3


def fill_template_with_summary(template_bytes: bytes, summary: pd.DataFrame) -> bytes:
    """
    지급조서 템플릿(엑셀 바이너리)에 summary 정보를 채워넣고,
    결과 엑셀을 bytes로 반환.

    기본 가정(필요시 여기서만 수정하면 됨):
      - '성명' 열 위치는 헤더를 검색해서 자동으로 찾음.
      - D열: 4시간 미만 횟수
      - E열: 4시간 이상 횟수
      - F열: 차량 사용 횟수
      - H열: PDF 기준 지급액
      - L열: 규칙에 따른 계산금액 (PDF 금액과 다를 때만 표기)
      - H열과 L열 값이 다르면 H열은 빨간색/굵게 표시
    """
    summary = summary.copy()
    summary = summary.fillna(0)
    # 타입 보정
    for col in ["총지급액_숫자", "4시간미만", "4시간이상", "차이", "차량사용횟수", "계산_총지급액"]:
        if col in summary.columns:
            summary[col] = summary[col].astype(int)

    # 이름 매칭을 좀 더 튼튼하게 하기 위해 공백 제거 버전도 만든다.
    # key: 공백 제거한 이름, value: 원래 index 이름
    name_map = {}
    for idx_name in summary.index:
        key = str(idx_name).replace(" ", "").strip()
        if key:  # 비어있지 않은 경우만
            name_map[key] = idx_name

    wb = load_workbook(io.BytesIO(template_bytes))
    ws = wb.active

    # 1) 헤더 위치 & 성명 열 위치 찾기
    header_row, name_col = _find_header_and_name_col(ws)
    data_start_row = header_row + 1

    # 2) L열 헤더 없으면 추가
    if not ws[f"L{header_row}"].value:
        ws[f"L{header_row}"] = "계산금액(규칙)"

    # 빨간 글씨용 폰트
    red_bold_font = Font(color="FF0000", bold=True)

    row = data_start_row
    last_filled_row = header_row  # 나중에 합계용으로 사용

    while True:
        cell_name = ws.cell(row=row, column=name_col)
        name_val = cell_name.value

        # 이름 셀이 비어 있거나 '합계' 등이면 데이터 끝으로 본다
        if not name_val or str(name_val).strip() in ["합계"]:
            break

        # 템플릿에 적힌 이름 정리
        name_str_raw = str(name_val)
        name_str = name_str_raw.strip()
        name_key = name_str.replace(" ", "")

        # 기본값
        under4 = over4 = car_cnt = 0
        amount_pdf = amount_calc = 0

        # summary에서 이름 매칭
        idx_name = None
        if name_str in summary.index:
            idx_name = name_str
        elif name_key in name_map:
            idx_name = name_map[name_key]

        if idx_name is not None:
            r = summary.loc[idx_name]
            under4 = int(r.get("4시간미만", 0) or 0)
            over4 = int(r.get("4시간이상", 0) or 0)
            car_cnt = int(r.get("차량사용횟수", 0) or 0)
            amount_pdf = int(r.get("총지급액_숫자", 0) or 0)
            amount_calc = int(r.get("계산_총지급액", 0) or 0)

        # D/E/F: 4시간 미만/이상/차량횟수
        ws[f"D{row}"] = "" if under4 == 0 else under4
        ws[f"E{row}"] = "" if over4 == 0 else over4
        ws[f"F{row}"] = "" if car_cnt == 0 else car_cnt

        # H열: PDF 기준 지급액
        cell_h = ws[f"H{row}"]
        cell_h.value = "" if amount_pdf == 0 else amount_pdf

        # L열: 계산금액 (PDF 금액과 다를 때만 표시)
        cell_l = ws[f"L{row}"]
        if amount_pdf != amount_calc:
            # 계산 금액이 0인데 PDF 금액도 0이면 굳이 표시할 필요 없음
            if amount_calc != 0:
                cell_l.value = amount_calc
                cell_h.font = red_bold_font
            else:
                cell_l.value = ""
        else:
            cell_l.value = ""

        last_filled_row = row
        row += 1

    # 3) 합계 행(H열) 자동 계산 (필요시 사용)
    if last_filled_row >= data_start_row:
        total_row = last_filled_row + 1
        if not ws[f"G{total_row}"].value:
            ws[f"G{total_row}"] = "합계"
        ws[f"H{total_row}"] = f"=SUM(H{data_start_row}:H{last_filled_row})"

    # bytes로 저장
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