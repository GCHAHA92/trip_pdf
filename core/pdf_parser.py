# core/pdf_parser.py
"""
출장 월별집계 PDF 파서 (최소 컬럼 버전)

- '출장_월별집계.pdf' 형태의 PDF를 읽어서
  "출장 1건 = 1행" DataFrame으로 변환한다.

반환 컬럼:
  - 성명
  - 시작일자
  - 시작시각
  - 종료일자
  - 종료시각
  - minutes        (총출장시간 분 단위, int)
  - car_used       (공무용차량 사용 여부: bool)
  - amount_pdf     (해당 출장 건의 금액, int)
"""

from __future__ import annotations
from pathlib import Path
from typing import List, Optional, Tuple, Union, BinaryIO
import re
import io

import pdfplumber
import pandas as pd


# ---------------- 헬퍼 함수 ----------------

def _find_col(df: pd.DataFrame, keywords: List[str]) -> Optional[str]:
    """
    열 이름들 중에서 지정한 키워드들이 모두 포함된 열을 찾아 열 이름을 반환.
    예: ["출장기간"] -> '출장기간', '출장기간_raw' 등
    """
    for col in df.columns:
        name = str(col)
        if all(k in name for k in keywords):
            return col
    return None


def _parse_minutes(text: str) -> int:
    """
    '1일 2시간 30분', '7시간21분', '59분', '4시간', '1일' 등을 총 '분'으로 변환.
    """
    if text is None:
        return 0
    s = str(text).strip().replace(" ", "")
    if not s:
        return 0

    days = hours = mins = 0

    m = re.search(r"(\d+)일", s)
    if m:
        days = int(m.group(1))

    m = re.search(r"(\d+)시간", s)
    if m:
        hours = int(m.group(1))

    m = re.search(r"(\d+)분", s)
    if m:
        mins = int(m.group(1))

    return days * 24 * 60 + hours * 60 + mins


def _parse_period(text: str) -> Tuple[str, str, str, str]:
    """
    '2025-11-14 09:00 ~ 2025-11-14 18:00'
    '2025-11-14 09:00\n~ 2025-11-14 18:00'
    과 같이 시작/종료 일시가 포함된 문자열에서
    (시작일자, 시작시각, 종료일자, 종료시각)을 반환한다.
    인식하지 못하면 ("", "", "", "")을 반환.
    """
    if text is None:
        return "", "", "", ""

    s = str(text).strip()
    if not s:
        return "", "", "", ""

    # 줄바꿈, 탭 등을 공백으로 통일하고 '~' 주변에 공백 부여
    s = s.replace("\r", " ").replace("\n", " ")
    s = s.replace("~", " ~ ")
    s = re.sub(r"\s+", " ", s)

    # (시작일자 시작시각) ~ (종료일자 종료시각)
    patterns = [
        r"(?P<start_date>\d{4}-\d{1,2}-\d{1,2})\s+"
        r"(?P<start_time>\d{1,2}:\d{2})\s*~\s*"
        r"(?P<end_date>\d{4}-\d{1,2}-\d{1,2})\s+"
        r"(?P<end_time>\d{1,2}:\d{2})",
        # 종료일자가 따로 없고, 같은 날짜로만 표기된 경우
        r"(?P<date>\d{4}-\d{1,2}-\d{1,2})\s+"
        r"(?P<start_time>\d{1,2}:\d{2}).*?(?P<end_time>\d{1,2}:\d{2})",
    ]

    for pattern in patterns:
        m = re.search(pattern, s)
        if not m:
            continue

        group_dict = m.groupdict()
        if "start_date" in group_dict:
            return (
                group_dict.get("start_date", ""),
                group_dict.get("start_time", ""),
                group_dict.get("end_date", ""),
                group_dict.get("end_time", ""),
            )

        # 동일한 날짜로만 표기된 경우
        date_only = group_dict.get("date", "")
        return (
            date_only,
            group_dict.get("start_time", ""),
            date_only,
            group_dict.get("end_time", ""),
        )

    return "", "", "", ""


def _extract_name(raw: str) -> str:
    """
    '정홍식\\n(A141714\\n7)' 같은 문자열에서 실제 이름(첫 줄 또는 괄호 앞)만 추출.
    """
    if raw is None:
        return ""
    s = str(raw).strip()
    s = s.replace("\r", "\n")
    if "\n" in s:
        s = s.split("\n")[0]
    if "(" in s:
        s = s.split("(", 1)[0]
    return s.strip()


def _to_int_safe(x: str) -> int:
    """
    금액 문자열(예: '10,000', '0', '', '0원')을 int로 변환.
    숫자가 없으면 0.
    """
    if x is None:
        return 0
    s = str(x).replace(",", "").replace(" ", "")
    if s in ["", "-", "0원"]:
        return 0
    if not re.search(r"\d", s):
        return 0
    try:
        return int(float(s))
    except Exception:
        return 0


# ---------------- 메인 파서 ----------------

def parse_pdf_to_rows(
    pdf_source: Union[bytes, BinaryIO, str, Path]
) -> pd.DataFrame:
    """
    '출장 월별집계' PDF를 읽어 계산용 최소 컬럼만 가진
    '출장 1건 = 행 1개' 형태의 DataFrame을 반환.

    반환 컬럼:
      - 성명
      - 시작일자
      - 시작시각
      - 종료일자
      - 종료시각
      - minutes        (총출장시간 분 단위)
      - car_used       (공무용차량 사용 여부: bool)
      - amount_pdf     (PDF의 합계 금액, int)
    """
    # pdfplumber가 받을 수 있는 형태로 정리
    if isinstance(pdf_source, (bytes, bytearray)):
        pdf_file = io.BytesIO(pdf_source)
    else:
        pdf_file = pdf_source  # 파일 경로나 파일 객체

    tables = []

    with pdfplumber.open(pdf_file) as pdf:
        for page in pdf.pages:
            extracted = page.extract_tables()
            for table in extracted:
                if not table:
                    continue
                header = table[0]
                if not header:
                    continue
                # '순번'으로 시작하는 표만 출장 데이터로 간주
                if str(header[0]).strip() != "순번":
                    continue
                rows = table[1:]
                df_page = pd.DataFrame(rows, columns=header)
                tables.append(df_page)

    if not tables:
        raise ValueError("PDF에서 '순번' 헤더를 가진 출장 표를 찾지 못했습니다.")

    df_all = pd.concat(tables, ignore_index=True)

    # 열 이름/값 정리
    df_all.columns = [str(c).strip() for c in df_all.columns]
    df_all = df_all.fillna("").astype(str)
    df_all = df_all.apply(lambda col: col.str.strip())

    # 필요한 컬럼 찾기
    col_name = _find_col(df_all, ["성명"])
    col_period = _find_col(df_all, ["출장기간"])
    col_dur = _find_col(df_all, ["총출장시간"])
    col_car = _find_col(df_all, ["공무용차량"])
    col_amount = _find_col(df_all, ["합계"]) or _find_col(df_all, ["청구"])

    if not (col_name and col_period and col_dur and col_car and col_amount):
        raise ValueError(
            f"필수 열(성명/출장기간/총출장시간/공무용차량/합계)을 찾지 못했습니다. 실제 열 목록: {list(df_all.columns)}"
        )

    # 유효 행 필터링: 성명/출장기간 비어 있는 행, '합계', '소계' 등 제외
    name_raw = df_all[col_name]
    period_raw = df_all[col_period]

    mask_valid = (
        name_raw.notna()
        & (name_raw.str.strip() != "")
        & period_raw.notna()
        & (period_raw.str.strip() != "")
        & (~name_raw.str.contains("합계"))
        & (~name_raw.str.contains("소계"))
    )

    df_use = df_all[mask_valid].copy()

    # 최소 컬럼 생성
    df_use["성명"] = df_use[col_name].apply(_extract_name)

    (
        df_use["시작일자"],
        df_use["시작시각"],
        df_use["종료일자"],
        df_use["종료시각"],
    ) = zip(*df_use[col_period].apply(_parse_period))

    df_use["minutes"] = df_use[col_dur].apply(_parse_minutes)
    df_use["car_used"] = df_use[col_car].eq("사용")
    df_use["amount_pdf"] = df_use[col_amount].apply(_to_int_safe)

    cols_final = [
        "성명",
        "시작일자",
        "시작시각",
        "종료일자",
        "종료시각",
        "minutes",
        "car_used",
        "amount_pdf",
    ]
    df_final = df_use[cols_final].copy()

    return df_final
