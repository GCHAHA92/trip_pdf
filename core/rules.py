# core/rules.py
"""
출장비 계산 규칙 모듈

규칙:
  - 1건 기준:
      minutes < 60          -> 0원
      60 <= minutes < 240   -> 10,000원
      minutes >= 240        -> 20,000원
      공용차량(car_used=True)이면서 금액 > 0이면 10,000원 차감 (최소 0원)
  - 1일(성명 + 시작일자) 최대 20,000원
  - 1달(성명 기준) 최대 280,000원
"""

from __future__ import annotations
import pandas as pd


def _normalize_minutes(minutes: int) -> int:
    """minutes 값이 None/빈 문자열/음수여도 0 이상 int로 정규화."""
    try:
        m = int(float(minutes))
    except Exception:
        m = 0
    return max(m, 0)


def _normalize_car_used(car_used) -> bool:
    """공용차량 사용 여부를 문자열/숫자/불리언 입력에 관계없이 bool로 변환."""
    true_values = {"사용", "y", "yes", "true", "1", 1, True}
    false_values = {"미사용", "n", "no", "false", "0", 0, False}

    if car_used in true_values:
        return True
    if car_used in false_values:
        return False
    if isinstance(car_used, str):
        normalized = car_used.strip().lower()
        if normalized in true_values:
            return True
        if normalized in false_values:
            return False
    return bool(car_used)


def calc_row_amount(minutes: int, car_used: bool) -> int:
    """출장 1건(1행)에 대한 기본 금액 계산."""
    m = _normalize_minutes(minutes)
    amt = 0
    if 60 <= m < 240:
        amt = 10000
    elif m >= 240:
        amt = 20000

    if car_used and amt > 0:
        amt -= 10000

    if amt < 0:
        amt = 0
    return amt


def compute_allowance_by_person(df_rows: pd.DataFrame) -> pd.Series:
    """
    행단위 데이터(df_rows)를 받아,
    규칙(건별 -> 일별 -> 월별 상한)을 적용한 뒤
    성명별 최종 금액(계산_총지급액)을 반환.
    """
    df = df_rows.copy()

    required_cols = ["성명", "시작일자", "minutes", "car_used"]
    for col in required_cols:
        if col not in df.columns:
            raise ValueError(f"df_rows에 '{col}' 컬럼이 없습니다.")

    # 1) 건별 기본 금액
    df["row_base"] = df.apply(
        lambda r: calc_row_amount(
            _normalize_minutes(r["minutes"]), _normalize_car_used(r["car_used"])
        ),
        axis=1,
    )

    # 2) 1일 상한 적용 (성명 + 시작일자 기준 최대 2만)
    df["row_daily_limited"] = 0
    for (name, day), g in df.groupby(["성명", "시작일자"]):
        cum = 0
        for idx in g.index:
            base = int(df.at[idx, "row_base"])
            if base <= 0 or cum >= 20000:
                add = 0
            elif cum + base <= 20000:
                add = base
            else:
                add = 20000 - cum
            cum += add
            df.at[idx, "row_daily_limited"] = add

    # 3) 1달 상한 적용 (성명 기준 최대 28만)
    df["row_final"] = 0
    for name, g in df.groupby("성명"):
        cum = 0
        for idx in g.index:
            base = int(df.at[idx, "row_daily_limited"])
            if base <= 0 or cum >= 280000:
                add = 0
            elif cum + base <= 280000:
                add = base
            else:
                add = 280000 - cum
            cum += add
            df.at[idx, "row_final"] = add

    person_totals = df.groupby("성명")["row_final"].sum()
    person_totals = person_totals.fillna(0).astype(int)
    return person_totals
