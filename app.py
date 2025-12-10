import streamlit as st
import pandas as pd
from io import BytesIO

from core.pdf_parser import parse_trip_pdf
from core.pdf_analyzer import analyze_pdf
from core.rules import compute_amount_for_rows

# 템플릿 경로
TEMPLATE_PATH = "templates/지급조서_템플릿.xlsx"

st.set_page_config(page_title="출장비 자동정산 시스템", layout="wide")

st.title("📄 출장 월별집계 PDF 기반 지급조서 자동 생성기")
st.write("PDF를 업로드하면 파싱 → 계산 → 지급조서를 자동으로 생성합니다.")

uploaded_pdf = st.file_uploader("출장 월별집계 PDF 업로드", type=["pdf"])

# -----------------------------------------
# PDF 업로드 처리
# -----------------------------------------
if uploaded_pdf is not None:
    st.info("PDF 파싱 중… 잠시만 기다려 주세요.")
    try:
        df_pdf = parse_trip_pdf(uploaded_pdf)
        st.success("PDF 파싱 완료!")
        st.dataframe(df_pdf, use_container_width=True)
    except Exception as e:
        st.error(f"PDF 파싱 중 오류 발생: {e}")
        st.stop()

    # 출장비 규칙 적용
    st.info("출장비 계산 중…")
    try:
        df_result = analyze_pdf(df_pdf)
        st.success("출장비 계산 완료!")
        st.dataframe(df_result, use_container_width=True)
    except Exception as e:
        st.error(f"출장비 계산 중 오류 발생: {e}")
        st.stop()

    # 지급조서 템플릿 불러오기
    try:
        template_df = pd.read_excel(TEMPLATE_PATH)
    except Exception as e:
        st.error(f"템플릿 파일을 읽는 중 오류 발생: {e}")
        st.stop()

    # 템플릿에 결과 매핑
    st.info("지급조서 생성 중...")

    # 템플릿의 이름과 계산된 df_result의 이름 매칭
    merged = template_df.copy()

    if "성명" not in merged.columns:
        st.error("템플릿에 '성명' 열이 없습니다.")
        st.stop()

    # L열 = 실제 계산 금액 / 차이가 있을 때만 표시
    excel_output = merged.merge(
        df_result[["성명", "총지급액_숫자", "올바른지급액", "차이"]],
        on="성명",
        how="left"
    )

    # 차이가 있는 경우만 L열에 표시
    excel_output["L열_계산금액"] = excel_output.apply(
        lambda r: r["올바른지급액"] if pd.notna(r["차이"]) and r["차이"] != 0 else "",
        axis=1
    )

    # 엑셀 다운로드용 버퍼 생성
    output = BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        excel_output.to_excel(writer, index=False, sheet_name="지급조서")

    st.success("🎉 지급조서 생성 완료!")

    st.download_button(
        label="📥 지급조서 Excel 다운로드",
        data=output.getvalue(),
        file_name="지급조서_from_pdf.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
