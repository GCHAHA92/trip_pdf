import streamlit as st
import pandas as pd

from core.pdf_analyzer import analyze_pdf_and_template


st.set_page_config(
    page_title="출장비 자동정산기 (PDF 버전)",
    layout="centered",
)

st.title("출장비 자동정산기 (PDF 버전)")
st.write(
    """
인사랑에서 출력한 **'출장 월별집계 PDF'**와  
이미 성명/직급/계좌 등이 기입된 **지급조서 템플릿 엑셀**을 업로드하면,  

규칙에 따라 실제 지급해야 할 금액을 다시 계산하고,  
PDF 금액과 차이가 있는 경우 **지급조서에서 차이를 표시**해 줍니다.
"""
)

st.markdown("---")

uploaded_pdf = st.file_uploader("1. 출장 월별집계 PDF 업로드", type=["pdf"])
uploaded_template = st.file_uploader("2. 지급조서 템플릿 엑셀 업로드", type=["xlsx"])

run_button = st.button("정산 실행")

if run_button:
    if not uploaded_pdf or not uploaded_template:
        st.error("PDF와 템플릿 엑셀 파일을 모두 업로드해 주세요.")
    else:
        with st.spinner("PDF 분석 및 지급조서 작성 중..."):
            pdf_bytes = uploaded_pdf.read()
            template_bytes = uploaded_template.read()

            try:
                summary_df, result_bytes = analyze_pdf_and_template(
                    pdf_bytes,
                    template_bytes,
                )
            except Exception as e:
                st.error(f"처리 중 오류가 발생했습니다: {e}")
            else:
                st.success("정산 완료!")

                st.subheader("성명별 요약 (PDF vs 계산금액)")
                # 차이 큰 순서대로 정렬
                summary_display = summary_df.sort_values("차이", ascending=False)
                st.dataframe(summary_display)

                # 차이 있는 사람만 따로
                diff_df = summary_display[summary_display["차이"] != 0]
                if not diff_df.empty:
                    st.subheader("PDF 금액과 계산 금액이 다른 대상자 목록")
                    st.dataframe(diff_df)
                else:
                    st.info("PDF 금액과 규칙 계산 금액이 모두 일치합니다. 🎉")

                st.markdown("---")
                st.download_button(
                    "지급조서 엑셀 다운로드",
                    data=result_bytes,
                    file_name="지급조서_from_pdf.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                )
