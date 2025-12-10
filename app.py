import streamlit as st
import pandas as pd

from core.pdf_analyzer import analyze_pdf_and_template

TEMPLATE_PATH = "templates/지급조서_템플릿.xlsx"

st.set_page_config(
    page_title="출장비 자동정산기 (PDF)",
    layout="centered",
)

st.title("📄 출장비 자동정산기 (PDF 버전)")
st.write(
    """
인사랑에서 출력한 **'출장 월별집계 PDF'**와  
깃허브에 포함된 **지급조서 템플릿 엑셀**을 이용해,

규칙에 따라 실제 지급해야 할 금액을 다시 계산하고,  
PDF 금액과 차이가 있는 경우 **지급조서에서 차이를 표시**합니다.
"""
)

st.markdown("---")

# 1) PDF 업로드
uploaded_pdf = st.file_uploader("1. 출장 월별집계 PDF 업로드", type=["pdf"])

run_button = st.button("정산 실행")

if run_button:
    if not uploaded_pdf:
        st.error("먼저 '출장 월별집계 PDF' 파일을 업로드해 주세요.")
    else:
        with st.spinner("PDF 분석 및 지급조서 작성 중..."):
            try:
                pdf_bytes = uploaded_pdf.read()

                # 템플릿 엑셀은 깃허브 repo 안에 있는 파일을 그대로 사용
                with open(TEMPLATE_PATH, "rb") as f:
                    template_bytes = f.read()

            except FileNotFoundError:
                st.error(f"템플릿 파일을 찾을 수 없습니다: {TEMPLATE_PATH}")
            except Exception as e:
                st.error(f"파일을 읽는 중 오류가 발생했습니다: {e}")
            else:
                try:
                    # 핵심 로직: PDF + 템플릿 → (summary_df, 결과엑셀 bytes)
                    summary_df, result_bytes = analyze_pdf_and_template(
                        pdf_bytes,
                        template_bytes,
                    )
                except Exception as e:
                    st.error(f"처리 중 오류가 발생했습니다: {e}")
                else:
                    st.success("정산 완료!")

                    st.subheader("요약 결과")

                    # 👉 여기서 '차이' 컬럼이 있을 때만 정렬/차이표 보여주기
                    if isinstance(summary_df, pd.DataFrame):
                        if "차이" in summary_df.columns:
                            # 차이 기준으로 정렬
                            summary_display = summary_df.sort_values("차이", ascending=False)
                            st.dataframe(summary_display, use_container_width=True)

                            # 차이 나는 사람만 따로
                            diff_df = summary_display[summary_display["차이"] != 0]
                            if not diff_df.empty:
                                st.subheader("PDF 금액과 계산 금액이 다른 대상자 목록")
                                st.dataframe(diff_df, use_container_width=True)
                            else:
                                st.info("PDF 금액과 규칙 계산 금액이 모두 일치합니다. 🎉")
                        else:
                            # 디버그용처럼 '차이'가 없는 경우 그냥 전체 출력
                            st.dataframe(summary_df, use_container_width=True)
                    else:
                        st.write(summary_df)

                    st.markdown("---")

                    # 3) 지급조서 엑셀 다운로드
                    st.download_button(
                        "📥 지급조서 엑셀 다운로드",
                        data=result_bytes,
                        file_name="지급조서_from_pdf.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    )