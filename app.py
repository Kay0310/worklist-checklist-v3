import streamlit as st
import pandas as pd
import io

st.set_page_config(page_title="작업목록표 입력 시스템", layout="wide")
st.title("📝 작업목록표 입력 시스템")

# 단위작업공정 여러 개 입력
task_units = []
num_units = st.number_input("단위작업공정 개수", min_value=1, step=1, value=1)
for i in range(int(num_units)):
    with st.expander(f"단위작업공정 {i+1} 입력"):
        with st.form(f"worklist_form_{i}"):
            col1, col2, col3 = st.columns(3)
            with col1:
                company = st.text_input("회사명", key=f"company_{i}")
            with col2:
                task_unit = st.text_input("단위작업명", key=f"task_unit_{i}")
            with col3:
                num_workers = st.number_input("작업자 수", min_value=1, step=1, key=f"num_workers_{i}")

            st.markdown("---")
            st.subheader("📦 중량물 정보")
            weights = []
            num_weights = st.number_input("중량물 종류 수", min_value=0, step=1, key=f"num_weights_{i}")
            for j in range(int(num_weights)):
                cols = st.columns(3)
                wtype = cols[0].text_input(f"중량물 종류 {j+1}", key=f"wtype_{i}_{j}")
                wcount = cols[1].number_input(f"중량물 개수 {j+1}", min_value=0, step=1, key=f"wcount_{i}_{j}")
                wweight = cols[2].number_input(f"중량물 무게(kg) {j+1}", min_value=0.0, step=0.1, key=f"wweight_{i}_{j}")
                weights.append((wtype, wcount, wweight))

            st.subheader("🔧 수공구 정보")
            tools = []
            num_tools = st.number_input("수공구 종류 수", min_value=0, step=1, key=f"num_tools_{i}")
            for j in range(int(num_tools)):
                cols = st.columns(3)
                ttype = cols[0].text_input(f"수공구 종류 {j+1}", key=f"ttype_{i}_{j}")
                tcount = cols[1].number_input(f"수공구 개수 {j+1}", min_value=0, step=1, key=f"tcount_{i}_{j}")
                tweight = cols[2].number_input(f"수공구 무게(kg) {j+1}", min_value=0.0, step=0.1, key=f"tweight_{i}_{j}")
                tools.append((ttype, tcount, tweight))

            submitted = st.form_submit_button("저장하기")
            if submitted:
                task_units.append({
                    "회사명": company,
                    "단위작업명": task_unit,
                    "작업자 수": num_workers,
                    "중량물": weights,
                    "수공구": tools
                })

# 저장 및 엑셀 다운로드
if task_units:
    from openpyxl import Workbook

    wb = Workbook()
    ws = wb.active
    ws.title = "작업목록표"

    # 헤더
    headers = [
        "회사명", "단위작업명", "작업자 수",
        "중량물 종류", "중량물 개수", "중량물 무게(kg)",
        "수공구 종류", "수공구 개수", "수공구 무게(kg)"
    ]
    ws.append(headers)

    for unit in task_units:
        weights = unit["중량물"]
        tools = unit["수공구"]
        max_len = max(len(weights), len(tools))
        for i in range(max_len):
            w = weights[i] if i < len(weights) else ("", "", "")
            t = tools[i] if i < len(tools) else ("", "", "")
            row = [unit["회사명"], unit["단위작업명"], unit["작업자 수"], *w, *t]
            ws.append(row)

    buffer = io.BytesIO()
    wb.save(buffer)
    buffer.seek(0)

    st.success("✅ 저장이 완료되었습니다!")
    st.download_button(
        "📥 엑셀파일 다운로드",
        data=buffer,
        file_name=f"작업목록표_전체_{pd.Timestamp.now().strftime('%Y%m%d_%H%M%S')}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
