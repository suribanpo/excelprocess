import streamlit as st
import pandas as pd
import openpyxl
from io import BytesIO

# 데이터 처리 함수
def drop_empty_rows(df):
    """빈 행 삭제"""
    return df.dropna(how='all', axis=0)

def drop_empty_columns(df):
    """빈 열 삭제"""
    return df.dropna(how='all', axis=1)

def drop_single_value_rows(df):
    """행에서 하나의 값만 존재하는 경우 삭제"""
    return df[df.apply(lambda row: row.count() > 1, axis=1)]

def drop_single_value_columns(df):
    """열에서 하나의 값만 존재하는 경우 삭제"""
    return df.loc[:, df.apply(lambda col: col.count() > 1, axis=0)]

def sanitize_columns(columns):
    """중복 또는 None 열 이름 처리"""
    sanitized = []
    seen = {}
    for col in columns:
        if col is None:
            col = "Unnamed"
        if col in seen:
            seen[col] += 1
            sanitized.append(f"{col}_{seen[col]}")
        else:
            seen[col] = 0
            sanitized.append(col)
    return sanitized

# Streamlit 앱
st.title("✨ 엑셀 시트별 병합 해제 및 옵션 처리 앱 ✨")
st.markdown("""
이 앱은 업로드된 엑셀 파일의 **각 시트 데이터를 병합 해제**하고, 데이터 처리 옵션을 제공합니다.  
처리된 데이터를 시트별로 다운로드할 수 있습니다.

### 🛠 주요 기능:
- **빈 행/열 삭제**
- **하나의 값만 있는 행/열 삭제**
- **시트별 데이터 미리보기**
- **처리된 데이터 다운로드**
""")

# 파일 업로드
uploaded_file = st.file_uploader("📤 엑셀 파일 업로드", type=["xlsx"], accept_multiple_files=False)

if uploaded_file:
    st.success(f"업로드된 파일: {uploaded_file.name}")
    with st.spinner("파일 처리 중... 잠시만 기다려 주세요 ⏳"):
        workbook = openpyxl.load_workbook(uploaded_file)
        sheet_names = workbook.sheetnames

        # 시트별 탭 생성
        tabs = st.tabs(sheet_names)
        processed_sheets = {}

        for tab, sheet_name in zip(tabs, sheet_names):
            with tab:
                sheet = workbook[sheet_name]
                data = sheet.values

                try:
                    columns = next(data)
                except StopIteration:
                    st.warning(f"시트 '{sheet_name}'에 데이터가 없습니다.")
                    continue

                # 중복 및 None 열 이름 처리
                columns = sanitize_columns(columns)
                df = pd.DataFrame(data, columns=columns)

                # 원본 데이터 출력
                st.subheader(f"📄 원본 데이터: {sheet_name}")
                st.dataframe(df, use_container_width=True)

                # 데이터 처리 옵션
                st.markdown("### 🛠 데이터 처리 옵션")
                col1, col2 = st.columns(2)
                with col1:
                    remove_empty_rows = st.checkbox("빈 행 삭제", key=f"rows_{sheet_name}")
                    remove_empty_columns = st.checkbox("빈 열 삭제", key=f"columns_{sheet_name}")
                with col2:
                    remove_single_value_rows = st.checkbox("하나의 값만 있는 행 삭제", key=f"single_rows_{sheet_name}")
                    remove_single_value_columns = st.checkbox("하나의 값만 있는 열 삭제", key=f"single_columns_{sheet_name}")

                # 처리 옵션 적용
                if remove_empty_rows:
                    df = drop_empty_rows(df)
                if remove_empty_columns:
                    df = drop_empty_columns(df)
                if remove_single_value_rows:
                    df = drop_single_value_rows(df)
                if remove_single_value_columns:
                    df = drop_single_value_columns(df)

                # 처리된 데이터 출력
                st.subheader(f"✅ 처리된 데이터: {sheet_name}")
                st.dataframe(df, use_container_width=True)

                # 시트별 다운로드 버튼
                sheet_output = BytesIO()
                with pd.ExcelWriter(sheet_output, engine='openpyxl') as writer:
                    df.to_excel(writer, index=False, sheet_name=sheet_name, header=False)
                sheet_output.seek(0)

                st.download_button(
                    label=f"💾 {sheet_name} 다운로드",
                    data=sheet_output,
                    file_name=f"{sheet_name}_병합해제.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                )
else:
    st.info("파일을 업로드하세요. 지원되는 파일 형식: **.xlsx**")
