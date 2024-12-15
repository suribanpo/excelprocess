import streamlit as st
import pandas as pd
import re
from io import BytesIO

# 학번 추출 함수
def extract_student_id(value):
    """
    문자열에서 연속된 숫자 5개를 찾아 학번으로 반환합니다.
    :param value: 문자열
    :return: 연속된 5자리 숫자 (학번) 또는 None
    """
    match = re.search(r'\d{5}', str(value))  # 연속된 5자리 숫자 찾기 (문자와 숫자가 붙어 있어도 동작)
    return match.group(0) if match else None

# 제목 및 소개
st.title("📊 구글 설문 응답 통합 앱")

st.markdown("""
<div style="font-size: 14px; background-color: #f8f9fa; padding: 10px; border-radius: 5px;">
<b>📑 구글 설문 응답 통합 앱</b><br>
업로드된 설문 데이터를 통합하여 키(예: 이메일, 학번)를 기준으로 응답을 하나로 병합할 수 있습니다.<br><br>
<b>🔑 주요 기능</b><br>
1️⃣ 여러 파일 업로드 및 키 열 설정<br>
2️⃣ 학번 추출(선택적) 및 통합<br>
3️⃣ 최종 데이터 미리보기 및 엑셀 다운로드<br><br>
<b>📂 지원 파일 형식</b><br>
- CSV 파일 (.csv)<br>
- 엑셀 파일 (.xls, .xlsx)<br><br>
<b>🔒 보안 안내</b><br>
- 업로드된 데이터는 로컬 세션 내에서만 처리되며, 외부 서버로 전송되지 않습니다.<br>
- 세션 종료 시 모든 데이터는 삭제됩니다.<br>
</div>
""", unsafe_allow_html=True)

# 파일 업로드
uploaded_files = st.file_uploader(
    "📂 파일을 업로드하세요 (CSV 또는 Excel 형식)", 
    accept_multiple_files=True, 
    type=["csv", "xls", "xlsx"]
)

if uploaded_files:
    key_columns = {}
    dataframes = []
    st.markdown("#### 업로드된 파일 처리하기")

    for file in uploaded_files:
        st.write("---")
        st.write(f"파일: **{file.name}**")
        
        # 파일 읽기
        try:
            if file.name.endswith('.csv'):
                df = pd.read_csv(file)
            elif file.name.endswith(('.xls', '.xlsx')):
                df = pd.read_excel(file)
            else:
                st.warning(f"지원되지 않는 파일 형식입니다: {file.name}")
                continue
        except Exception as e:
            st.error(f"{file.name} 파일을 읽는 중 오류 발생: {e}")
            continue
        
        # 열 이름에서 파일명 제거 (파일 이름 접두사 제거)
        original_columns = [col.split("_", 1)[-1] for col in df.columns]
        df.columns = original_columns
        
        # 키 열 선택
        key_col = st.selectbox(
            f"**{file.name}**에서 키로 사용할 열을 선택하세요.",
            df.columns,
            index=1 if len(df.columns) > 1 else 0,  # 기본값: 두 번째 열
            key=file.name + "_key"
        )
        key_columns[file.name] = key_col
        
        # 학번 추출 체크박스
        extract_checkbox = st.checkbox(
            f"**{file.name}**에서 학번 추출(다섯 자리 숫자)",
            key=file.name + "_extract"
        )
        
        if extract_checkbox:
            # 학번 열 추가 (체크박스가 체크된 경우 학번 추출)
            df["학번"] = df[key_col].apply(extract_student_id)
            merge_key = "학번"  # 병합 기준: 학번
        else:
            # 체크박스가 체크되지 않은 경우 기존 키 열을 사용
            df["학번"] = df[key_col]  # 기존 키 열을 '학번' 열로 대체
            merge_key = key_col  # 병합 기준: 기존 키 열
        st.write(key_col)
        # 병합할 응답 열 선택
        columns_to_merge = st.multiselect(
            f"**{file.name}**에서 병합할 응답 열을 선택하세요.",
            [col for col in df.columns if col != key_col],  # 키 열 제외
            key=file.name + "_cols"
        )

        if columns_to_merge:
            # 응답 병합 함수
            def combine_responses(row):
                pairs = []
                for c in columns_to_merge:
                    pairs.append(f"✅[질문]{c}\n✅[답변]{row[c]}\n\n")
                return "\n".join(pairs)

            # 병합 기준 열 추가
            df[file.name + "_응답"] = df.apply(combine_responses, axis=1)

            # 병합 기준으로 사용할 열 추가
            df = df[[merge_key, file.name + "_응답"]]
            df.rename(columns={merge_key: "병합키"}, inplace=True)  # 병합 키 열 이름 통일

            dataframes.append(df)
        else:
            st.warning(f"**{file.name}**에서 병합할 열을 선택하지 않았습니다.")


    if dataframes:
        # 데이터 병합 (학번 열을 기준으로 병합)
        try:
            merged_df = pd.concat(dataframes, axis=1, join='outer')  # 학번 기준으로 병합
        except Exception as e:
            st.error(f"데이터 병합 중 오류 발생: {e}")
            st.stop()

        # 중복 열 이름 제거
        merged_df = merged_df.loc[:, ~merged_df.columns.duplicated()]

        st.markdown("### 2️⃣ 병합된 데이터 미리보기")
        st.dataframe(merged_df)

        # 엑셀 다운로드
        output = BytesIO()
        with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
            merged_df.to_excel(writer, index=False, sheet_name='합쳐진 응답')

            # 텍스트 줄바꿈 및 열 너비 설정
            workbook = writer.book
            worksheet = writer.sheets['합쳐진 응답']
            wrap_format = workbook.add_format({'text_wrap': True})
            for idx, col in enumerate(merged_df.columns):
                worksheet.set_column(idx, idx, 60, wrap_format)
        
        processed_data = output.getvalue()
        
        st.markdown("### 3️⃣ 다운로드")
        st.download_button(
            label="📥 병합된 데이터 다운로드 (Excel)",
            data=processed_data,
            file_name="설문데이터_병합.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
    else:
        st.error("데이터를 병합할 수 없습니다. 최소 하나의 파일에서 응답 열을 선택하세요.")
