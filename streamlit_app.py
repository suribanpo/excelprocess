import streamlit as st
import pandas as pd
import unicodedata
from io import BytesIO
from openpyxl import load_workbook
from openpyxl.utils import get_column_letter
from openpyxl.styles import Font, Alignment

st.set_page_config(
    page_title="엑셀 데이터 통합 및 처리",
    page_icon="📑",
    layout="wide"
)

#####################
# 헬퍼 함수 정의
#####################
def normalize_text(value):
    if isinstance(value, str):
        return unicodedata.normalize('NFC', value)
    return value

def extract_fields(sheet_name):
    # 파일명 형태: "영역명_세부파일명"
    parts = sheet_name.split('_', 1)
    if len(parts) == 2:
        return parts[0], parts[1]
    else:
        return parts[0], ""

# 1단계 파일 처리 수정
def process_uploaded_files(uploaded_files):
    processed_files_data = {}
    output = BytesIO()
    with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
        for uploaded_file in uploaded_files:
            file_name = uploaded_file.name
            base_sheet_name = '_'.join(file_name.split('_')[:2])
            excel_file = pd.ExcelFile(uploaded_file)
            sheet_dfs = []
            for sheet_name in excel_file.sheet_names:
                df = excel_file.parse(sheet_name=sheet_name, header=None)
                name_row_index = df[df.apply(lambda row: row.astype(str).str.contains('이름|성명').any(), axis=1)].index[0]
                df.columns = df.iloc[name_row_index].str.replace('성명', '이름')
                df = df[name_row_index + 1:]

                # '학년' 열의 값에서 숫자만 추출
                if '학년' in df.columns:
                    df['학년'] = df['학년'].astype(str).str.extract('(\d+)').astype(int)

                if len(excel_file.sheet_names) == 1:
                    new_sheet_name = base_sheet_name[:31]
                else:
                    new_sheet_name = f"{base_sheet_name}_{sheet_name}"

                df.to_excel(writer, sheet_name=new_sheet_name, index=False)
                sheet_dfs.append((new_sheet_name, df))
            processed_files_data[file_name] = sheet_dfs
    output.seek(0)
    return output, processed_files_data

def process_step2_data(step1_data):
    with pd.ExcelFile(step1_data) as excel_file:
        all_data = []
        for sheet_name in excel_file.sheet_names:
            normalized_sheet_name = unicodedata.normalize('NFC', sheet_name)
            df = excel_file.parse(sheet_name=sheet_name)

            max_length_col = df.apply(lambda col: col.astype(str).str.len().max(), axis=0).idxmax()
            df.columns = df.columns.str.replace(max_length_col, '기재내용', regex=False)

            if '학년' in df.columns:
                df['학년'] = df['학년'].astype(str).str.extract('(\d+)').astype(int)
                df['반'] = df['반'].astype(str).str.extract('(\d+)').astype(int)
                df['번호'] = df['번호'].astype(str).str.extract('(\d+)').astype(int)

            if '학번' in df.columns:
                df['학년'] = df['학번'].astype(str).str[0].astype(int)
                df['반'] = df['학번'].astype(str).str[1:3].astype(int)
                df['번호'] = df['학번'].astype(str).str[3:].astype(int)
                st.warning(f"⚠️ [{sheet_name}] 파일의 '학번'이 학년, 반, 번호로 분리되었습니다.")

            df['영역'] = normalized_sheet_name
            df = df[['학년', '반', '번호', '이름', '영역', '기재내용']]
            df['기재내용'] = df['기재내용'].apply(lambda x: x[:x.rfind('.')+1] + ' ' if isinstance(x, str) and '.' in x else x)
            all_data.append(df)

        final_df = pd.concat(all_data, ignore_index=True)

        for col in ['이름', '기재내용', '영역']:
            final_df[col] = final_df[col].apply(normalize_text)

        # 영역명, 세부영역명 추출 (이제 "영역명_세부파일명" 형식)
        final_df[['영역명', '세부영역명']] = final_df['영역'].apply(lambda x: pd.Series(extract_fields(x)))
        for col in ['영역명', '세부영역명']:
            final_df[col] = final_df[col].apply(normalize_text)

        final_df = final_df[['학년', '반', '번호', '이름', '영역명', '세부영역명', '기재내용']]
        return final_df

def create_pivot_tables(final_df):
    section_df_list = []
    for section_name in final_df['영역명'].unique():
        section_df = final_df[final_df['영역명'] == section_name]
        section_df_pivot = section_df.pivot(index=['학년', '반', '번호', '이름'], columns='세부영역명', values='기재내용')
        section_df_pivot.reset_index(inplace=True)
        section_df_list.append((section_name, section_df_pivot))
    return section_df_list

def add_excel_formulas(section_name, df):
    output_step4 = BytesIO()
    file_name = f"{section_name}_특기사항_학생별로_모아보기_엑셀수식포함.xlsx"
    df.to_excel(output_step4, index=False, sheet_name="특기사항")
    wb = load_workbook(output_step4)
    ws = wb["특기사항"]

    start_col = 5
    num_cols = len(df.columns) - start_col + 1
    combine_col_index = start_col + num_cols
    byte_col_index = combine_col_index + 1

    for idx in range(2, len(df) + 2):
        concat_formula = "=" + "CONCATENATE(" + ",".join(
            [f"{chr(64 + col)}{idx}" for col in range(start_col, start_col + num_cols)]
        ) + ")"
        ws[f"{chr(64 + combine_col_index)}{idx}"] = concat_formula

    for idx in range(2, ws.max_row + 1):
        ws[f"{get_column_letter(byte_col_index)}{idx}"] = (
            f'=LENB({get_column_letter(combine_col_index)}{idx})*2-LEN({get_column_letter(combine_col_index)}{idx})'
        )

    byte_column_letter = get_column_letter(byte_col_index)
    ws.column_dimensions[byte_column_letter].width = 15
    font_style = Font(size=22, bold=True)
    alignment_style = Alignment(horizontal="center", vertical="center")

    for row_idx in range(2, ws.max_row + 1):
        cell = ws[f"{byte_column_letter}{row_idx}"]
        cell.font = font_style
        cell.alignment = alignment_style

    ws[f"{get_column_letter(combine_col_index)}1"] = "특기사항 합본"
    ws[f"{get_column_letter(byte_col_index)}1"] = "바이트"

    for col_idx in range(start_col, byte_col_index + 1):
        col_letter = get_column_letter(col_idx)
        ws.column_dimensions[col_letter].width = 50 if col_idx != byte_col_index else 20
        for row_idx in range(2, ws.max_row + 1):
            cell = ws[f"{col_letter}{row_idx}"]
            cell.alignment = cell.alignment.copy(wrap_text=True)

    temp_output = BytesIO()
    wb.save(temp_output)
    wb.close()

    preview_data = pd.DataFrame(ws.values)
    preview_data.columns = preview_data.iloc[0]
    preview_data = preview_data[1:]
    return temp_output, preview_data

#####################
# UI 시작
#####################

st.title("📑 엑셀 데이터 처리 앱")
st.info("""
학년, 반, 학번, 학생 이름 등이 포함된 여러 엑셀 파일을 하나로 통합하고 데이터를 학생별로 정렬하는 앱입니다.

**주요 기능**  
1️⃣ 엑셀 파일 업로드 및 통합    
2️⃣ 데이터 처리 및 변환  
3️⃣ 피벗 테이블 생성  
4️⃣ 엑셀 수식 추가 및 반별 시트 생성

**파일명 규칙**  
"영역명_세부파일명.xlsx" 형태로 파일명 지정해주세요.(예: "자율활동_000행사 특기사항.xlsx")
동일한 영역명의 파일만 한 번에 업로드해주세요.(자율은 자율끼리, 진로는 진로끼리)

각 단계에서는 미리보기를 통해 결과를 확인하고, 처리된 데이터를 다운로드할 수 있습니다.
""")

if 'step1_data' not in st.session_state:
    st.session_state.step1_data = None  
if 'step2_data' not in st.session_state:
    st.session_state.step2_data = None  
if 'step3_data' not in st.session_state:
    st.session_state.step3_data = []  
if 'step4_data' not in st.session_state:
    st.session_state.step4_data = []  

# 1단계: 파일 업로드 및 통합
with st.expander("1단계: 엑셀 파일 업로드 및 통합", expanded=True):
    uploaded_files = st.file_uploader("📤 엑셀 파일을 업로드하세요", type=["xls", "xlsx"], accept_multiple_files=True)
    if uploaded_files:
        output, processed_files_data = process_uploaded_files(uploaded_files)
        st.session_state.step1_data = output
        st.success("🎉 1단계 처리 완료! '이름'으로 통일하고 통합 문서를 생성했습니다.")

        # 업로드한 모든 파일을 tabs로 보기
        tab_names = [f"파일: {name}" for name in processed_files_data.keys()]
        tabs = st.tabs(tab_names)
        for i, (file_name, sheet_dfs) in enumerate(processed_files_data.items()):
            with tabs[i]:
                st.write(f"**{file_name} 처리 결과**")
                for sheet_name, df in sheet_dfs:
                    st.write(f"**시트명**: {sheet_name}")
                    st.dataframe(df.head(10))

        st.download_button(
            type='primary', 
            label="1단계 결과 다운로드: 통합 문서",
            data=st.session_state.step1_data,
            file_name="특기사항_모든파일_통합문서.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

# 2단계: 데이터 처리 및 변환
with st.expander("2단계: 데이터 처리 및 변환", expanded=True):
    if st.session_state.step1_data:
        final_df = process_step2_data(st.session_state.step1_data)
        st.session_state.step2_data = final_df
        st.write("📋 2단계 처리 결과 (미리보기)")
        st.dataframe(final_df.head(10))

        output_step2 = BytesIO()
        final_df.to_excel(output_step2, index=False, engine='xlsxwriter')
        output_step2.seek(0)
        st.success("🎉 2단계 처리 완료! 모든 데이터를 하나의 시트로 통합하였습니다.")
        st.download_button(
            type='primary', 
            label="2단계 결과 다운로드",
            data=output_step2,
            file_name="특기사항_하나의시트.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

# 3단계: 피벗 테이블 생성
with st.expander("3단계: 피벗 테이블 생성", expanded=True):
    if st.session_state.step2_data is not None:
        st.write("### 🗂️ 영역별 피벗 테이블 생성")
        section_df_list = create_pivot_tables(st.session_state.step2_data)
        st.session_state.step3_data = section_df_list

        for section_name, df in section_df_list:
            st.write(f"#### 📋 피벗 테이블 미리보기: **{section_name}**")
            st.dataframe(df.head(10))
            output_step3 = BytesIO()
            with pd.ExcelWriter(output_step3, engine="xlsxwriter") as writer:
                df.to_excel(writer, index=False, sheet_name="특기사항")
            output_step3.seek(0)
            st.download_button(
                type='primary',
                label=f"3단계 결과 다운로드: {section_name}",
                data=output_step3,
                file_name=f"{section_name}_특기사항_학생별모음.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

# 4단계: 엑셀 수식 및 열 설정 추가
with st.expander("4단계: 엑셀 수식 및 열 설정 추가", expanded=True):
    if st.session_state.step3_data:
        st.write("### ✏️ 특기사항 합본 및 바이트 계산 수식 추가")
        for section_name, df in st.session_state.step3_data:
            temp_output, preview_data = add_excel_formulas(section_name, df)
            st.write(f"#### 📝 4단계 결과 미리보기: **{section_name}**")
            st.dataframe(preview_data.head(10))
            temp_output.seek(0)
            st.download_button(
                type='primary',
                label=f"4단계 결과 다운로드: {section_name}",
                data=temp_output,
                file_name=f"{section_name}_특기사항_합본_바이트추가.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

st.markdown("---")
st.markdown("""
<div style="text-align: center; margin-top: 20px;">
    <p style="font-size: 14px; color: gray;">© 2024 <strong>Excel Process</strong></p>
    <p style="font-size: 12px; color: gray;">
        Designed with ❤️ to simplify Excel workflows.
    </p>
</div>
""", unsafe_allow_html=True)
