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

# 1단계 파일 처리 수정 부분
def process_uploaded_files(uploaded_files):
    processed_files_data = {}
    output = BytesIO()
    with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
        for uploaded_file in uploaded_files:
            file_name = uploaded_file.name
            # 파일명에서 영역명_세부영역명 추출
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

                # 모든 시트에 동일한 영역명_세부영역명 부여
                df['영역'] = base_sheet_name

                if len(excel_file.sheet_names) == 1:
                    new_sheet_name = base_sheet_name[:31]
                else:
                    new_sheet_name = f"{base_sheet_name}_{sheet_name}"[:31]

                df.to_excel(writer, sheet_name=new_sheet_name, index=False)
                sheet_dfs.append((new_sheet_name, df))
            processed_files_data[file_name] = sheet_dfs
    output.seek(0)
    return output, processed_files_data


# 2단계 데이터 처리 수정 부분 (영역 추출 시 sheet_name 대신 1단계에서 이미 넣어준 '영역' 컬럼 사용)
def process_step2_data(step1_data):
    with pd.ExcelFile(step1_data) as excel_file:
        all_data = []
        for sheet_name in excel_file.sheet_names:
            df = excel_file.parse(sheet_name=sheet_name)

            max_length_col = df.apply(lambda col: col.astype(str).str.len().max(), axis=0).idxmax()
            df.columns = df.columns.str.replace(max_length_col, '기재내용', regex=False)

            if '학년' in df.columns:
                df['학년'] = df['학년'].astype(str).str.extract('(\d+)').astype(int)
                df['반'] = df['반'].astype(str).str.extract('(\d+)').astype(int)
                st.write(sheet_name)
                st.write(df.head())
                df['번호'] = df['번호'].astype(str).str.extract('(\d+)').astype(int)

            if '학번' in df.columns:
                df['학년'] = df['학번'].astype(str).str[0].astype(int)
                df['반'] = df['학번'].astype(str).str[1:3].astype(int)
                df['번호'] = df['학번'].astype(str).str[3:].astype(int)
                st.warning(f"⚠️ [{sheet_name}] 파일의 '학번'이 학년, 반, 번호로 분리되었습니다.")


            df = df[['학년', '반', '번호', '이름', '영역', '기재내용']]
            df['기재내용'] = df['기재내용'].apply(lambda x: x[:x.rfind('.')+1] + ' ' if isinstance(x, str) and '.' in x else x)
            all_data.append(df)

        final_df = pd.concat(all_data, ignore_index=True)

        for col in ['이름', '기재내용', '영역']:
            final_df[col] = final_df[col].apply(normalize_text)

        # 영역명, 세부영역명 추출 (이제 "영역명_세부파일명"은 1단계에서 영역 컬럼에 동일하게 할당됨)
        final_df[['영역명', '세부영역명']] = final_df['영역'].apply(lambda x: pd.Series(extract_fields(x)))
        for col in ['영역명', '세부영역명']:
            final_df[col] = final_df[col].apply(normalize_text)

        final_df = final_df[['학년', '반', '번호', '이름', '영역명', '세부영역명', '기재내용']]
        return final_df

def create_pivot_tables(final_df):
    section_df_list = []
    for section_name in final_df['영역명'].unique():
        section_df = final_df[final_df['영역명'] == section_name]

        # 중복 데이터 해결: 중복된 '학년', '반', '번호', '이름', '세부영역명'에 대해 기재내용을 병합
        section_df = section_df.groupby(['학년', '반', '번호', '이름', '세부영역명'], as_index=False).agg({
            '기재내용': lambda x: ' | '.join(x.dropna().astype(str))
        })

        # 피벗 테이블 생성
        section_df_pivot = section_df.pivot(index=['학년', '반', '번호', '이름'], columns='세부영역명', values='기재내용')
        section_df_pivot.reset_index(inplace=True)  # 인덱스 초기화
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

st.markdown("""
<div style="font-size: 14px; background-color: #f0f2f6; padding: 10px; border-radius: 5px;">
<b>📊 엑셀 데이터 처리 앱</b><br>
학번이나 학년/반/번호/이름이 포함된 학생 데이터를 통합, 정렬, 학생별 모아보기 등의 기능을 제공하는 앱입니다.
<br><br>
<b>🔑 주요 기능</b><br>
1️⃣ 파일 업로드 및 통합<br>
2️⃣ 데이터 처리 및 변환<br>
3️⃣ 피벗 테이블 생성<br>
4️⃣ 수식 추가 및 반별 시트 생성<br><br>
<b>📂 파일명 규칙</b><br>
- "영역명_세부파일명.xlsx" 형식 필수<br>
- 동일 영역명 파일끼리 업로드<br><br>
<b>🔒 보안 안내</b><br>
- 업로드된 데이터는 로컬 세션 내에서만 처리되며, 외부 서버로 전송되지 않습니다.<br>
- 즉, 데이터는 사용자가 세션을 종료하면 즉시 삭제됩니다.<br><br>
<b>🔍 각 단계에서 미리보기 제공 및 다운로드 가능</b><br>
👤 creator : Subhin Hwang, 💻 language : python</div>
""", unsafe_allow_html=True)



if 'step1_data' not in st.session_state:
    st.session_state.step1_data = None  
if 'step2_data' not in st.session_state:
    st.session_state.step2_data = None  
if 'step3_data' not in st.session_state:
    st.session_state.step3_data = []  
if 'step4_data' not in st.session_state:
    st.session_state.step4_data = []  


# 학생 명렬표 업로드 기능 추가 (1단계 전에)
st.write(" ")
st.write("##### 📤 [선택]전체 학생 명렬표 업로드")

roster_file = st.file_uploader(
    "학생 명렬표를 업로드하세요. 명렬표를 업로드하면 최종 파일에 특기사항이 없는 학생들(모든 행사에 참여하지 않은 학생)도 포함되어 파일이 생성됩니다. ",
    type=["xls", "xlsx"], key="roster"
)
if roster_file is not None:
    roster_df = pd.read_excel(roster_file)
    # '성명' 컬럼이 있으면 '이름'으로 변경
    if '성명' in roster_df.columns:
        roster_df.rename(columns={'성명': '이름'}, inplace=True)
    # 필요하다면 roster_df 처리 (학년, 반, 번호 추출 등)
    if '학번' in roster_df.columns:
        roster_df['학년'] = roster_df['학번'].astype(str).str[0].astype(int)
        roster_df['반'] = roster_df['학번'].astype(str).str[1:3].astype(int)
        roster_df['번호'] = roster_df['학번'].astype(str).str[3:].astype(int)
    # 이름 정규화
    if '이름' in roster_df.columns:
        roster_df['이름'] = roster_df['이름'].apply(normalize_text)
else:
    roster_df = None


# 1단계: 파일 업로드 및 통합
with st.expander("1단계: 엑셀 파일 업로드 및 통합", expanded=True):
    st.write("### 📤 엑셀 파일을 업로드하기")
    uploaded_files = st.file_uploader("엑셀파일을 업로드해주세요. ", type=["xls", "xlsx"], accept_multiple_files=True)
    if uploaded_files:
        output, processed_files_data = process_uploaded_files(uploaded_files)
        st.session_state.step1_data = output
        st.success("🎉 1단계 처리 완료! '성명'을 '이름'으로 통일하고 통합 문서를 생성했습니다.")

        # 업로드한 모든 파일을 tabs로 보기
        tab_names = [f"파일: {name}" for name in processed_files_data.keys()]
        tabs = st.tabs(tab_names)
        for i, (file_name, sheet_dfs) in enumerate(processed_files_data.items()):
            with tabs[i]:
                # st.write(f"**{file_name} 처리 결과**")
                for sheet_name, df in sheet_dfs:
                    n, m = df.shape
                    st.info(f"파일명 : {file_name}\n\n시트명: {sheet_name}, **{n}명의 학생 데이터가** 포함되어 있습니다. ")
                    st.dataframe(df.head(5))

        st.download_button(
            label="1단계 결과 다운로드: 여러 통합문서를 하나의 통합문서로",
            data=st.session_state.step1_data,
            file_name="특기사항_모든파일_통합문서.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

# 2단계: 데이터 처리 및 변환
with st.expander("2단계: 데이터 처리 및 변환", expanded=True):
    if st.session_state.step1_data:
        st.write("### ⌨️ 하나의 시트로 모든 데이터 모으기")
        final_df = process_step2_data(st.session_state.step1_data)
        st.session_state.step2_data = final_df
        st.write("📋 2단계 처리 결과 (미리보기)")
        st.dataframe(final_df.head(10))

        output_step2 = BytesIO()
        final_df.to_excel(output_step2, index=False, engine='xlsxwriter')
        output_step2.seek(0)
        st.success("🎉 2단계 처리 완료! 모든 데이터를 하나의 시트로 통합하였습니다.")
        st.download_button(
            label="2단계 결과 다운로드: 모든 데이터를 하나의 시트로",
            data=output_step2,
            file_name="특기사항_하나의시트.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

# 3단계: 피벗 테이블 생성
with st.expander("3단계: 학생별 데이터 모아보기 생성", expanded=True):
    if st.session_state.step2_data is not None:
        st.write("### 🗂️ 학생별 데이터 모아보기 생성")
        section_df_list = create_pivot_tables(st.session_state.step2_data)
        st.session_state.step3_data = section_df_list

        for section_name, df in section_df_list:
            st.dataframe(df.head(10))
            output_step3 = BytesIO()
            with pd.ExcelWriter(output_step3, engine="xlsxwriter") as writer:
                df.to_excel(writer, index=False, sheet_name="특기사항")
            output_step3.seek(0)
            st.download_button(
                label=f"3단계 결과 다운로드: {section_name} 학생별 합본",
                data=output_step3,
                file_name=f"{section_name}_특기사항_학생별모음.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
    # 3단계 끝부분 수정: 3단계 결과에 학생명렬표 기반 누락 학생 추가 후 정렬
    if roster_df is not None and st.session_state.step3_data:
        updated_section_df_list = []
        for section_name, df in st.session_state.step3_data:
            df = pd.merge(roster_df[['학년','반','번호','이름']], df, on=['학년','반','번호','이름'], how='left')
            df = df.sort_values(['학년','반','번호'])
            updated_section_df_list.append((section_name, df))
        st.session_state.step3_data = updated_section_df_list


# 4단계: 엑셀 수식 및 열 설정 추가
with st.expander("4단계: 엑셀 수식 및 열 설정 추가", expanded=True):
    # 4단계 시작부분 수정: 4단계에서도 학생명렬표와 비교하여 누락 학생 추가 후 정렬
    if roster_df is not None and st.session_state.step3_data:
        updated_section_df_list = []
        for section_name, df in st.session_state.step3_data:
            df = pd.merge(roster_df[['학년','반','번호','이름']], df, on=['학년','반','번호','이름'], how='left')
            df = df.sort_values(['학년','반','번호'])
            # df를 정리한 후 수식 추가 함수에 전달
            temp_output, preview_data = add_excel_formulas(section_name, df)
            updated_section_df_list.append((section_name, preview_data))
        # 4단계에서는 preview_data는 미리보기용이므로 실제 데이터 정리는 위 merge/정렬 단계에서 이미 완료됨
        # 여기서는 st.session_state.step4_data 등에 할당할 필요가 있다면 추가
        st.session_state.step4_data = updated_section_df_list

        import datetime
        import pytz

        # 한국 시간대로 현재 날짜와 시간 포맷팅
        kst = pytz.timezone('Asia/Seoul')
        current_datetime_kst = datetime.datetime.now(kst).strftime("%Y%m%d_%H%M")

        if st.session_state.step3_data:
            st.write("### ✏️ 특기사항 합본 및 바이트 계산 수식 추가")
            # 업로드 파일 개수 추출 (1단계 처리 시 저장한 processed_files_data 사용 가정)
            combined_files_count = len(processed_files_data) if 'processed_files_data' in locals() else 0

            for section_name, df in st.session_state.step3_data:
                temp_output, preview_data = add_excel_formulas(section_name, df)
                st.dataframe(preview_data.head(10))
                temp_output.seek(0)
                st.download_button(
                    label=f"4단계 결과 다운로드: {section_name} 모든 특기사항 합친 데이터 및 바이트 추가한 최종본",
                    data=temp_output,
                    file_name=f"{section_name}_특기사항_{combined_files_count}개_합본_바이트추가_{current_datetime_kst}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )

st.markdown("---")
st.markdown("""
<div style="text-align: center; margin-top: 20px;">
    <p style="font-size: 14px; color: gray;">© 2024 <strong>Excel Process</strong> made by <strong>Subhin Hwang</strong></p>
    <p style="font-size: 12px; color: gray;">
        Designed with ❤️ to simplify Excel workflows.
    </p>
    <p style="font-size: 14px; color: gray;">
        📧 Contact: <a href="mailto:sbhath17@gmail.com">sbhath17@gmail.com</a>
    </p>
</div>
""", unsafe_allow_html=True)
