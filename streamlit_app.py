import streamlit as st
import pandas as pd
import unicodedata
from io import BytesIO
from openpyxl import load_workbook
from openpyxl.utils import get_column_letter

# 페이지 설정
st.set_page_config(
    page_title="엑셀 데이터 통합 및 처리",
    page_icon="📑",
)

# 페이지 요약 정보
st.title("📑 엑셀 데이터 처리 앱")
st.info("""
학번, 학년 반 번호와 학생이름 등이 담긴 여러 엑셀파일을 하나로 통합하는 앱입니다. 
1️⃣ **엑셀 파일 업로드 및 통합**: 여러 엑셀 파일을 업로드하여 하나의 통합 문서로 생성합니다.  
2️⃣ **데이터 처리 및 변환**: 통합된 데이터를 학년, 반, 번호 등 기준으로 정리하고 변환합니다.  
3️⃣ **피벗 테이블 생성**: 영역별 피벗 테이블을 생성하여 데이터를 재구성합니다.  
4️⃣ **엑셀 수식 추가**: 특기사항 합본과 바이트 계산 수식을 엑셀에 추가하고, 결과를 다운로드할 수 있습니다.  
5️⃣ **반별로 시트 생성**: 학년과 반을 기준으로 데이터를 구분하여 반별 시트를 생성합니다.

⚠️ **파일명 규칙**:
- 업로드하는 파일명은 반드시 **"창체_영역명_세부파일명.xlsx"** 형식을 따라야 합니다.  
  예: **"창체_자율활동_활동명1.xlsx"**, **"창체_진로활동_활동명2.xlsx"**
- 업로드하는 파일들의 '영역명'은 일치해야 합니다. 즉, 자율활동에 해당하는 파일들만 올리거나 진로활동에 해당하는 파일들만 올려야 합니다. 

✅ 올바른 파일명을 사용하면, 데이터 처리 및 변환 과정에서 오류 없이 작업이 진행됩니다.

각 단계에서 **미리보기**를 통해 데이터를 확인하고, 처리된 데이터를 바로 다운로드할 수 있습니다.  
앱을 활용하여 데이터 통합과 정리를 간편하게 수행하세요! 🎉
""")



# 세션 상태 초기화
if 'step1_data' not in st.session_state:
    st.session_state.step1_data = None  # 1단계 통합 데이터
if 'step2_data' not in st.session_state:
    st.session_state.step2_data = None  # 2단계 처리 데이터
if 'step3_data' not in st.session_state:
    st.session_state.step3_data = []  # 3단계 피벗화 데이터
if 'step4_data' not in st.session_state:
    st.session_state.step4_data = []  # 4단계 처리 데이터

# 1단계: 파일 업로드 및 통합
with st.expander("1단계: 엑셀 파일 업로드 및 통합", expanded=True):
    uploaded_files = st.file_uploader("엑셀 파일을 업로드하세요", type=["xls", "xlsx"], accept_multiple_files=True)

    if uploaded_files:
        missing_files = []
        output = BytesIO()
        with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
            for uploaded_file in uploaded_files:
                try:
                    file_name = uploaded_file.name
                    base_sheet_name = '_'.join(file_name.split('_')[:3])
                    
                    excel_file = pd.ExcelFile(uploaded_file)
                    for sheet_name in excel_file.sheet_names:
                        df = excel_file.parse(sheet_name=sheet_name, header=None)
                        name_row_index = df[df.apply(lambda row: row.astype(str).str.contains('이름|성명').any(), axis=1)].index[0]
                        df.columns = df.iloc[name_row_index].str.replace('성명', '이름')
                        df = df[name_row_index + 1:]
                        
                        if len(excel_file.sheet_names) == 1:
                            new_sheet_name = base_sheet_name[:31]
                        else:
                            new_sheet_name = f"{base_sheet_name}_{sheet_name}"
                        
                        df.to_excel(writer, sheet_name=new_sheet_name, index=False)
                except Exception as e:
                    missing_files.append(file_name)
        
        output.seek(0)
        st.session_state.step1_data = output  # 1단계 데이터 저장
        st.success("1단계 처리가 완료되었습니다!")
        st.download_button(
            label="1단계 데이터 다운로드",
            data=st.session_state.step1_data,
            file_name="창체_특기사항_모든파일을_통합문서로.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

# 2단계: 추가 데이터 처리
with st.expander("2단계: 데이터 처리 및 변환", expanded=True):
    if st.session_state.step1_data:
        with pd.ExcelFile(st.session_state.step1_data) as excel_file:
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
                
                df['영역'] = normalized_sheet_name
                df = df[['학년', '반', '번호', '이름', '영역', '기재내용']]
                df['기재내용'] = df['기재내용'].apply(lambda x: x[:x.rfind('.')+1] + ' ' if isinstance(x, str) and '.' in x else x)
                all_data.append(df)
            
            final_df = pd.concat(all_data, ignore_index=True)
            for col in ['이름', '기재내용', '영역']:
                final_df[col] = final_df[col].apply(lambda x: unicodedata.normalize('NFC', str(x)) if isinstance(x, str) else x)
            
            def extract_fields(val):
                parts = val.split('_', 2)
                if len(parts) == 3:
                    return parts[1], parts[2]
                else:
                    return (parts[1] if len(parts) > 1 else ""), (parts[2] if len(parts) > 2 else "")
            
            final_df[['영역명', '세부영역명']] = final_df['영역'].apply(lambda x: pd.Series(extract_fields(x)))
            for col in ['영역명', '세부영역명']:
                final_df[col] = final_df[col].apply(lambda x: unicodedata.normalize('NFC', str(x)) if isinstance(x, str) else x)
            
            final_df = final_df[['학년', '반', '번호', '이름', '영역명', '세부영역명', '기재내용']]
            st.write("2단계 처리 결과 (미리보기)")
            st.dataframe(final_df.head(10))
            
            st.session_state.step2_data = final_df  # 2단계 데이터 저장
            
            output_step2 = BytesIO()
            final_df.to_excel(output_step2, index=False, engine='xlsxwriter')
            output_step2.seek(0)
            st.download_button(
                label="2단계 데이터 다운로드",
                data=output_step2,
                file_name="창체_특기사항_하나의시트로.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

# 3단계: 피벗화 및 결과 저장
with st.expander("3단계: 피벗 테이블 생성", expanded=True):
    if st.session_state.step2_data is not None:
        section_df_list = []
        for section_name in st.session_state.step2_data['영역명'].unique():
            section_df = st.session_state.step2_data[st.session_state.step2_data['영역명'] == section_name]
            section_df_pivot = section_df.pivot(index=['학년', '반', '번호', '이름'], columns='세부영역명', values='기재내용')
            section_df_pivot.reset_index(inplace=True)
            section_df_list.append((section_name, section_df_pivot))
        
        st.session_state.step3_data = section_df_list  # 3단계 데이터 저장
        
        for section_name, df in section_df_list:
            st.write(f"피벗 테이블: {section_name} (미리보기)")
            st.dataframe(df.head(10))
            
            output_step3 = BytesIO()
            with pd.ExcelWriter(output_step3, engine="xlsxwriter") as writer:
                df.to_excel(writer, index=False, sheet_name="특기사항")
            
            output_step3.seek(0)
            st.download_button(
                label=f"{section_name} 데이터 다운로드",
                data=output_step3,
                file_name=f"{section_name}_특기사항_통합.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

# 4단계: 엑셀 수식 및 열 설정 추가
with st.expander("4단계: 엑셀 수식 및 열 설정 추가", expanded=True):
    if st.session_state.step3_data:
        for section_name, df in st.session_state.step3_data:
            output_step4 = BytesIO()
            file_name = f"{section_name}_특기사항_통합_엑셀수식포함.xlsx"
            
            # Excel 저장 및 수식 추가
            df.to_excel(output_step4, index=False, sheet_name="특기사항")
            
            # 엑셀 파일 열기 및 수식 추가
            wb = load_workbook(output_step4)
            ws = wb["특기사항"]
            
            # 열 계산
            start_col = 5  # 데이터 시작 열 (E 열부터)
            num_cols = len(df.columns) - start_col + 1
            combine_col_index = start_col + num_cols  # 특기사항 합본 열
            byte_col_index = combine_col_index + 1  # 바이트 열
            # "특기사항 합본" 열 수식 추가
            for idx in range(2, len(df) + 2):  # 데이터는 2행부터 시작
                # CONCATENATE에 사용할 열 주소 생성 (E열부터 마지막 데이터 열까지)
                concat_formula = "=" + "CONCATENATE(" + ",".join([
                    f"{chr(64 + col)}{idx}" for col in range(start_col, start_col + num_cols)
                ]) + ")"
                # "특기사항 합본" 열에 수식 추가
                ws[f"{chr(64 + combine_col_index)}{idx}"] = concat_formula

            # 바이트 계산 수식 추가
            for idx in range(2, ws.max_row + 1):
                ws[f"{get_column_letter(byte_col_index)}{idx}"] = (
                    f'=LENB({get_column_letter(combine_col_index)}{idx})*2-LEN({get_column_letter(combine_col_index)}{idx})'
                )
            
            # 헤더 추가
            ws[f"{get_column_letter(combine_col_index)}1"] = "특기사항 합본"
            ws[f"{get_column_letter(byte_col_index)}1"] = "바이트"

            # 열 너비와 자동 줄바꿈 설정
            for col_idx in range(start_col, byte_col_index + 1):  # 5열부터 마지막 열까지
                col_letter = get_column_letter(col_idx)  # 열 이름 (A, B, C...)
                ws.column_dimensions[col_letter].width = 50  # 열 너비 설정
                for row_idx in range(2, ws.max_row + 1):  # 각 행에 대해
                    cell = ws[f"{col_letter}{row_idx}"]
                    cell.alignment = cell.alignment.copy(wrap_text=True)  # 자동 줄바꿈 설정
            
            # 수정된 파일 저장
            temp_output = BytesIO()
            wb.save(temp_output)
            wb.close()
            
            # 미리보기 데이터 생성
            preview_data = pd.DataFrame(ws.values)  # 워크시트 내용을 DataFrame으로 변환
            preview_data.columns = preview_data.iloc[0]  # 첫 번째 행을 헤더로 사용
            preview_data = preview_data[1:]  # 첫 번째 행 제외

            # 미리보기 표시
            st.write(f"4단계 결과 미리보기: {section_name}")
            st.dataframe(preview_data.head(10))  # 상위 10개 데이터 표시
            
            # 다운로드 버튼
            temp_output.seek(0)
            st.download_button(
                label=f"{section_name} 합본 추가된 버전 다운로드",
                data=temp_output,
                file_name=file_name,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )



st.markdown("---")
st.markdown("""
<div style="text-align: center; margin-top: 20px;">
    <p style="font-size: 14px; color: gray;">© 2024 <strong>Excel Process</strong> made by <strong>Subhin Hwang</strong></p>
    <p style="font-size: 12px; color: gray;">
        Designed with ❤️ to simplify Excel workflows.
    </p>
</div>
""", unsafe_allow_html=True)
