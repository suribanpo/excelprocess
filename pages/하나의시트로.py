import streamlit as st
import pandas as pd
import os
from io import BytesIO

# Streamlit App
def main():
    # 타이틀과 소개 설명 추가
    st.title("📊 Excel Sheet Combiner")
    st.subheader("여러 개의 엑셀 파일과 시트를 한 번에 병합해보세요!")
    st.markdown("""
    이 도구를 사용하면 다중 Excel 파일의 여러 시트를 병합하여 하나의 데이터로 만들어줍니다.
    
    ### 사용 방법:
    1. **[Upload Excel files]** 버튼을 클릭하여 엑셀 파일(.xlsx)을 업로드하세요.
    2. 업로드된 파일의 시트가 병합되고, 결과를 화면에서 확인할 수 있습니다.
    3. **[Download Combined Excel]** 버튼으로 병합된 데이터를 다운로드하세요.
    """)

    # 파일 업로드
    uploaded_files = st.file_uploader(
        "여러 개의 Excel 파일을 업로드하세요",
        type=['xlsx'],
        accept_multiple_files=True,
        help="여러 파일을 선택하려면 Ctrl(Cmd)을 누른 상태에서 선택하세요."
    )

    if uploaded_files:
        st.success(f"총 {len(uploaded_files)}개의 파일이 업로드되었습니다!")
        combined_data = []

        for uploaded_file in uploaded_files:
            file_name = os.path.splitext(uploaded_file.name)[0]  # Get file name without extension
            excel_data = pd.ExcelFile(uploaded_file)

            # Iterate through all sheets in the file
            for sheet_name in excel_data.sheet_names:
                sheet_data = pd.read_excel(excel_data, sheet_name=sheet_name, index_col=None)

                # Create a new DataFrame with FileName and SheetName as headers in the first row
                header_data = pd.DataFrame([[file_name, sheet_name] + [None] * sheet_data.shape[1]])
                header_data.columns = sheet_data.columns.insert(0, 'SheetName').insert(0, 'FileName')
                sheet_data = pd.concat([header_data, sheet_data], ignore_index=True)

                combined_data.append(sheet_data)

        # Combine all dataframes
        if combined_data:
            combined_df = pd.concat(combined_data, ignore_index=True)
            combined_df = combined_df.dropna(how='all', axis=0)  # Remove completely empty rows
            st.write("### 🗂 병합된 데이터")
            st.dataframe(combined_df)

            # Option to download the combined data
            download_data = BytesIO()
            combined_df.to_excel(download_data, index=False, engine='openpyxl', header=False)
            download_data.seek(0)

            st.download_button(
                label="📥 병합된 Excel 다운로드",
                data=download_data,
                file_name="combined_data.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
            st.success("병합된 파일을 다운로드할 준비가 되었습니다!")

    else:
        st.info("업로드된 파일이 없습니다. Excel 파일을 업로드해 주세요.")

if __name__ == "__main__":
    main()
