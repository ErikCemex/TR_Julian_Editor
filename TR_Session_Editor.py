import streamlit as st 
from utils.TR_Session_Editor_backend import read_woorkbook, process_workbook, save_workbook_to_bytes
import re

st.set_page_config(page_title="Excel_Editor_TR", layout="wide")

st.image("Cemex Logo.png", width=300)
st.markdown("<h1 style='text-align: center;'>TR Session Processor</h1>", unsafe_allow_html=True)


Upload_TR_Session = st.file_uploader(
    label='Upload the Talent Review Session Report', 
     key="uploader1")




if Upload_TR_Session is not None:
    st.write('---')
    try:
        df_TR_Session_input = read_woorkbook(Upload_TR_Session)
        processed_wb = process_workbook(df_TR_Session_input)
        file_bytes = save_workbook_to_bytes(processed_wb)

        sheet_list = processed_wb['List View']
        file_name = sheet_list['A1'].value
        if not file_name:
            file_name = "Modified_TR_List_View"
        safe_filename = re.sub(r'[\\/*?:"<>|]', "_", str(file_name)) + ".xlsx"
        spacer1, center_col, spacer2 = st.columns([2, 2.2, 2])
        with center_col:
            st.download_button(
                label="📥 Download the cleaned Talent Review session file",
                data=file_bytes,
                file_name=safe_filename,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
    except Exception as e:
        st.error(f"❌ An error occurred while processing the files: {e}")

else:
    st.write('---')
    st.markdown(
        """
        <div style='text-align: center; background-color: #e1f5fe; padding: 10px; border-radius: 5px; color: #31708f; border: 1px solid #bce8f1;'>
            🔄 Please upload the required file to generate the processed data.
        </div>
        """,
        unsafe_allow_html=True
    )
