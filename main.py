# Importing useful libraries
import streamlit as st
from scripts import utils
import pandas as pd
import os

# Creating a title and icon to the webpage
st.set_page_config(
    page_title="Report Generating App",
    page_icon="📚"
)

st.header("SubID based revenue checking Simple App")

st.subheader("Data Uploading Section")

comprehensive_report = st.file_uploader(
    'Please Upload the comprehensive report',
    type = ['xlsx']
)


if st.button('Generate Report'):

    if comprehensive_report:

        with st.spinner("Generating..."):

            df = pd.read_excel(comprehensive_report)

            st.dataframe(df)

            # Specifying file paths for saving 
            temp_file_path = "uploaded_file.xlsx"
            output_file_path = "output_file.xlsx"

            # Save the uploaded file to the local directory
            with open(temp_file_path, "wb") as f:
                f.write(comprehensive_report.getbuffer())

            try:
                utils.generate_report(temp_file_path,output_file_path)

                # Reading the revenue data
                xls = pd.read_excel(output_file_path, sheet_name=None)
                data_frames = {}

                # xls is a dictionary where the keys are the sheet names and the values are the DataFrames
                for sheet_name, sheet_df in xls.items():
                    data_frames[sheet_name] = sheet_df

                downloadable_data = utils.generate_excel_file(data_frames)

                # Providing a download button
                st.download_button(
                    label="Download Excel File",
                    data=downloadable_data,
                    file_name="generated_file.xlsx",
                    mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
                )

                os.remove(temp_file_path)
                os.remove(output_file_path)

            except:
                st.warning("Please upload only the comprehensive report, the app is currently working for that report!!!")

    else:

        st.warning("Upload a file first please")