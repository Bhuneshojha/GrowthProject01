import openpyxl
import streamlit as st
import pandas as pd
import os
from io import BytesIO

st.set_page_config(page_title="Data Sweeper", page_icon="🧹", layout='wide')

# Css custom
st.markdown(
    """
    <style>
    .stApp {
        background-color: #f0f0f0;
        color: #000000;
    }
    </style>
    """,
    unsafe_allow_html=True
)

# Title
st.title("Data Sweeper Integrator By Bhunesh Kumar")
st.write("Transform your files between CSV and Excel formats with built-in data cleaning and analysis.")

uploaded_file = st.file_uploader("Upload your files (CSV or Excel):", type=["csv", "xlsx"], accept_multiple_files=True)

if uploaded_file:
    for file in uploaded_file:
        file_ext = os.path.splitext(file.name)[-1].lower()

        try:
            if file_ext == ".csv":
                df = pd.read_csv(file)
            elif file_ext == ".xlsx":
                df = pd.read_excel(file, engine='openpyxl')
            else:
                st.error(f"File type {file_ext} not supported")
                continue

            st.write(f"### Preview: {file.name}")
            st.dataframe(df.head())

            # Data Cleaning Options
            st.subheader(f"Data Cleaning Options for {file.name}")

            if st.checkbox(f"Clean Data for {file.name}"):
                col1, col2 = st.columns(2)

                with col1:
                    if st.button(f"Drop Duplicates {file.name}", key=f"{file.name}_drop"):
                        df.drop_duplicates(inplace=True)
                        st.success("Duplicates Dropped")

                with col2:
                    if st.button(f"Fill Missing Values {file.name}", key=f"{file.name}_fill"):
                        numeric_cols = df.select_dtypes(include=["number"]).columns
                        df[numeric_cols] = df[numeric_cols].fillna(df[numeric_cols].mean())
                        st.success("Missing Values Filled")

            # Select Columns
            columns = st.multiselect(f"Select Columns to Keep from {file.name}", df.columns, default=list(df.columns))
            df = df[columns]

            # Data Analysis
            st.subheader(f"Data Analysis for {file.name}")
            if st.checkbox(f"Show Data Analysis {file.name}"):
                st.bar_chart(df.select_dtypes(include='number').iloc[:, :2])

            # Conversion
            st.subheader(f"Conversion Options for {file.name}")
            conversion_type = st.radio(f"Convert {file.name} to:", ["csv", "excel"], key=f"{file.name}_convert")

            if st.button(f"Convert {file.name} to {conversion_type}", key=f"{file.name}_download"):
                output = BytesIO()
                if conversion_type == "csv":
                    df.to_csv(output, index=False)
                    file_name = file.name.replace(".xlsx", ".csv")
                    mime_type = "text/csv"
                elif conversion_type == "excel":
                    with pd.ExcelWriter(output, engine='openpyxl') as writer:
                        df.to_excel(writer, index=False)
                    file_name = file.name.replace(".csv", ".xlsx")
                    mime_type = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"

                output.seek(0)
                st.download_button(
                    label=f"Download {file_name}",
                    data=output,
                    file_name=file_name,
                    mime=mime_type
                )

        except Exception as e:
            st.error(f"An error occurred with file {file.name}: {e}")

st.success("All files processed successfully")
