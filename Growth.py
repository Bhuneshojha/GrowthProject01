import streamlit as st
import pandas as pd
import os
from io import BytesIO

st.set_page_config(page_title="Data Sweeper", page_icon="🧹", layout='wide')

# Css custom
st.markdown(
 """
    <style>
    .stApp{
            background-color:#f0f0f0;
            colour: #000000;

    }
   </style>

 """, 

    unsafe_allow_html=True

)

# Title
st.title("Data Sweeper integrator By Bhunesh Kumar")
st.write("tansform your files between csv and excel formats with built in dta cleaning and data analysis")


uploaded_file = st.file_uploader("Upload your files (Csv or excel):", type=["csv","xlsx"], accept_multiple_files=(True))

if uploaded_file:
    for file in uploaded_file:
        file_ext = os.path.splitext(file.name)[-1].lower()
        
        if file_ext == ".csv":
            df = pd.read_csv(file)
        elif file_ext == ".xlsx":
            df = pd.read_excel(file)
        else:
            st.error(f"File  type{file_ext} format not supported") 
            continue 
# File details
    st.write("Preview the head of data frame")
    st.dataframe(df.head())
# Data cleaning options
    st.subheader("Data Cleaning Options") 
    if st.checkbox(f"clean data for {file.name}"):
        col1 , col2 = st.columns(2)

        with col1:
            if st.button(f"Drop Duplicates from {file.name}"):
                df.drop_duplicates(inplace=True)
                st.write("Duplicates dropped")
        with col2:
            if st.button(f"Fill missing values for {file.name}"):
                numeric_cols = df.select_dtypes(include=["number"]).columns  
                df[numeric_cols] = df[numeric_cols].fillna(df[numeric_cols].mean())
                st.write("Missing values filled")


    st.subheader("Select columns to keep")  
    coloumns = st.multiselect(f"Select columns to keep from {file.name}", df.columns, default=df.columns)
    df = df[coloumns]


    # Data Analysis
    st.header("Data Analysis")
    if st.checkbox(f"Show data analysis for {file.name}"):
        st.bar_chart(df.select_dtypes (include='number').iloc[:, :2])

  # Conversrion options
    st.subheader("Conversion Options")
    conversion_type = st.radio(f"convert {file.name} to", ["csv", "excel"], key=file.name)
    if st.button(f"Convert {file.name} to {conversion_type}"):
        output = BytesIO()
        if conversion_type == "csv":
            df.to_csv(output, index=False)
            file_name = file.name.replace(".xlsx", ".csv")
            mime_type = "text/csv"
        elif conversion_type == "excel":
            df.to_excel(output, index==False)
            file_name = file.name.replace(file_ext, ".xlsx")
            mime_type = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        output.seek(0)

        st.download_button(
            label = f"Download {file_name} as {conversion_type}",
            data = output,
            file_name=file_name,
            mime=mime_type
        )

st.write("All files processed successfully")

                
