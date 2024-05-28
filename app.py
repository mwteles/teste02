import streamlit as st
import pandas as pd
import plotly_express as px
import openpyxl

st.set_page_config(layout='wide')

uploaded_file = st.file_uploader("Choose a file",
                                 type='.xlsx')
if uploaded_file is not None:
    # To read file as bytes:

    # Can be used wherever a "file-like" object is accepted:
    df = pd.read_excel(uploaded_file,sheet_name='teste')
    st.write(df)
    xx=st.line_chart(df,x='letra',y='num')