from process import *
import streamlit as st
from io import BytesIO
import time

with open('../contents/fish_3ss.xlsx', 'rb') as rfile:
    st.download_button(
        label="Скачать пример xlsx",
        data=BytesIO(rfile.read()),
        file_name="example.xlsx",
    )

with open('../contents/test_3ss.pptx', 'rb') as rfile:
    st.download_button(
        label="Скачать пример pptx",
        data=BytesIO(rfile.read()),
        file_name="example.pptx",
    )

uploaded_file = st.file_uploader("Загрузите xlsx")

if uploaded_file:
    with open('../contents/file.xlsx', 'wb') as wfile:
        wfile.write(uploaded_file.read())
        process('../contents/file.xlsx', 'file.pptx')

    with open('../contents/file.pptx', 'rb') as rfile:
        buffer = BytesIO(rfile.read())
    buffer.seek(0)
    st.download_button(
        label="Скачать pptx",
        data=buffer,
        file_name=f"pptx-{time.time()}.pptx",
    )
