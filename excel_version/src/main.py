import os
from process import *
import streamlit as st
from io import BytesIO
import time

CONTENT_DIR = '../contents'

with open(os.path.join(CONTENT_DIR, 'fish_3ss.xlsx'), 'rb') as rfile:
    st.download_button(
        label="Скачать пример xlsx",
        data=BytesIO(rfile.read()),
        file_name="example.xlsx",
    )

with open(os.path.join(CONTENT_DIR, 'test_3ss.pptx'), 'rb') as rfile:
    st.download_button(
        label="Скачать пример pptx",
        data=BytesIO(rfile.read()),
        file_name="example.pptx",
    )

uploaded_file = st.file_uploader("Загрузите xlsx")

if uploaded_file:
    with open(os.path.join(CONTENT_DIR, 'file.xlsx'), 'wb') as wfile:
        wfile.write(uploaded_file.read())
        process(os.path.join(CONTENT_DIR, 'file.xlsx'), 'file.pptx')

    with open(os.path.join(CONTENT_DIR, 'file.pptx'), 'rb') as rfile:
        buffer = BytesIO(rfile.read())
    buffer.seek(0)
    st.download_button(
        label="Скачать pptx",
        data=buffer,
        file_name=f"pptx-{time.time()}.pptx",
    )
