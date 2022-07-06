import streamlit as st
from pathlib import Path
import pandas as pd  # pip install pandas openpyxl
from docxtpl import DocxTemplate  # pip install docxtpl
from docx import Document
from docxcompose.composer import Composer
import numpy as np
from num2words import num2words
st.set_page_config(page_title='Dede Saputra App', page_icon = ":coffee:", layout = 'centered', initial_sidebar_state = 'auto')

form = st.form("Upload file")
with form:
    st.header("Proses pembuatan kuitansi")
    st.markdown(
    "**PERHATIAN: ** Untuk menggunakan aplikasi ini, silahkan unduh format nominatif[ di sini](https://docs.google.com/spreadsheets/d/10LDb3d-pbSRKuio8zY-HIS2kmP03ehhn/edit#gid=1302452135). Dimohon untuk mengisi nominatif sesuai dengan format file unduhan, tanpa nama kolom dan nama sheet. "
    )
    st.write("**Ini adalah format kuitansi perjalanan dinas ke daerah, jika ingin membuat kuitansi jenis lain silahkan klik menu**")

    excelfile = st.file_uploader("Unggah file nominatif")
    submit = st.form_submit_button("Proses")
if submit:
    df = pd.read_excel(excelfile, sheet_name="Sheet2")
    tiket = df["tiket"]
    terkata = tiket
    currency = "Rp{:,.2f}".format(terkata.any())
    thousands_separator = "."
    fractional_separator = ","
    if thousands_separator == ".":
        main_currency, fractional_currency = currency.split(".")[0], currency.split(".")[1]
        new_main_currency = main_currency.replace(",", ".")
        currency = new_main_currency + fractional_separator + fractional_currency
    st.write(tiket)
    st.write(currency)

    

