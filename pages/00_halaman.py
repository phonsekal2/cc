import streamlit as st
from pathlib import Path
import pandas as pd  # pip install pandas openpyxl
from docxtpl import DocxTemplate  # pip install docxtpl
from docx import Document
from docxcompose.composer import Composer
import numpy as np
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
    #try:
        # Convert Excel sheet to pandas dataframe
        df = pd.read_excel(excelfile, sheet_name="Sheet1")
        #st.write(df.to_dict(orient="records"))
        #st.write(df.to_dict(orient="list"))
        #st.write(df.to_dict(orient="series"))

        # Keep only date part YYYY-MM-DD (not the time)
        df["TODAY"] = pd.to_datetime(df["TODAY"]).dt.date
        df["TODAY_IN_ONE_WEEK"] = pd.to_datetime(df["TODAY_IN_ONE_WEEK"]).dt.date
        np_array = df["VENDOR"].to_numpy()
        #st.write(list(np_array))
        # Iterate over each row in df and render word document
        for record in df.to_dict(orient="records"):
            doc = DocxTemplate(f"pages/vendor-contract.docx")
            doc.render(record)
            output_path = f"pages/OUTPUT/{record['VENDOR']}.docx"
            doc.save(output_path)
        st.success("🎉 File kuitansi telah selesai dibuat, lanjutkan unggah lagi file nominatif untuk mengunduh kuitansi")
    #except:
       # st.warning("Unggah dulu file nominatif")   
form2 = st.form("Proses gabung kuitansi")
with form2:
    excelfile2 = st.file_uploader("Upload File Excel")
    submit2 = st.form_submit_button("Proses")
    

if submit2:
    #try:
            df = pd.read_excel(excelfile2, sheet_name="Sheet1")
            np_array = df["VENDOR"].to_numpy()
            files2 = list("pages/OUTPUT/" + (np_array) + ".docx")
            #files2 = files
            #st.write(files2)
            composed = f"pages/gabung.docx"
            result = Document(files2[0])
            result.add_page_break()
            composer = Composer(result)
            for i in range(1, len(files2)):
                doc2 = Document(files2[i])
                if i != len(files2) -1:
                    doc2.add_page_break()
                composer.append(doc2)
            composer.save(composed)
            with open(composed, "rb") as file:
                st.success("🎉 File kuitansi telah selesai dibuat")
                st.download_button(
                    label = "⬇️ Download File",
                    data=file,
                    file_name="kuitansi.docx",
                    mime="application/octet-stream",
                    key="10000009"
                )
   # except:
      #  st.warning("Unggah file dulu")