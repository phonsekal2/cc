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
    #try:
        df = pd.read_excel(excelfile, sheet_name="Sheet1")
        st.write(df.to_dict(orient="records"))
        np_array = df["Nama"].to_numpy()
        #st.write(np_array)
        #st.write(list(np_array))
        # Iterate over each row in df and render word document
        #for record in df.to_dict(orient="records"):
        for r_idx, r_val in df.iterrows():
            st.write(df.iterrows())

            doc = DocxTemplate(f"pages/formatkuitansimonev2.docx")
            #total_hari = uang_harian_per_hari * lama_hari
            #total_malam = hotel_per_hari * (lama_hari - 1)
            #total = tiket + taksi_asal + taksi_daerah + total_hari + total_malam
            #total_spd = tiket + taksi_asal + taksi_daerah
            context = {
                "nama" : r_val['Nama'],
                "nip" : r_val['NIP'],
                "jabatan" : r_val['Jabatan'],
                "kegiatan" : r_val['nama_kegiatan'],
                "layanan" : r_val['dalam_rangka'],
                "lokasi" : r_val['lokasi'],
                "tujuan" : r_val['Kota_Tujuan'],
                "tanggal" : r_val['tanggal_kegiatan'],
                "nomor st" : r_val['no_st'],
                "tanggal st" : r_val['tanggal_st'],
                "kota tujuan" : r_val['Kota_Tujuan'],
                "asal tujuan" : "Jakarta " + "kota tujuan",
                "tiket" : r_val['pesawat'],
                "taksi asal" : r_val['taksi_asal'],
                "taksi daerah" : r_val['taksi_daerah'],
                "hari" : r_val['lama_hari'],
                "hari rp" : r_val['uang_harian_per_hari'],
                #"total hari" : (int("hari")) * (int("hari rp")),
                "malam" : r_val['lama_hari'],
                "malam rp" : r_val['hotel_per_hari'],
                #"total malam" : ("malam") * ("malam rp"),
                #"total" : "tiket" + "taksi_asal" + "taksi_daerah" + "total_hari" + "total_malam",
                #"terbilang" : num2words(int(total), lang='id').title() + " Rupiah",
               # "total spd" : total_spd,
               # "bulan" : bulan,
               # "hari terbilang" : num2words(int(total_hari)),
               # "tanggal berangkat" : tanggal_berangkat,
              #  "tanggal kembali" : tanggal_kembali,
               # "mak" : mak



            }
            st.write(context)
            doc.render(context)
            output_path = f"pages/OUTPUT/{context['Nama']}.docx"
            doc.save(output_path)
       # button = st.button("a")
        a = st.success("🎉 File kuitansi telah selesai dibuat, silahkan unduh")
        b = st.success("")
        with b:
            #st.write(df)
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