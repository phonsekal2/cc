import streamlit as st
import pandas as pd
from docxtpl import DocxTemplate
from num2words import num2words
from datetime import datetime
def kalender_indo(value):
    a = (value).strftime("%d %B %Y")
    kal = a.replace("January", "Januari").replace("February", "Februari").replace("March", "Maret").replace("May", "Mei").replace("June", "Juni").replace("July", "Juli").replace("August", "Agustus").replace("October", "Oktober").replace("December", "Desember")
    return kal

def transform_to_rupiah_format(value):
    str_value = str(value)
    separate_decimal = str_value.split(".")
    after_decimal = separate_decimal[0]
    before_decimal = separate_decimal[1]
    reverse = after_decimal[::-1]
    temp_reverse_value = ""
    for index, val in enumerate(reverse):
        if (index + 1) % 3 == 0 and index + 1 != len(reverse):
            temp_reverse_value = temp_reverse_value + val + "."
        else:
            temp_reverse_value = temp_reverse_value + val
    temp_result = temp_reverse_value[::-1]
    return "Rp" + temp_result + ",-" 
    #+ "," + before_decimal ##buat koma di belakang

def rupiah_strip(value):
    a = str(transform_to_rupiah_format(float(value)))
    ubah = a.replace("Rp0,-", "Rp-")
    return ubah
    
form = st.form("Upload file")
with form:
    excelfile = st.file_uploader("Unggah file nominatif")
    submit = st.form_submit_button("Proses")
if submit:
    df2 = pd.read_excel(excelfile, sheet_name="Sheet2")
    for r_idx, r_val in df2.iterrows():
        if (r_val['asal'] == 'jambi'):
            doctemp = f"pages/templates/formatjambi.docx"
        elif (r_val['asal'] == 'jakarta'):
            doctemp = f"pages/templates/formatjakarta.docx"
        doc = DocxTemplate(doctemp)
        context = {
            'name' : r_val['nama'],
            'surname' : r_val['nama_panggilan'],
            'from' : r_val['asal'],
            'age2' :str(r_val['umur']) + "Rupiah",
            'age11' : num2words(int(r_val['umur']), lang='id').title() + " Rupiah",
            'age' : transform_to_rupiah_format(float(r_val['umur'])),
            'age1' : r_val['umur'] * r_val['umur'],
            'date1' : r_val['tanggal'].strftime("%d %B %Y"),
            'date' : kalender_indo(r_val['tanggal']),
            'rupiah' : rupiah_strip((r_val['rupiah'])),                         
                }
        st.write(context)
        doc.render(context)
        output_path = f"pages/OUTPUT/{context['name']}.docx"
        doc.save(output_path)
        a = st.success("🎉 File kuitansi telah selesai dibuat, silahkan unduh")
        b = st.success("")
