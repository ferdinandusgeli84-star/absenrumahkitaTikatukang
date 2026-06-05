import streamlit as st
import pandas as pd
from datetime import datetime
import io

# Set konfigurasi halaman
st.set_page_config(page_title="Aplikasi Absensi Guru", layout="wide")

st.title("📋 Aplikasi Absensi Guru")
st.write("Isi absensi harian guru di bawah ini berdasarkan format template.")

# Input Tanggal
tanggal_pilih = st.date_input("Pilih Tanggal Absensi", datetime.now())
format_tanggal = tanggal_pilih.strftime("%A, %d-%m-%Y")

st.info(f"Mencatat Absensi untuk tanggal: **{format_tanggal}**")

# Daftar Guru dari file "Untitled spreadsheet (1).xlsx"
daftar_guru = [
    'Sabililah Pata Ola, S.Pd.',
    'FATMAWATI KASIM, S.S.', 
    'Ferdinandus Payong Geli, S.Pd.',
    'Agnes Beata Lewomuda, S.Pd.', 
    'Anania Kartini Beni, S.Pd.', 
    'Flasidius Henimus Ngaga, S.Pd.', 
    'Lintiana Bte Serilus, S.Pd.', 
    'Marselinus Dalo Wara, S.Pd.', 
    'Regina R, S.Pd.', 
    'Rikardus Kopong Boli, S.Pd.', 
    'Rusini Usman, S.Pd.', 
    'Rusmini Saputri, S.Pd.', 
    'Veronika Lepan Buran Keban, S.Pd.', 
    'Visensia Selviana Horowura, S.Pd.', 
    'Yohanes Juang Saso Assan, S.Pd.'
    'Yohanes Maria Vianey, S.E.', 
    'Kristina Setya Oreng Ola, S.M.', 
    'Junior Timu Meong.'
]

data_absensi = []
st.markdown("---")

# Form Input Absensi per Guru
for i, guru in enumerate(daftar_guru):
    st.subheader(f"👤 {guru}")
    col1, col2, col3 = st.columns([3, 4, 5])
    
    with col1:
        status = st.radio(
            f"Status Kehadiran ({guru})", 
            ["Hadir", "Sakit (S)", "Izin (I)", "Alpha (A)"], 
            key=f"status_{i}"
        )
        
    waktu_datang = ""
    waktu_pulang = ""
    s, i_stat, a = "", "", ""
    
    if status == "Hadir":
        with col2:
            waktu_datang = st.text_input("Waktu Datang", value="07:00", key=f"datang_{i}")
            waktu_pulang = st.text_input("Waktu Pulang", value="14:00", key=f"pulang_{i}")
    elif status == "Sakit (S)":
        s = "✓"
    elif status == "Izin (I)":
        i_stat = "✓"
    elif status == "Alpha (A)":
        a = "✓"
        
    with col3:
        keterangan = st.text_input("Keterangan Tambahan", value="-", key=f"ket_{i}")
        
    st.markdown("---")
    
    data_absensi.append([guru, waktu_datang, waktu_pulang, s, i_stat, a, keterangan])

# Tombol Cetak Excel
if st.button("Proses Data Absensi Hari Ini", type="primary"):
    # Rekonstruksi format agar persis seperti "Untitled spreadsheet (1).xlsx"
    header_data = [
        ["HARI DAN TANGGAL :", format_tanggal, "", "", "", "", ""],
        ["NAMA GURU", "HADIR", "", "TIDAK HADIR", "", "", "KET"],
        ["", "WAKTU DATANG", "WAKTU PULANG", "S", "I", "A", ""]
    ]
    
    for row in data_absensi:
        header_data.append(row)
        
    df_output = pd.DataFrame(header_data)
    
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        df_output.to_excel(writer, index=False, header=False, sheet_name='Sheet1')
        workbook  = writer.book
        worksheet = writer.sheets['Sheet1']
        
        # Atur lebar kolom biar rapi
        worksheet.set_column('A:A', 30)
        worksheet.set_column('B:G', 15)
        
    excel_data = output.getvalue()
    
    st.success("✅ Data berhasil diproses!")
    st.download_button(
        label="📥 Unduh File Excel Hasil Absensi",
        data=excel_data,
        file_name=f"Absensi_{tanggal_pilih.strftime('%Y%m%d')}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
