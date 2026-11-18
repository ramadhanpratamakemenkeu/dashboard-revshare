import streamlit as st
import pandas as pd
import sqlite3
import plotly.express as px
import numpy as np
import streamlit.components.v1 as components
import io
from pptx import Presentation
from pptx.util import Inches, Pt
import os
import gzip
import shutil

# ==========================================
# KODE OTOMATIS EKSTRAK DATABASE (.gz ke .db)
# ==========================================
if not os.path.exists('database_dbh.db'):
    if os.path.exists('database_dbh.db.gz'):
        with gzip.open('database_dbh.db.gz', 'rb') as f_in:
            with open('database_dbh.db', 'wb') as f_out:
                shutil.copyfileobj(f_in, f_out)
    else:
        st.error("File database tidak ditemukan! Pastikan database_dbh.db.gz sudah diunggah.")

# ==========================================
# APLIKASI UTAMA
# ==========================================
st.set_page_config(page_title="Dashboard DBH", layout="wide")
st.title("📊 Dashboard DBH")

DB = 'database_dbh.db'

@st.cache_data(show_spinner="Memuat data...")
def load_data():
    cols = ['tahun','siklus','namaprovinsi','pemda', 'level0','keterangan', 'nilai']
    conn = sqlite3.connect(DB)
    df = pd.read_sql_query(f"SELECT {','.join(cols)} FROM database_utama", conn)
    conn.close()
    df['tahun'] = pd.to_numeric(df['tahun'], errors='coerce').fillna(0).astype(int)
    return df

df = load_data()

tab1, tab2, tab3, tab4 = st.tabs([
    "📈 Visualisasi Chart", 
    "📄 Laporan Tabel (KB/LB)", 
    "🏦 Alokasi & Realisasi",
    "📋 Data Mentah"
])

# ------------------------------------------
# TAB 1: CHART
# ------------------------------------------
with tab1:
    st.subheader("Visualisasi Chart")
    grup = st.selectbox("Kelompokkan per", ['level0','level1','namaprovinsi'], key='grup_chart')
    agg = df.groupby(grup, as_index=False)['nilai'].sum().sort_values('nilai', ascending=False).head(15)
    st.plotly_chart(px.bar(agg, x='nilai', y=grup, orientation='h'), use_container_width=True)

# ------------------------------------------
# TAB 2: TABEL KBLB
# ------------------------------------------
with tab2:
    st.subheader("Laporan Sisa KB & LB DBH")
    
    semua_prov = sorted(df['namaprovinsi'].dropna().unique())
    idx_jateng = semua_prov.index("Jawa Tengah") if "Jawa Tengah" in semua_prov else 0
    pilih_prov_kblb = st.selectbox("Pilih Provinsi untuk Laporan KB/LB:", semua_prov, index=idx_jateng, key='prov_kblb')
    
    df_prov = df[(df['namaprovinsi'] == pilih_prov_kblb) & (df['level0'] == 'DBH')].copy()
    
    if df_prov.empty:
        st.warning(f"Data DBH tidak ditemukan untuk provinsi {pilih_prov_kblb}.")
    else:
        with st.expander("⚙️ Pengaturan Teks Transaksi KB/LB"):
            opsi_ket = sorted(df_prov['keterangan'].dropna().unique())
            
            def_kb23 = next((k for k in opsi_ket if 'Sisa KB s.d. TA 2023 KMK 198' in k), opsi_ket[0] if opsi_ket else "")
            def_kb24 = next((k for k in opsi_ket if 'Penyaluran KB TA 2024 KMK 198' in k), opsi_ket[0] if opsi_ket else "")
            def_lb23 = next((k for k in opsi_ket if 'Sisa LB s.d. TA 2023 KMK 44' in k), opsi_ket[0] if opsi_ket else "")
            def_lb24 = next((k for k in opsi_ket if 'Penetapan LB TA 2024' in k or 'Sisa LB s.d. TA 2024 KMK 198' in k), opsi_ket[0] if opsi_ket else "")

            c1, c2, c3, c4 = st.columns(4)
            ket_kb23 = c1.selectbox("SiSA KB s.d. TA 2023", opsi_ket, index=opsi_ket.index(def_kb23) if def_kb23 in opsi_ket else 0)
            ket_kb24 = c2.selectbox("KB TA 2024", opsi_ket, index=opsi_ket.index(def_kb24) if def_kb24 in opsi_ket else 0)
            ket_lb23 = c3.selectbox("SiSA LB s.d. TA 2023", opsi_ket, index=opsi_ket.index(def_lb23) if def_lb23 in opsi_ket else 0)
            ket_lb24 = c4.selectbox("LB TA 2024", opsi_ket, index=opsi_ket.index(def_lb24) if def_lb24 in opsi_ket else 0)

        def get_sum(keterangan_target):
            if not keterangan_target: return pd.Series(dtype=float)
            return df_prov[df_prov['keterangan'] == keterangan_target].groupby('pemda')['nilai'].sum()

        s_kb23, s_kb24 = get_sum(ket_kb23), get_sum(ket_kb24)
        s_lb23, s_lb24 = get_sum(ket_lb23), get_sum(ket_lb24)

        hasil = pd.DataFrame({'KB_2023': s_kb23, 'KB_2024': s_kb24, 'LB_2023': s_lb23, 'LB_2024': s_lb24}).fillna(0)

        prov_name = f"Provinsi {pilih_prov_kblb}" if "Provinsi" not in pilih_prov_kblb else pilih_prov_kblb
        if prov_name not in hasil.index:
            hasil.loc[prov_name] = hasil.sum()

        hasil = hasil.reset_index().rename(columns={'index': 'pemda'})
        hasil['is_prov'] = hasil['pemda'].str.contains('Provinsi', case=False, na=False)
        hasil = hasil.sort_values(by=['is_prov', 'pemda'], ascending=[False, True]).drop(columns=['is_prov'])

        hasil['KB_Total'] = hasil['KB_2023'] + hasil['KB_2024']
        hasil['LB_Total'] = hasil['LB_2023'] + hasil['LB_2024']
        hasil.insert(0, 'No', range(1, len(hasil) + 1))

        css = "<style>.custom-table { width: 100%; border-collapse: collapse; font-family: sans-serif; font-size: 13px; background-color: white; color: black; } .custom-table th, .custom-table td { border: 1px solid #000; padding: 6px 8px; } .custom-table th { background-color: #d9e1f2; text-align: center; font-weight: bold; vertical-align: middle; } .custom-table td { text-align: right; } .custom-table td.text-left { text-align: left; } .custom-table tr:nth-child(even) { background-color: #f2f2f2; } .total-row { background-color: #d9e1f2 !important; font-weight: bold; }</style>"
        
        html = css + "<table class='custom-table'><thead>"
        html += "<tr><th rowspan='3'>No</th><th rowspan='3'>Pemerintah Daerah</th><th colspan='3'>Kurang Bayar</th><th colspan='3'>Lebih Bayar</th></tr>"
        html += "<tr><th>SiSA KB s.d. TA 2023</th><th>KB TA 2024</th><th>KB s.d. TA 2024</th><th>SiSA LB s.d. TA 2023</th><th>LB TA 2024</th><th>LB s.d. TA 2024</th></tr></thead><tbody>"
        
        for _, row in hasil.iterrows():
            html += f"<tr><td class='text-left'>{row['No']}</td>"
            style_prov = "style='font-weight:bold; background-color:#e6f0ff;'" if "Provinsi" in str(row['pemda']) else ""
            html += f"<td class='text-left' {style_prov}>{row['pemda']}</td>"
            html += f"<td>{row['KB_2023']:,.0f}</td><td>{row['KB_2024']:,.0f}</td><td>{row['KB_Total']:,.0f}</td>"
            html += f"<td>{row['LB_2023']:,.0f}</td><td>{row['LB_2024']:,.0f}</td><td>{row['LB_Total']:,.0f}</td></tr>"
            
        html += f"<tr class='total-row'><td colspan='2' class='text-left' style='text-align: center;'>JUMLAH</td>"
        html += f"<td>{hasil['KB_2023'].sum():,.0f}</td><td>{hasil['KB_2024'].sum():,.0f}</td><td>{hasil['KB_Total'].sum():,.0f}</td>"
        html += f"<td>{hasil['LB_2023'].sum():,.0f}</td><td>{hasil['LB_2024'].sum():,.0f}</td><td>{hasil['LB_Total'].sum():,.0f}</td></tr></tbody></table>"
        
        components.html(html, height=800, scrolling=True)
        st.download_button("⬇️ Download Tabel KB/LB (CSV)", data=hasil.to_csv(index=False).encode('utf-8'), file_name=f"KB_LB_{pilih_prov_kblb}.csv", mime="text/csv")

# ------------------------------------------
# TAB 3: ALOKASI & REALISASI + PPT
# ------------------------------------------
with tab3:
    st.subheader("Laporan Alokasi dan Realisasi DBH")
    
    semua_prov = sorted(df['namaprovinsi'].dropna().unique())
    idx_jateng = semua_prov.index("Jawa Tengah") if "Jawa Tengah" in semua_prov else 0
    pilih_prov_ar = st.selectbox("Pilih Provinsi untuk Laporan Alokasi:", semua_prov, index=idx_jateng, key='prov_ar')
    
    df_prov = df[(df['namaprovinsi'] == pilih_prov_ar) & (df['level0'] == 'DBH')].copy()
    
    list_tahun = [2025, 2026, 2027]
    
    with st.expander("⚙️ Pengaturan Teks Transaksi Alokasi/Realisasi"):
        opsi_ket = sorted(df_prov['keterangan'].dropna().unique())
        mapping_ket = {}
        cols_ui = st.columns(3)
        
        for i, thn in enumerate(list_tahun):
            with cols_ui[i]:
                st.markdown(f"**Tahun {thn}**")
                def_alokasi = next((k for k in opsi_ket if f'Alokasi DBH pada TA {thn}' in k), "")
                def_penyaluran = next((k for k in opsi_ket if f'Penyaluran DBH Reguler TA {thn}' in k), "")
                def_kb = next((k for k in opsi_ket if f'Sisa KB s.d. TA {thn}' in k or f'Penetapan KB TA {thn}' in k), "")
                
                idx_alokasi = opsi_ket.index(def_alokasi) if def_alokasi in opsi_ket else 0
                idx_penyaluran = opsi_ket.index(def_penyaluran) if def_penyaluran in opsi_ket else 0
                idx_kb = opsi_ket.index(def_kb) if def_kb in opsi_ket else 0

                mapping_ket[thn] = {
                    'alokasi': st.selectbox(f"Alokasi {thn}", opsi_ket, index=idx_alokasi, key=f"a{thn}"),
                    'penyaluran': st.selectbox(f"Penyaluran {thn}", opsi_ket, index=idx_penyaluran, key=f"p{thn}"),
                    'kb': st.selectbox(f"Kurang Bayar {thn}", opsi_ket, index=idx_kb, key=f"k{thn}"),
                }

    if df_prov.empty:
        st.warning("Data tidak ditemukan.")
    else:
        def get_val(pemda, ket, thn):
            if not ket: return 0
            val = df_prov[(df_prov['pemda'] == pemda) & (df_prov['keterangan'] == ket) & (df_prov['tahun'] == thn)]['nilai'].sum()
            return val if not pd.isna(val) else 0

        list_pemda = sorted(df_prov['pemda'].dropna().unique())
        prov_name = next((p for p in list_pemda if 'Provinsi' in p), None)
        if prov_name:
            list_pemda.remove(prov_name)
            list_pemda.insert(0, prov_name)

        data_rows = []
        for p in list_pemda:
            row = {'pemda': p}
            for thn in list_tahun:
                alok = get_val(p, mapping_ket[thn]['alokasi'], thn)
                penyal = get_val(p, mapping_ket[thn]['penyaluran'], thn)
                kb = get_val(p, mapping_ket[thn]['kb'], thn)
                pct = (penyal / alok * 100) if alok > 0 else 0
                
                row[f'A_{thn}'], row[f'P_{thn}'], row[f'Pct_{thn}'], row[f'KB_{thn}'] = alok, penyal, pct, kb
            data_rows.append(row)
            
        df_hasil = pd.DataFrame(data_rows)
        
        total_row = {'pemda': 'JUMLAH'}
        for thn in list_tahun:
            total_row[f'A_{thn}'] = df_hasil[f'A_{thn}'].sum()
            total_row[f'P_{thn}'] = df_hasil[f'P_{thn}'].sum()
            total_row[f'Pct_{thn}'] = (total_row[f'P_{thn}'] / total_row[f'A_{thn}'] * 100) if total_row[f'A_{thn}'] > 0 else 0
            total_row[f'KB_{thn}'] = df_hasil[f'KB_{thn}'].sum()
        
        df_hasil = pd.concat([df_hasil, pd.DataFrame([total_row])], ignore_index=True)

        css = "<style>.custom-table { width: 100%; border-collapse: collapse; font-family: sans-serif; font-size: 11px; background-color: white; color: black; } .custom-table th, .custom-table td { border: 1px solid #000; padding: 4px 6px; } .custom-table th { background-color: #d9e1f2; text-align: center; font-weight: bold; vertical-align: middle; } .custom-table td { text-align: right; } .custom-table td.text-left { text-align: left; } .custom-table tr:nth-child(even) { background-color: #f2f2f2; } .total-row { background-color: #d9e1f2 !important; font-weight: bold; }</style>"
        
        html = css + "<table class='custom-table'><thead>"
        html += "<tr><th rowspan='3'>PEMERINTAH DAERAH</th>"
        for thn in list_tahun: html += f"<th colspan='4'>{thn}</th>"
        html += "</tr><tr>"
        for _ in list_tahun: html += "<th rowspan='2'>Alokasi</th><th colspan='2'>Penyaluran</th><th rowspan='2'>Kurang Bayar</th>"
        html += "</tr><tr>"
        for _ in list_tahun: html += "<th>Reguler</th><th>%</th>"
        html += "</tr></thead><tbody>"

        for _, row in df_hasil.iterrows():
            is_total = row['pemda'] == 'JUMLAH'
            tr_class = "class='total-row'" if is_total else ""
            style_prov = "style='font-weight:bold; background-color:#e6f0ff;'" if "Provinsi" in str(row['pemda']) else ""
            
            html += f"<tr {tr_class}><td class='text-left' {style_prov}>{row['pemda']}</td>"
            for thn in list_tahun:
                html += f"<td>{row[f'A_{thn}']:,.0f}</td><td>{row[f'P_{thn}']:,.0f}</td><td>{row[f'Pct_{thn}']:.2f}%</td><td>{row[f'KB_{thn}']:,.0f}</td>"
            html += "</tr>"
        html += "</tbody></table>"
        
        components.html(html, height=600, scrolling=True)
        
        st.markdown("### 📥 Export Laporan ke PowerPoint")
        if st.button("Buat Slide PowerPoint (.pptx)"):
            with st.spinner("Membuat file PPT..."):
                prs = Presentation()
                slide = prs.slides.add_slide(prs.slide_layouts[5])
                
                title = slide.shapes.title
                title.text = "ALOKASI DAN REALISASI DBH"
                title.text_frame.paragraphs[0].font.size = Pt(18)
                title.text_frame.paragraphs[0].font.bold = True
                
                rows, cols = len(df_hasil) + 3, 1 + (len(list_tahun) * 4)
                shape = slide.shapes.add_table(rows, cols, Inches(0.2), Inches(1.2), Inches(13.0), Inches(5.5))
                table = shape.table
                
                for r in range(rows):
                    for c in range(cols):
                        cell = table.cell(r, c)
                        cell.text_frame.paragraphs[0].font.size = Pt(9)
                        cell.margin_left, cell.margin_right, cell.margin_top, cell.margin_bottom = Inches(0.02), Inches(0.02), Inches(0.02), Inches(0.02)
                
                table.cell(0, 0).text = "PEMERINTAH DAERAH"
                table.cell(0, 0).merge(table.cell(2, 0))
                
                col_idx = 1
                for thn in list_tahun:
                    table.cell(0, col_idx).text = str(thn)
                    table.cell(0, col_idx).merge(table.cell(0, col_idx + 3))
                    
                    table.cell(1, col_idx).text = "Alokasi"
                    table.cell(1, col_idx).merge(table.cell(2, col_idx))
                    
                    table.cell(1, col_idx+1).text = "Penyaluran"
                    table.cell(1, col_idx+1).merge(table.cell(1, col_idx+2))
                    
                    table.cell(2, col_idx+1).text = "Reguler"
                    table.cell(2, col_idx+2).text = "%"
                    
                    table.cell(1, col_idx+3).text = "Kurang Bayar"
                    table.cell(1, col_idx+3).merge(table.cell(2, col_idx+3))
                    col_idx += 4
                
                for r_idx, row_data in df_hasil.iterrows():
                    excel_row = r_idx + 3
                    table.cell(excel_row, 0).text = str(row_data['pemda'])
                    col_idx = 1
                    for thn in list_tahun:
                        table.cell(excel_row, col_idx).text = f"{row_data[f'A_{thn}']:,.0f}"
                        table.cell(excel_row, col_idx+1).text = f"{row_data[f'P_{thn}']:,.0f}"
                        table.cell(excel_row, col_idx+2).text = f"{row_data[f'Pct_{thn}']:.2f}%"
                        table.cell(excel_row, col_idx+3).text = f"{row_data[f'KB_{thn}']:,.0f}"
                        col_idx += 4
                
                ppt_stream = io.BytesIO()
                prs.save(ppt_stream)
                ppt_stream.seek(0)
                
                st.success("✅ File PPT berhasil dibuat!")
                st.download_button("⬇️ Download File PPT", data=ppt_stream, file_name=f"Alokasi_Realisasi_{pilih_prov_ar}.pptx", mime="application/vnd.openxmlformats-officedocument.presentationml.presentation")

# ------------------------------------------
# TAB 4: DATA MENTAH
# ------------------------------------------
with tab4:
    st.subheader("Data Mentah")
    st.dataframe(df.head(10000), use_container_width=True)
