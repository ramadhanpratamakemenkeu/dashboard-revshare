import streamlit as st
import pandas as pd
import sqlite3
import plotly.express as px
import numpy as np
import streamlit.components.v1 as components
import io
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN
import os
import gzip
import shutil

# ==========================================
# 1. KONFIGURASI HALAMAN & UI/UX (Dark Mode)
# ==========================================
st.set_page_config(page_title="RevShare 1.0", layout="wide", page_icon="🚀")

# Custom CSS untuk mempercantik UI
st.markdown("""
<style>
    @import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;600;700&display=swap');
    
    html, body, [class*="css"] {
        font-family: 'Inter', sans-serif;
    }
    h1, h2, h3 {
        font-weight: 700;
        color: #ffffff;
    }
    .stTabs [data-baseweb="tab-list"] button {
        font-size: 16px;
        font-weight: 600;
        padding: 10px 20px;
        border-radius: 10px 10px 0 0;
    }
    .stButton>button {
        border-radius: 8px;
        font-weight: 600;
        transition: all 0.3s ease;
    }
    .stButton>button:hover {
        transform: scale(1.02);
    }
</style>
""", unsafe_allow_html=True)

st.title("🚀 RevShare 1.0")

# ==========================================
# 2. OTOMATIS EKSTRAK DATABASE & UPLOAD CSV
# ==========================================
DB = 'database_dbh.db'

# Ekstrak .gz jika .db belum ada
if not os.path.exists(DB):
    if os.path.exists(DB + '.gz'):
        with gzip.open(DB + '.gz', 'rb') as f_in:
            with open(DB, 'wb') as f_out:
                shutil.copyfileobj(f_in, f_out)
    else:
        st.error("File database tidak ditemukan!")

# Inisialisasi session state untuk data
if 'data_updated' not in st.session_state:
    @st.cache_data(show_spinner="Memuat data...")
    def load_data():
        conn = sqlite3.connect(DB)
        df = pd.read_sql_query("SELECT * FROM database_utama", conn)
        conn.close()
        df['tahun'] = pd.to_numeric(df['tahun'], errors='coerce').fillna(0).astype(int)
        return df
    st.session_state['data_updated'] = load_data()

df = st.session_state['data_updated']

# --- SIDEBAR: FITUR UPLOAD ---
st.sidebar.header("📤 Upload Data Baru")
uploaded_file = st.sidebar.file_uploader("Pilih file CSV atau Excel", type=['csv', 'xlsx'])

if uploaded_file is not None:
    if st.sidebar.button("Proses Upload"):
        try:
            if uploaded_file.name.endswith('.csv'):
                new_data = pd.read_csv(uploaded_file)
            else:
                new_data = pd.read_excel(uploaded_file)
            
            # Validasi kolom
            required_cols = ['siklus', 'pemda', 'uraian4', 'nilai', 'tanggal', 'tahun']
            if not all(col in new_data.columns for col in required_cols):
                st.sidebar.error(f"Kolom tidak lengkap! Harus ada: {', '.join(required_cols)}")
            else:
                # Konversi tipe data
                new_data['nilai'] = pd.to_numeric(new_data['nilai'], errors='coerce')
                new_data['tahun'] = pd.to_numeric(new_data['tahun'], errors='coerce').fillna(0).astype(int)
                
                # Gabungkan (Append)
                df_combined = pd.concat([df, new_data], ignore_index=True)
                # Hapus duplikat jika ada (berdasarkan siklus, pemda, uraian4, tahun)
                df_combined = df_combined.drop_duplicates(subset=['siklus', 'pemda', 'uraian4', 'tahun'], keep='last')
                
                st.session_state['data_updated'] = df_combined
                df = df_combined
                
                # Update DB SQLite sementara
                conn = sqlite3.connect(DB)
                df.to_sql('database_utama', conn, if_exists='replace', index=False)
                conn.close()
                
                st.sidebar.success("✅ Data berhasil ditambahkan!")
                st.sidebar.info("💡 Agar permanen, unduh file DB di bawah dan ganti file .db.gz di GitHub.")
        except Exception as e:
            st.sidebar.error(f"Error: {e}")

# Tombol unduh database terbaru (jika ada update)
if os.path.exists(DB):
    with open(DB, 'rb') as f:
        st.sidebar.download_button(
            label="⬇️ Unduh Database Terbaru (.db)",
            data=f,
            file_name="database_dbh_terbaru.db",
            mime="application/x-sqlite3",
            help="Upload file ini ke GitHub (ganti yang lama) agar perubahan permanen."
        )

# ==========================================
# 3. TABS & KONTEN
# ==========================================
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
    fig = px.bar(agg, x='nilai', y=grup, orientation='h', template='plotly_dark')
    fig.update_layout(paper_bgcolor='rgba(0,0,0,0)', plot_bgcolor='rgba(0,0,0,0)')
    st.plotly_chart(fig, use_container_width=True)

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

        # Tampilan HTML Tabel
        css = "<style>.custom-table { width: 100%; border-collapse: collapse; font-family: 'Inter', sans-serif; font-size: 13px; background-color: white; color: black; } .custom-table th, .custom-table td { border: 1px solid #000; padding: 6px 8px; } .custom-table th { background-color: #d9e1f2; text-align: center; font-weight: 700; vertical-align: middle; } .custom-table td { text-align: right; } .custom-table td.text-left { text-align: left; } .custom-table tr:nth-child(even) { background-color: #f2f2f2; } .total-row { background-color: #d9e1f2 !important; font-weight: 700; }</style>"
        
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

        # Tampilan HTML
        css = "<style>.custom-table { width: 100%; border-collapse: collapse; font-family: 'Inter', sans-serif; font-size: 11px; background-color: white; color: black; } .custom-table th, .custom-table td { border: 1px solid #000; padding: 4px 6px; } .custom-table th { background-color: #d9e1f2; text-align: center; font-weight: 700; vertical-align: middle; } .custom-table td { text-align: right; } .custom-table td.text-left { text-align: left; } .custom-table tr:nth-child(even) { background-color: #f2f2f2; } .total-row { background-color: #d9e1f2 !important; font-weight: 700; }</style>"
        
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
        
        # --- PERBAIKAN EXPORT PPT ---
        st.markdown("### 📥 Export Laporan ke PowerPoint")
        if st.button("Buat Slide PowerPoint (.pptx)"):
            with st.spinner("Membuat file PPT yang rapi..."):
                prs = Presentation()
                # Set ukuran slide ke Widescreen (16:9)
                prs.slide_width = Inches(13.333)
                prs.slide_height = Inches(7.5)
                
                slide = prs.slides.add_slide(prs.slide_layouts[5]) # Title Only
                
                # Judul
                title = slide.shapes.title
                title.text = "ALOKASI DAN REALISASI DBH"
                title.text_frame.paragraphs[0].font.size = Pt(24)
                title.text_frame.paragraphs[0].font.bold = True
                title.text_frame.paragraphs[0].font.name = 'Arial'
                title.text_frame.paragraphs[0].alignment = PP_ALIGN.CENTER
                
                rows, cols = len(df_hasil) + 3, 1 + (len(list_tahun) * 4)
                left, top, width, height = Inches(0.2), Inches(1.2), Inches(12.9), Inches(5.8)
                
                shape = slide.shapes.add_table(rows, cols, left, top, width, height)
                table = shape.table
                
                # Atur lebar kolom proporsional
                table.columns[0].width = Inches(2.5)
                for i in range(1, cols):
                    table.columns[i].width = Inches((12.9 - 2.5) / (cols - 1))
                
                # Helper untuk mewarnai sel dan memberi border
                def style_cell(cell, text, is_header=False, is_total=False, align=PP_ALIGN.RIGHT, bg_color=None):
                    cell.text = str(text)
                    cell.text_frame.paragraphs[0].font.size = Pt(10)
                    cell.text_frame.paragraphs[0].font.name = 'Arial'
                    cell.text_frame.paragraphs[0].alignment = align
                    cell.vertical_anchor = 3 # MIDDLE
                    
                    if is_header or is_total:
                        cell.fill.solid()
                        cell.fill.fore_color.rgb = RGBColor(217, 225, 242) # Biru muda #D9E1F2
                        cell.text_frame.paragraphs[0].font.bold = True
                    elif bg_color:
                        cell.fill.solid()
                        cell.fill.fore_color.rgb = bg_color
                    else:
                        cell.fill.solid()
                        cell.fill.fore_color.rgb = RGBColor(255, 255, 255)

                # Header Row 1 & 2 & 3
                style_cell(table.cell(0, 0), "PEMERINTAH DAERAH", is_header=True, align=PP_ALIGN.CENTER)
                table.cell(0, 0).merge(table.cell(2, 0))
                
                col_idx = 1
                for thn in list_tahun:
                    style_cell(table.cell(0, col_idx), str(thn), is_header=True, align=PP_ALIGN.CENTER)
                    table.cell(0, col_idx).merge(table.cell(0, col_idx + 3))
                    
                    style_cell(table.cell(1, col_idx), "Alokasi", is_header=True, align=PP_ALIGN.CENTER)
                    table.cell(1, col_idx).merge(table.cell(2, col_idx))
                    
                    style_cell(table.cell(1, col_idx+1), "Penyaluran", is_header=True, align=PP_ALIGN.CENTER)
                    table.cell(1, col_idx+1).merge(table.cell(1, col_idx+2))
                    
                    style_cell(table.cell(2, col_idx+1), "Reguler", is_header=True, align=PP_ALIGN.CENTER)
                    style_cell(table.cell(2, col_idx+2), "%", is_header=True, align=PP_ALIGN.CENTER)
                    
                    style_cell(table.cell(1, col_idx+3), "Kurang Bayar", is_header=True, align=PP_ALIGN.CENTER)
                    table.cell(1, col_idx+3).merge(table.cell(2, col_idx+3))
                    col_idx += 4
                
                # Isi Data
                for r_idx, row_data in df_hasil.iterrows():
                    excel_row = r_idx + 3
                    is_total = row_data['pemda'] == 'JUMLAH'
                    is_prov = "Provinsi" in str(row_data['pemda'])
                    bg_color = RGBColor(230, 240, 255) if is_prov and not is_total else None
                    
                    style_cell(table.cell(excel_row, 0), row_data['pemda'], is_total=is_total, align=PP_ALIGN.LEFT, bg_color=bg_color)
                    
                    col_idx = 1
                    for thn in list_tahun:
                        style_cell(table.cell(excel_row, col_idx), f"{row_data[f'A_{thn}']:,.0f}", is_total=is_total, bg_color=bg_color)
                        style_cell(table.cell(excel_row, col_idx+1), f"{row_data[f'P_{thn}']:,.0f}", is_total=is_total, bg_color=bg_color)
                        style_cell(table.cell(excel_row, col_idx+2), f"{row_data[f'Pct_{thn}']:.2f}%", is_total=is_total, bg_color=bg_color)
                        style_cell(table.cell(excel_row, col_idx+3), f"{row_data[f'KB_{thn}']:,.0f}", is_total=is_total, bg_color=bg_color)
                        col_idx += 4
                
                ppt_stream = io.BytesIO()
                prs.save(ppt_stream)
                ppt_stream.seek(0)
                
                st.success("✅ File PPT berhasil dibuat dengan format yang rapi!")
                st.download_button("⬇️ Download File PPT", data=ppt_stream, file_name=f"Alokasi_Realisasi_{pilih_prov_ar}.pptx", mime="application/vnd.openxmlformats-officedocument.presentationml.presentation")

# ------------------------------------------
# TAB 4: DATA MENTAH
# ------------------------------------------
with tab4:
    st.subheader("Data Mentah")
    st.dataframe(df.head(10000), use_container_width=True)
