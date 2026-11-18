import streamlit as st
import pandas as pd
import sqlite3
import plotly.express as px
import numpy as np
import streamlit.components.v1 as components
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN
from fpdf import FPDF
import os
import gzip
import shutil
import io

# ==========================================
# 1. KONFIGURASI HALAMAN & UI/UX
# ==========================================
st.set_page_config(page_title="RevShare 1.0", layout="wide", page_icon="🚀")

st.markdown("""
<style>
    @import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;600;700&display=swap');
    html, body, [class*="css"] { font-family: 'Inter', sans-serif; }
    h1, h2, h3 { font-weight: 700; color: #ffffff; }
    .stTabs [data-baseweb="tab-list"] button { font-size: 16px; font-weight: 600; padding: 10px 20px; border-radius: 10px 10px 0 0; }
    .stButton>button { border-radius: 8px; font-weight: 600; transition: all 0.3s ease; }
    .stButton>button:hover { transform: scale(1.02); }
</style>
""", unsafe_allow_html=True)

st.title("🚀 RevShare 1.0")

# ==========================================
# 2. DATABASE & UPLOAD
# ==========================================
DB = 'database_dbh.db'
if not os.path.exists(DB):
    if os.path.exists(DB + '.gz'):
        with gzip.open(DB + '.gz', 'rb') as f_in:
            with open(DB, 'wb') as f_out:
                shutil.copyfileobj(f_in, f_out)
    else:
        st.error("File database tidak ditemukan!")

if 'data_updated' not in st.session_state:
    @st.cache_data(show_spinner="Memuat data...")
    def load_data():
        conn = sqlite3.connect(DB)
        df = pd.read_sql_query("SELECT * FROM database_utama", conn)
        conn.close()
        
        required_cols = ['tahun', 'siklus', 'namaprovinsi', 'pemda', 'level0', 'level1', 'level2', 'level3', 'uraian4', 'keterangan', 'nilai']
        for col in required_cols:
            if col not in df.columns:
                df[col] = "" 
                
        df['tahun'] = pd.to_numeric(df['tahun'], errors='coerce').fillna(0).astype(int)
        df['nilai'] = pd.to_numeric(df['nilai'], errors='coerce').fillna(0)
        return df
    st.session_state['data_updated'] = load_data()

df = st.session_state['data_updated']

st.sidebar.header("📤 Upload Data Baru")
uploaded_file = st.sidebar.file_uploader("Pilih file CSV atau Excel", type=['csv', 'xlsx'])

if uploaded_file is not None:
    if st.sidebar.button("Proses Upload"):
        try:
            if uploaded_file.name.endswith('.csv'):
                new_data = pd.read_csv(uploaded_file)
            else:
                new_data = pd.read_excel(uploaded_file)
            
            required_cols = ['siklus', 'pemda', 'uraian4', 'nilai', 'tanggal', 'tahun']
            if not all(col in new_data.columns for col in required_cols):
                st.sidebar.error(f"Kolom tidak lengkap! Harus ada: {', '.join(required_cols)}")
            else:
                new_data['nilai'] = pd.to_numeric(new_data['nilai'], errors='coerce')
                new_data['tahun'] = pd.to_numeric(new_data['tahun'], errors='coerce').fillna(0).astype(int)
                
                df_combined = pd.concat([df, new_data], ignore_index=True)
                df_combined = df_combined.drop_duplicates(subset=['siklus', 'pemda', 'uraian4', 'tahun'], keep='last')
                
                st.session_state['data_updated'] = df_combined
                df = df_combined
                
                conn = sqlite3.connect(DB)
                df.to_sql('database_utama', conn, if_exists='replace', index=False)
                conn.close()
                
                st.sidebar.success("✅ Data berhasil ditambahkan!")
                st.sidebar.info("💡 Agar permanen, unduh file DB di bawah dan ganti file .db.gz di GitHub.")
        except Exception as e:
            st.sidebar.error(f"Error: {e}")

if os.path.exists(DB):
    with open(DB, 'rb') as f:
        st.sidebar.download_button(
            label="⬇️ Unduh Database Terbaru (.db)",
            data=f,
            file_name="database_dbh_terbaru.db",
            mime="application/x-sqlite3"
        )

# ==========================================
# 3. FUNGSI BANTUAN EKSPOR
# ==========================================
def generate_pdf_kblb(df_pivot, pemda_name):
    pdf = FPDF(orientation='L', unit='mm', format='A4')
    pdf.add_page()
    
    pdf.set_font("Helvetica", style="B", size=16)
    pdf.cell(0, 10, f"SISA KB dan SISA LB DBH {pemda_name.upper()}", new_x="LMARGIN", new_y="NEXT", align="C")
    pdf.set_font("Helvetica", size=10)
    pdf.cell(0, 5, "(dalam rupiah)", new_x="LMARGIN", new_y="NEXT", align="R")
    pdf.ln(5)
    
    col_widths = [60, 30, 30, 30, 30, 30, 30]
    headers = ["Jenis DBH", "Sisa KB 2023", "KB 2024", "KB Total", "Sisa LB 2023", "LB 2024", "LB Total"]
    
    pdf.set_font("Helvetica", style="B", size=10)
    pdf.set_fill_color(217, 225, 242)
    for i, h in enumerate(headers):
        pdf.cell(col_widths[i], 10, h, border=1, align="C", fill=True)
    pdf.ln()
    
    pdf.set_font("Helvetica", size=9)
    for _, row in df_pivot.iterrows():
        pdf.cell(col_widths[0], 8, str(row['Jenis DBH']), border=1, align="L")
        pdf.cell(col_widths[1], 8, f"{row['KB 2023']:,.0f}", border=1, align="R")
        pdf.cell(col_widths[2], 8, f"{row['KB 2024']:,.0f}", border=1, align="R")
        pdf.cell(col_widths[3], 8, f"{row['KB Total']:,.0f}", border=1, align="R")
        pdf.cell(col_widths[4], 8, f"{row['LB 2023']:,.0f}", border=1, align="R")
        pdf.cell(col_widths[5], 8, f"{row['LB 2024']:,.0f}", border=1, align="R")
        pdf.cell(col_widths[6], 8, f"{row['LB Total']:,.0f}", border=1, align="R")
        pdf.ln()
        
    return bytes(pdf.output())

def generate_ppt_kblb_pemda(df_hasil, pemda_name):
    prs = Presentation()
    prs.slide_width = Inches(13.333)
    prs.slide_height = Inches(7.5)
    slide = prs.slides.add_slide(prs.slide_layouts[5])
    
    title = slide.shapes.title
    title.text = f"SISA KB dan SISA LB DBH {pemda_name.upper()}"
    title.text_frame.paragraphs[0].font.size = Pt(20)
    title.text_frame.paragraphs[0].font.bold = True
    title.text_frame.paragraphs[0].font.name = 'Arial'
    title.text_frame.paragraphs[0].alignment = PP_ALIGN.CENTER
    
    rows = len(df_hasil) + 2
    cols = 7
    shape = slide.shapes.add_table(rows, cols, Inches(0.5), Inches(1.5), Inches(12.3), Inches(5.5))
    table = shape.table
    
    table.columns[0].width = Inches(3.3)
    for i in range(1, 7): table.columns[i].width = Inches(1.5)
    
    def style_cell(cell, text, is_header=False, is_parent=False, is_total=False, align=PP_ALIGN.RIGHT):
        cell.text = str(text)
        cell.text_frame.paragraphs[0].font.size = Pt(10)
        cell.text_frame.paragraphs[0].font.name = 'Arial'
        cell.text_frame.paragraphs[0].alignment = align
        cell.vertical_anchor = 3
        
        cell.fill.solid()
        if is_header:
            cell.fill.fore_color.rgb = RGBColor(217, 225, 242)
            cell.text_frame.paragraphs[0].font.bold = True
        elif is_total:
            cell.fill.fore_color.rgb = RGBColor(217, 217, 217)
            cell.text_frame.paragraphs[0].font.bold = True
        elif is_parent:
            cell.fill.fore_color.rgb = RGBColor(255, 242, 204)
            cell.text_frame.paragraphs[0].font.bold = True
        else:
            cell.fill.fore_color.rgb = RGBColor(255, 255, 255)

    style_cell(table.cell(0, 0), "Jenis DBH", is_header=True, align=PP_ALIGN.CENTER)
    table.cell(0, 0).merge(table.cell(1, 0))
    
    style_cell(table.cell(0, 1), "Kurang Bayar (KB)", is_header=True, align=PP_ALIGN.CENTER)
    table.cell(0, 1).merge(table.cell(0, 3))
    
    style_cell(table.cell(0, 4), "Lebih Bayar (LB)", is_header=True, align=PP_ALIGN.CENTER)
    table.cell(0, 4).merge(table.cell(0, 6))
    
    style_cell(table.cell(1, 1), "Sisa KB s.d. TA 2023", is_header=True, align=PP_ALIGN.CENTER)
    style_cell(table.cell(1, 2), "KB TA 2024", is_header=True, align=PP_ALIGN.CENTER)
    style_cell(table.cell(1, 3), "KB s.d. TA 2024", is_header=True, align=PP_ALIGN.CENTER)
    style_cell(table.cell(1, 4), "Sisa LB s.d. TA 2023", is_header=True, align=PP_ALIGN.CENTER)
    style_cell(table.cell(1, 5), "LB TA 2024", is_header=True, align=PP_ALIGN.CENTER)
    style_cell(table.cell(1, 6), "LB s.d. TA 2024", is_header=True, align=PP_ALIGN.CENTER)
    
    for r_idx, row_data in df_hasil.iterrows():
        excel_row = r_idx + 2
        is_parent = row_data.get('is_parent', False)
        is_total = row_data.get('is_total', False)
        
        style_cell(table.cell(excel_row, 0), row_data['Jenis DBH'], is_parent=is_parent, is_total=is_total, align=PP_ALIGN.LEFT)
        for c_idx, col in enumerate(['KB 2023', 'KB 2024', 'KB Total', 'LB 2023', 'LB 2024', 'LB Total']):
            val = f"{row_data[col]:,.0f}" if row_data[col] != 0 else "-"
            style_cell(table.cell(excel_row, c_idx + 1), val, is_parent=is_parent, is_total=is_total)
            
    ppt_stream = io.BytesIO()
    prs.save(ppt_stream)
    ppt_stream.seek(0)
    return ppt_stream

# ==========================================
# 4. TABS
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
    st.subheader("Visualisasi Data DBH")
    
    col1, col2 = st.columns(2)
    
    with col1:
        st.markdown("**Total Alokasi Reguler DBH per Tahun**")
        df_alokasi = df[df['keterangan'].str.contains('Alokasi DBH', case=False, na=False)]
        if not df_alokasi.empty:
            agg_alokasi = df_alokasi.groupby('tahun')['nilai'].sum().reset_index()
            fig1 = px.bar(agg_alokasi, x='tahun', y='nilai', text_auto='.2s', template='plotly_dark')
            st.plotly_chart(fig1, use_container_width=True)
        else:
            st.info("Data Alokasi tidak ditemukan.")
            
    with col2:
        st.markdown("**Total Sisa Kurang Bayar (KB) per Tahun**")
        df_kb = df[df['keterangan'].str.contains('Sisa KB', case=False, na=False)]
        if not df_kb.empty:
            agg_kb = df_kb.groupby('tahun')['nilai'].sum().reset_index()
            fig2 = px.bar(agg_kb, x='tahun', y='nilai', text_auto='.2s', template='plotly_dark', color_discrete_sequence=['#ff7f0e'])
            st.plotly_chart(fig2, use_container_width=True)
        else:
            st.info("Data Sisa KB tidak ditemukan.")

# ------------------------------------------
# TAB 2: LAPORAN KB/LB
# ------------------------------------------
with tab2:
    sub_tab1, sub_tab2 = st.tabs(["Pencarian per Provinsi", "Pencarian per Pemda/Jenis"])
    
    # -- SUB TAB 1: Provinsi --
    with sub_tab1:
        st.subheader("Laporan Sisa KB & LB per Provinsi")
        semua_prov = sorted(df['namaprovinsi'].dropna().unique())
        idx_jateng = semua_prov.index("Jawa Tengah") if "Jawa Tengah" in semua_prov else 0
        pilih_prov_kblb = st.selectbox("Pilih Provinsi:", semua_prov, index=idx_jateng, key='prov_kblb')
        
        df_prov = df[(df['namaprovinsi'] == pilih_prov_kblb) & (df['level0'] == 'DBH')].copy()
        
        if df_prov.empty:
            st.warning("Data tidak ditemukan.")
        else:
            with st.expander("⚙️ Pengaturan Teks Transaksi KB/LB"):
                opsi_ket = sorted(df_prov['keterangan'].dropna().unique())
                def_kb23 = next((k for k in opsi_ket if 'Sisa KB s.d. TA 2023 KMK 198' in k), opsi_ket[0] if opsi_ket else "")
                def_kb24 = next((k for k in opsi_ket if 'Penetapan KB TA 2024 PMK 120' in k), opsi_ket[0] if opsi_ket else "")
                def_lb23 = next((k for k in opsi_ket if 'Sisa LB s.d. TA 2023 KMK 44' in k), opsi_ket[0] if opsi_ket else "")
                def_lb24 = next((k for k in opsi_ket if 'Penetapan LB TA 2024 PMK 120' in k or 'Sisa LB s.d. TA 2024 KMK 198' in k), opsi_ket[0] if opsi_ket else "")

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
            if prov_name not in hasil.index: hasil.loc[prov_name] = hasil.sum()

            hasil = hasil.reset_index().rename(columns={'index': 'pemda'})
            hasil['is_prov'] = hasil['pemda'].str.contains('Provinsi', case=False, na=False)
            hasil = hasil.sort_values(by=['is_prov', 'pemda'], ascending=[False, True]).drop(columns=['is_prov'])

            hasil['KB_Total'] = hasil['KB_2023'] + hasil['KB_2024']
            hasil['LB_Total'] = hasil['LB_2023'] + hasil['LB_2024']
            hasil.insert(0, 'No', range(1, len(hasil) + 1))

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

    # -- SUB TAB 2: Pemda/Jenis (SESUAI GAMBAR) --
    with sub_tab2:
        st.subheader("Laporan Sisa KB & LB per Pemda (KMK 198/2026)")
        semua_pemda = sorted(df['pemda'].dropna().unique())
        idx_pemda = next((i for i, p in enumerate(semua_pemda) if 'Indragiri Hulu' in str(p)), 0)
        pilih_pemda = st.selectbox("Pilih Pemda (Kab/Kota/Provinsi):", semua_pemda, index=idx_pemda)
        
        df_pemda = df[(df['pemda'] == pilih_pemda) & (df['level0'] == 'DBH')].copy()
        
        if df_pemda.empty:
            st.warning("Data tidak ditemukan.")
        else:
            # Tambahkan Pengaturan Dropdown agar Anda bisa memilih teks yang tepat
            with st.expander("⚙️ Pengaturan Teks Transaksi (Klik untuk menyesuaikan)"):
                opsi_ket = sorted(df_pemda['keterangan'].dropna().unique())
                
                # Tebakan otomatis berdasarkan list teks Anda
                def_kb23 = next((k for k in opsi_ket if 'Sisa KB s.d. TA 2023 KMK 198' in k), opsi_ket[0] if opsi_ket else "")
                def_kb24 = next((k for k in opsi_ket if 'Penetapan KB TA 2024 PMK 120' in k), opsi_ket[0] if opsi_ket else "")
                def_lb23 = next((k for k in opsi_ket if 'Sisa LB s.d. TA 2023 KMK 198' in k), opsi_ket[0] if opsi_ket else "")
                def_lb24 = next((k for k in opsi_ket if 'Penetapan LB TA 2024 PMK 120' in k), opsi_ket[0] if opsi_ket else "")

                c1, c2, c3, c4 = st.columns(4)
                KET_KB_2023 = c1.selectbox("Sisa KB s.d. TA 2023", opsi_ket, index=opsi_ket.index(def_kb23) if def_kb23 in opsi_ket else 0)
                KET_KB_2024 = c2.selectbox("KB TA 2024", opsi_ket, index=opsi_ket.index(def_kb24) if def_kb24 in opsi_ket else 0)
                KET_LB_2023 = c3.selectbox("Sisa LB s.d. TA 2023", opsi_ket, index=opsi_ket.index(def_lb23) if def_lb23 in opsi_ket else 0)
                KET_LB_2024 = c4.selectbox("LB TA 2024", opsi_ket, index=opsi_ket.index(def_lb24) if def_lb24 in opsi_ket else 0)

            def get_val(l1=None, l2=None, ket=None):
                d = df_pemda
                if l1: d = d[d['level1'] == l1]
                if l2: d = d[d['level2'] == l2]
                if ket: d = d[d['keterangan'] == ket]
                val = d['nilai'].sum()
                return 0 if pd.isna(val) else val

            # Hierarki kaku sesuai gambar
            hierarchy = [
                {'l1': 'Pajak', 'l2s': ['PPh', 'PBB', 'CHT']},
                {'l1': 'SDA', 'l2s': ['Migas', 'Minerba', 'Panas Bumi', 'Kehutanan', 'Perikanan']},
                {'l1': 'Lainnya', 'l2s': ['Sawit', 'Perkebunan Sawit']}
            ]

            rows = []
            for h in hierarchy:
                # Baris Parent (Pajak, SDA, Lainnya)
                kb23 = get_val(l1=h['l1'], ket=KET_KB_2023)
                kb24 = get_val(l1=h['l1'], ket=KET_KB_2024)
                lb23 = get_val(l1=h['l1'], ket=KET_LB_2023)
                lb24 = get_val(l1=h['l1'], ket=KET_LB_2024)
                
                rows.append({
                    'Jenis DBH': h['l1'],
                    'KB 2023': kb23, 'KB 2024': kb24, 'KB Total': kb23 + kb24,
                    'LB 2023': lb23, 'LB 2024': lb24, 'LB Total': lb23 + lb24,
                    'is_parent': True
                })

                # Baris Child (PPh, PBB, dll)
                for l2 in h['l2s']:
                    kb23_c = get_val(l1=h['l1'], l2=l2, ket=KET_KB_2023)
                    kb24_c = get_val(l1=h['l1'], l2=l2, ket=KET_KB_2024)
                    lb23_c = get_val(l1=h['l1'], l2=l2, ket=KET_LB_2023)
                    lb24_c = get_val(l1=h['l1'], l2=l2, ket=KET_LB_2024)

                    display_name = 'Perkebunan Sawit' if l2 in ['Sawit', 'Perkebunan Sawit'] else l2

                    if any(r['Jenis DBH'] == display_name for r in rows): continue

                    rows.append({
                        'Jenis DBH': display_name,
                        'KB 2023': kb23_c, 'KB 2024': kb24_c, 'KB Total': kb23_c + kb24_c,
                        'LB 2023': lb23_c, 'LB 2024': lb24_c, 'LB Total': lb23_c + lb24_c,
                        'is_parent': False
                    })

            # Baris Total Keseluruhan
            tot_kb23 = sum(r['KB 2023'] for r in rows if r['is_parent'])
            tot_kb24 = sum(r['KB 2024'] for r in rows if r['is_parent'])
            tot_lb23 = sum(r['LB 2023'] for r in rows if r['is_parent'])
            tot_lb24 = sum(r['LB 2024'] for r in rows if r['is_parent'])

            rows.append({
                'Jenis DBH': 'Total',
                'KB 2023': tot_kb23, 'KB 2024': tot_kb24, 'KB Total': tot_kb23 + tot_kb24,
                'LB 2023': tot_lb23, 'LB 2024': tot_lb24, 'LB Total': tot_lb23 + tot_lb24,
                'is_parent': False, 'is_total': True
            })

            df_hasil_pemda = pd.DataFrame(rows)

            # Tampilan HTML Tabel
            css_kblb = """
            <style>
            .tbl-kblb { width: 100%; border-collapse: collapse; font-family: 'Inter', sans-serif; font-size: 13px; background-color: white; color: black; }
            .tbl-kblb th, .tbl-kblb td { border: 1px solid #000; padding: 8px; }
            .tbl-kblb th { text-align: center; font-weight: bold; vertical-align: middle; }
            .tbl-kblb td { text-align: right; }
            .tbl-kblb td.text-left { text-align: left; }
            .bg-kb { background-color: #92d050; }
            .bg-lb { background-color: #9bc2e6; }
            .bg-title { background-color: #d9d9d9; }
            .bg-parent { background-color: #fff2cc; font-weight: bold; }
            .bg-total { background-color: #d9d9d9; font-weight: bold; }
            </style>
            """
            html_kblb = css_kblb + "<table class='tbl-kblb'><thead>"
            html_kblb += "<tr><th rowspan='2' class='bg-title' style='width: 25%;'>Jenis DBH</th><th colspan='3' class='bg-kb'>Kurang Bayar (KB)</th><th colspan='3' class='bg-lb'>Lebih Bayar (LB)</th></tr>"
            html_kblb += "<tr><th class='bg-kb'>Sisa KB s.d. TA 2023</th><th class='bg-kb'>KB TA 2024</th><th class='bg-kb'>KB s.d. TA 2024</th>"
            html_kblb += "<th class='bg-lb'>Sisa LB s.d. TA 2023</th><th class='bg-lb'>LB TA 2024</th><th class='bg-lb'>LB s.d. TA 2024</th></tr></thead><tbody>"
            
            for _, r in df_hasil_pemda.iterrows():
                is_tot = r.get('is_total', False)
                is_par = r.get('is_parent', False)
                tr_cls = "class='bg-total'" if is_tot else ("class='bg-parent'" if is_par else "")
                
                html_kblb += f"<tr {tr_cls}><td class='text-left'>{r['Jenis DBH']}</td>"
                for c in ['KB 2023', 'KB 2024', 'KB Total', 'LB 2023', 'LB 2024', 'LB Total']:
                    val = f"{r[c]:,.0f}" if r[c] != 0 else "-"
                    html_kblb += f"<td>{val}</td>"
                html_kblb += "</tr>"
            html_kblb += "</tbody></table>"
            
            components.html(html_kblb, height=600, scrolling=True)
            
            col_pdf, col_ppt = st.columns(2)
            with col_pdf:
                pdf_bytes = generate_pdf_kblb(df_hasil_pemda, pilih_pemda)
                st.download_button(
                    label="⬇️ Export Tabel ke PDF",
                    data=pdf_bytes,
                    file_name=f"KB_LB_{pilih_pemda}.pdf",
                    mime="application/pdf",
                    use_container_width=True
                )
            
            with col_ppt:
                if st.button("📊 Buat Slide PowerPoint (.pptx)", use_container_width=True):
                    with st.spinner("Membuat file PPT..."):
                        ppt_stream = generate_ppt_kblb_pemda(df_hasil_pemda, pilih_pemda)
                        st.success("✅ File PPT berhasil dibuat!")
                        st.download_button(
                            label="⬇️ Download File PPT",
                            data=ppt_stream,
                            file_name=f"KB_LB_{pilih_pemda}.pptx",
                            mime="application/vnd.openxmlformats-officedocument.presentationml.presentation",
                            use_container_width=True
                        )

# ------------------------------------------
# TAB 3: ALOKASI & REALISASI
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

        css = "<style>.custom-table { width: 100%; border-collapse: collapse; font-family: 'Inter', sans-serif; font-size: 11px; background-color: white; color: black; } .custom-table th, .custom-table td { border: 1px solid #000; padding: 4px 6px; } .custom-table th { background-color: #d9e1f2; text-align: center; font-weight: 700; vertical-align: middle; } .custom-table td { text-align: right; } .custom-table td.text-left { text-align: left; } .custom-table tr:nth-child(even) { background-color: #f2f2f2; } .total-row { background-color: #d9e1f2 !important; font-weight: 700; }</style>"
        
        html = css + "<table class='custom-table'><thead>"
        html += "<tr><th rowspan='3' style='width:200px;'>PEMERINTAH DAERAH</th>"
        html += "<th colspan='4'>2025</th>"
        html += "<th colspan='3'>2026</th>"
        html += "<th colspan='1'>2027</th>"
        html += "</tr><tr>"
        html += "<th rowspan='2'>Alokasi</th><th colspan='2'>Penyaluran</th><th rowspan='2'>Kurang Bayar</th>"
        html += "<th rowspan='2'>Alokasi</th><th colspan='2'>Penyaluran</th>"
        html += "<th rowspan='2'>Alokasi</th>"
        html += "</tr><tr>"
        html += "<th>Reguler</th><th>%</th>"
        html += "<th>Reguler</th><th>%</th>"
        html += "</tr></thead><tbody>"

        for _, row in df_hasil.iterrows():
            is_total = row['pemda'] == 'JUMLAH'
            tr_class = "class='total-row'" if is_total else ""
            style_prov = "style='font-weight:bold; background-color:#e6f0ff;'" if "Provinsi" in str(row['pemda']) else ""
            
            html += f"<tr {tr_class}><td class='text-left' {style_prov}>{row['pemda']}</td>"
            html += f"<td>{row['A_2025']:,.0f}</td><td>{row['P_2025']:,.0f}</td><td>{row['Pct_2025']:.2f}%</td><td>{row['KB_2025']:,.0f}</td>"
            html += f"<td>{row['A_2026']:,.0f}</td><td>{row['P_2026']:,.0f}</td><td>{row['Pct_2026']:.2f}%</td>"
            html += f"<td>{row['A_2027']:,.0f}</td>"
            html += "</tr>"
        html += "</tbody></table>"
        
        components.html(html, height=600, scrolling=True)
        
        st.markdown("### 📥 Export Laporan ke PowerPoint")
        if st.button("Buat Slide PowerPoint (.pptx)"):
            with st.spinner("Membuat file PPT..."):
                prs = Presentation()
                prs.slide_width = Inches(13.333)
                prs.slide_height = Inches(7.5)
                slide = prs.slides.add_slide(prs.slide_layouts[5])
                
                title = slide.shapes.title
                title.text = "ALOKASI DAN REALISASI DBH"
                title.text_frame.paragraphs[0].font.size = Pt(24)
                title.text_frame.paragraphs[0].font.bold = True
                title.text_frame.paragraphs[0].font.name = 'Arial'
                title.text_frame.paragraphs[0].alignment = PP_ALIGN.CENTER
                
                rows, cols = len(df_hasil) + 3, 9
                shape = slide.shapes.add_table(rows, cols, Inches(0.2), Inches(1.2), Inches(12.9), Inches(5.8))
                table = shape.table
                
                table.columns[0].width = Inches(2.9)
                for i in range(1, 9): table.columns[i].width = Inches(1.25)
                
                def style_cell_alokasi(cell, text, is_header=False, is_total=False, align=PP_ALIGN.RIGHT, bg_color=None):
                    cell.text = str(text)
                    cell.text_frame.paragraphs[0].font.size = Pt(10)
                    cell.text_frame.paragraphs[0].font.name = 'Arial'
                    cell.text_frame.paragraphs[0].alignment = align
                    cell.vertical_anchor = 3
                    
                    if is_header or is_total:
                        cell.fill.solid()
                        cell.fill.fore_color.rgb = RGBColor(217, 225, 242)
                        cell.text_frame.paragraphs[0].font.bold = True
                    elif bg_color:
                        cell.fill.solid()
                        cell.fill.fore_color.rgb = bg_color
                    else:
                        cell.fill.solid()
                        cell.fill.fore_color.rgb = RGBColor(255, 255, 255)

                style_cell_alokasi(table.cell(0, 0), "PEMERINTAH DAERAH", is_header=True, align=PP_ALIGN.CENTER)
                table.cell(0, 0).merge(table.cell(2, 0))
                
                style_cell_alokasi(table.cell(0, 1), "2025", is_header=True, align=PP_ALIGN.CENTER)
                table.cell(0, 1).merge(table.cell(0, 4))
                style_cell_alokasi(table.cell(0, 5), "2026", is_header=True, align=PP_ALIGN.CENTER)
                table.cell(0, 5).merge(table.cell(0, 7))
                style_cell_alokasi(table.cell(0, 8), "2027", is_header=True, align=PP_ALIGN.CENTER)
                
                style_cell_alokasi(table.cell(1, 1), "Alokasi", is_header=True, align=PP_ALIGN.CENTER)
                table.cell(1, 1).merge(table.cell(2, 1))
                
                style_cell_alokasi(table.cell(1, 2), "Penyaluran", is_header=True, align=PP_ALIGN.CENTER)
                table.cell(1, 2).merge(table.cell(1, 3))
                
                style_cell_alokasi(table.cell(1, 4), "Kurang Bayar", is_header=True, align=PP_ALIGN.CENTER)
                table.cell(1, 4).merge(table.cell(2, 4))
                
                style_cell_alokasi(table.cell(1, 5), "Alokasi", is_header=True, align=PP_ALIGN.CENTER)
                table.cell(1, 5).merge(table.cell(2, 5))
                
                style_cell_alokasi(table.cell(1, 6), "Penyaluran", is_header=True, align=PP_ALIGN.CENTER)
                table.cell(1, 6).merge(table.cell(1, 7))
                
                style_cell_alokasi(table.cell(1, 8), "Alokasi", is_header=True, align=PP_ALIGN.CENTER)
                table.cell(1, 8).merge(table.cell(2, 8))
                
                style_cell_alokasi(table.cell(2, 2), "Reguler", is_header=True, align=PP_ALIGN.CENTER)
                style_cell_alokasi(table.cell(2, 3), "%", is_header=True, align=PP_ALIGN.CENTER)
                style_cell_alokasi(table.cell(2, 6), "Reguler", is_header=True, align=PP_ALIGN.CENTER)
                style_cell_alokasi(table.cell(2, 7), "%", is_header=True, align=PP_ALIGN.CENTER)
                
                for r_idx, row_data in df_hasil.iterrows():
                    excel_row = r_idx + 3
                    is_total = row_data['pemda'] == 'JUMLAH'
                    is_prov = "Provinsi" in str(row_data['pemda'])
                    bg_color = RGBColor(230, 240, 255) if is_prov and not is_total else None
                    
                    style_cell_alokasi(table.cell(excel_row, 0), row_data['pemda'], is_total=is_total, align=PP_ALIGN.LEFT, bg_color=bg_color)
                    
                    style_cell_alokasi(table.cell(excel_row, 1), f"{row_data['A_2025']:,.0f}", is_total=is_total, bg_color=bg_color)
                    style_cell_alokasi(table.cell(excel_row, 2), f"{row_data['P_2025']:,.0f}", is_total=is_total, bg_color=bg_color)
                    style_cell_alokasi(table.cell(excel_row, 3), f"{row_data['Pct_2025']:.2f}%", is_total=is_total, bg_color=bg_color)
                    style_cell_alokasi(table.cell(excel_row, 4), f"{row_data['KB_2025']:,.0f}", is_total=is_total, bg_color=bg_color)
                    
                    style_cell_alokasi(table.cell(excel_row, 5), f"{row_data['A_2026']:,.0f}", is_total=is_total, bg_color=bg_color)
                    style_cell_alokasi(table.cell(excel_row, 6), f"{row_data['P_2026']:,.0f}", is_total=is_total, bg_color=bg_color)
                    style_cell_alokasi(table.cell(excel_row, 7), f"{row_data['Pct_2026']:.2f}%", is_total=is_total, bg_color=bg_color)
                    
                    style_cell_alokasi(table.cell(excel_row, 8), f"{row_data['A_2027']:,.0f}", is_total=is_total, bg_color=bg_color)
                
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
