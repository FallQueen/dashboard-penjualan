import streamlit as st
import pandas as pd
import plotly.express as px
from pptx import Presentation
from pptx.util import Inches
import io

# 1. KONFIGURASI HALAMAN
st.set_page_config(
    page_title="Sales Analytics Pro",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded" 
)

# 2. CUSTOM CSS (Jurus Rata Tengah & Sejajar)
st.markdown("""
    <style>
    .stApp {
        background-color: #0E1117 !important;
        color: #FFFFFF !important;
    }

    /* Judul & Subtitle Sejajar (Horizontal) */
    .header-container {
        display: flex;
        align-items: center;
        justify-content: center;
        gap: 20px;
        width: 100%;
        margin-bottom: 30px;
        text-align: center;
    }
    .header-container h1 {
        font-size: 48px !important;
        font-weight: 850 !important;
        color: #60A5FA !important;
        margin: 0 !important;
    }
    .header-container h2 {
        font-size: 22px !important;
        font-weight: 400 !important;
        color: #CBD5E1 !important;
        margin: 0 !important;
        border-bottom: none !important;
    }

    /* Kunci Lebar Dashboard agar di Tengah */
    [data-testid="stAppViewBlockContainer"] {
        max-width: 1200px !important;
        margin: auto !important;
    }

    /* Card Instruksi */
    .instruction-card {
        background-color: #1E293B;
        padding: 30px;
        border-radius: 20px;
        border: 2px solid #3B82F6;
        text-align: center;
        min-height: 280px;
        display: flex;
        flex-direction: column;
        justify-content: center;
        align-items: center;
    }

    /* Sidebar Kaku (Hapus Tombol Tutup di Awal) */
    [data-testid="stSidebarCollapseButton"] {
        display: none !important;
    }
    </style>
    """, unsafe_allow_html=True)

# --- SIDEBAR ---
with st.sidebar:
    st.markdown("## 📥 Panel Kontrol")
    uploaded_file = st.file_uploader("Upload File Laporan Papa", type=['xlsx', 'csv'])
    st.markdown("---")
    st.info("💡 Pilih file Excel/CSV untuk memunculkan Dashboard Full Screen.")

# --- LOGIKA DASHBOARD ---
if uploaded_file:
    # JURUS FULL SCREEN: Sembunyikan Sidebar saat data tampil
    st.markdown("""
        <style>
        [data-testid="stSidebar"] { display: none !important; }
        .main .block-container { max-width: 95% !important; padding-top: 1rem !important; }
        </style>
        """, unsafe_allow_html=True)

    try:
        # 1. PROSES DATA
        if uploaded_file.name.endswith('.csv'):
            df_raw = pd.read_csv(uploaded_file, header=None)
        else:
            excel = pd.ExcelFile(uploaded_file)
            # Ambil sheet pertama saja agar tidak ribet
            df_raw = pd.read_excel(uploaded_file, sheet_name=0, header=None)

        # Pembersihan Data Format Papa
        df_raw = df_raw.dropna(how='all', axis=0).dropna(how='all', axis=1).reset_index(drop=True)
        weeks, prods = df_raw.iloc[0].ffill(), df_raw.iloc[1]
        headers = [f"{str(w).replace('nan','')} - {str(p).replace('nan','')}".strip(" -") for w, p in zip(weeks[1:], prods[1:])]
        df_temp = pd.DataFrame(df_raw.iloc[2:, 1:].values, columns=headers)
        df_temp['Metrik'] = df_raw.iloc[2:, 0].values
        mask = df_temp.drop('Metrik', axis=1).apply(lambda r: pd.to_numeric(r, errors='coerce').notnull().any(), axis=1)
        df_temp = df_temp[mask].reset_index(drop=True)
        df_final = df_temp.melt(id_vars=['Metrik'], var_name='Kategori', value_name='Nilai').pivot_table(index='Kategori', columns='Metrik', values='Nilai', aggfunc='first').reset_index()
        for c in df_final.columns:
            if c != 'Kategori': df_final[c] = pd.to_numeric(df_final[c], errors='coerce').fillna(0)

        # --- TAMPILAN DASHBOARD ---
        st.markdown(f"<h1>📊 Laporan: {uploaded_file.name}</h1>", unsafe_allow_html=True)
        
        # Tombol Kembali (Biar Sidebar Muncul Lagi)
        if st.button("🔄 Ganti File / Kembali ke Menu Utama"):
            st.rerun()

        st.markdown("---")

        # 2. GRAFIK
        pilihan = st.selectbox("🎯 Pilih Metrik Penjualan:", [c for c in df_final.columns if c != 'Kategori'])
        fig = px.bar(df_final, x='Kategori', y=pilihan, text_auto='.2s', color_discrete_sequence=['#60A5FA'])
        fig.update_layout(template="plotly_dark", height=600, title=f"Tren {pilihan}", title_x=0.5)
        st.plotly_chart(fig, width="stretch")

        # 3. TOMBOL EKSPOR & TABEL (Nggak bakal hilang)
        col_down, col_tab = st.columns([1, 1])
        
        with col_down:
            st.markdown("### 📽️ Presentasi")
            if st.button("🚀 Buat Slide PowerPoint"):
                try:
                    prs = Presentation()
                    slide = prs.slides.add_slide(prs.slide_layouts[5])
                    slide.shapes.title.text = f"Laporan {pilihan}"
                    img_io = io.BytesIO(fig.to_image(format="png", width=1200, height=700))
                    slide.shapes.add_picture(img_io, Inches(0.5), Inches(1.5), width=Inches(9))
                    out = io.BytesIO()
                    prs.save(out)
                    st.download_button("📥 Download .pptx", out.getvalue(), f"Laporan_{pilihan}.pptx")
                except:
                    st.error("Gagal membuat PPT. Coba ganti versi Python ke 3.11 di Settings.")

        with col_tab:
            st.markdown("### 📋 Tabel Hasil Konversi")
            st.dataframe(df_final, width="stretch", height=400)

    except Exception as e:
        st.error(f"Error: {e}")
        if st.button("Reset Aplikasi"): st.rerun()

else:
    # --- TAMPILAN AWAL (SEJAJAR) ---
    st.markdown("""
        <div class="header-container">
            <h1>Portal Analisis Data Anda</h1>
            <h2>Dashboard eksekutif monitoring laporan mingguan</h2>
        </div>
    """, unsafe_allow_html=True)
    
    

    c1, c2, c3 = st.columns(3)
    steps = [
        ("📁", "1. Unggah", "Gunakan menu di sebelah kiri untuk memasukkan file Excel Papa."),
        ("📊", "2. Pantau", "Dashboard akan otomatis Full Screen untuk grafik yang besar."),
        ("🎞️", "3. Ekspor", "Download hasil ke PowerPoint untuk bahan presentasi rapat.")
    ]
    for i, (icon, title, desc) in enumerate(steps):
        with [c1, c2, c3][i]:
            st.markdown(f"<div class='instruction-card'><div style='font-size: 70px;'>{icon}</div><h3>{title}</h3><p>{desc}</p></div>", unsafe_allow_html=True)
