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

# 2. CUSTOM CSS (Jurus Paksa & Layout Sejajar)
st.markdown("""
    <style>
    .stApp {
        background-color: #0E1117 !important;
        color: #FFFFFF !important;
    }

    /* 2a. Layout Judul & Subtitle Sejajar (Horizontal) */
    .header-container {
        display: flex;
        align-items: baseline;
        justify-content: center;
        gap: 30px;
        width: 100%;
        margin-bottom: 40px;
        flex-wrap: wrap;
    }
    .header-container h1 {
        font-size: 55px !important;
        font-weight: 850 !important;
        color: #60A5FA !important;
        margin: 0 !important;
    }
    .header-container h2 {
        font-size: 24px !important;
        font-weight: 400 !important;
        color: #CBD5E1 !important;
        margin: 0 !important;
        border-bottom: none !important;
    }

    /* 2b. Sidebar Styling */
    [data-testid="stSidebar"] {
        background-color: #1E293B !important;
        min-width: 400px !important;
    }
    [data-testid="stSidebar"] h2 {
        font-size: 35px !important;
        text-align: left !important;
    }

    /* 2c. Kartu Instruksi Simetris */
    .instruction-container {
        display: flex;
        justify-content: center;
        gap: 25px;
        width: 100%;
        max-width: 1100px;
    }
    .instruction-card {
        background-color: #1E293B;
        padding: 40px 20px;
        border-radius: 20px;
        border: 3px solid #3B82F6;
        text-align: center;
        flex: 1;
        display: flex;
        flex-direction: column;
        align-items: center;
        justify-content: center;
        min-height: 320px;
    }

    /* 2d. Memperbesar Elemen UI */
    .stSelectbox, .stButton {
        max-width: 600px !important;
        margin: auto !important;
    }
    [data-testid="stMetric"] {
        background: rgba(255, 255, 255, 0.05);
        padding: 20px;
        border-radius: 15px;
    }

    /* CSS Khusus Mobile */
    @media (max-width: 768px) {
        .header-container { flex-direction: column; align-items: center; text-align: center; }
        .header-container h1 { font-size: 35px !important; }
        .instruction-container { flex-direction: column; }
    }
    </style>
    """, unsafe_allow_html=True)

# --- SIDEBAR & LOGIKA FULL SCREEN ---
with st.sidebar:
    st.markdown("## 📥 Panel Kontrol")
    uploaded_file = st.file_uploader("Upload File Laporan Papa", type=['xlsx', 'csv'])
    st.markdown("---")

if not uploaded_file:
    # SEMBUNYIKAN TOMBOL MINIMIZE (Tanda Panah) SAAT AWAL
    st.markdown("<style>div[data-testid='stSidebarCollapseButton'] { display: none; }</style>", unsafe_allow_html=True)
else:
    # JURUS FULL SCREEN: Sembunyikan Sidebar sepenuhnya saat data muncul
    st.markdown("""
        <style>
        [data-testid="stSidebar"] { display: none !important; }
        [data-testid="stSidebarCollapseButton"] { display: none !important; }
        .main .block-container { max-width: 95% !important; padding-top: 2rem !important; }
        </style>
        """, unsafe_allow_html=True)

# --- LOGIKA DASHBOARD ---
if uploaded_file:
    try:
        # BACA DATA
        if uploaded_file.name.endswith('.csv'):
            df_raw = pd.read_csv(uploaded_file, header=None)
        else:
            excel = pd.ExcelFile(uploaded_file)
            sheet = st.sidebar.selectbox("Pilih Sheet:", excel.sheet_names) if len(excel.sheet_names) > 1 else excel.sheet_names[0]
            df_raw = pd.read_excel(uploaded_file, sheet_name=sheet, header=None)

        # BERSIHKAN DATA
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

        # TAMPILAN DASHBOARD (FULL SCREEN)
        st.markdown(f"<h1>Laporan: {uploaded_file.name}</h1>", unsafe_allow_html=True)
        
        m1, m2 = st.columns(2)
        m1.metric("Periode Terdeteksi", f"{len(df_final)} Kolom")
        m2.metric("Jenis Data", f"{len(df_final.columns)-1} Baris")

        st.markdown("<br>", unsafe_allow_html=True)
        pilihan = st.selectbox("🎯 Pilih Metrik Penjualan:", [c for c in df_final.columns if c != 'Kategori'])
        
        fig = px.bar(df_final, x='Kategori', y=pilihan, text_auto='.2s', color_discrete_sequence=['#60A5FA'], title=f"Visualisasi {pilihan}")
        fig.update_layout(template="plotly_dark", font=dict(size=18), height=650, title_x=0.5)
        st.plotly_chart(fig, use_container_width=True)

        # Bagian Bawah
        st.markdown("---")
        if st.button("🚀 Buat Slide PowerPoint"):
            prs = Presentation()
            slide = prs.slides.add_slide(prs.slide_layouts[5])
            slide.shapes.title.text = f"Analisis {pilihan}"
            img_io = io.BytesIO(fig.to_image(format="png", width=1200, height=700))
            slide.shapes.add_picture(img_io, Inches(0.5), Inches(1.5), width=Inches(9))
            out = io.BytesIO()
            prs.save(out)
            st.download_button("📥 Download PPT", out.getvalue(), f"Laporan_{pilihan}.pptx")
        
        # Tombol Reset (Agar bisa upload lagi karena Sidebar hilang)
        if st.button("🔄 Ganti File / Upload Ulang"):
            st.rerun()

    except Exception as e:
        st.error(f"Error: {e}")

else:
    # --- TAMPILAN AWAL (SEJAJAR) ---
    st.markdown("""
        <div class="header-container">
            <h1>Portal Analisis Data Anda</h1>
            <h2>Dashboard eksekutif untuk monitoring laporan mingguan</h2>
        </div>
    """, unsafe_allow_html=True)
    
    
    
    st.markdown("""
    <div style="display: flex; justify-content: center; width: 100%;">
        <div class="instruction-container">
            <div class="instruction-card">
                <div style='font-size: 70px;'>📁</div>
                <h3>1. Unggah</h3>
                <p>Gunakan menu di sebelah kiri untuk memasukkan file laporan Papa.</p>
            </div>
            <div class="instruction-card">
                <div style='font-size: 70px;'>📊</div>
                <h3>2. Pantau</h3>
                <p>Dashboard akan otomatis melebar untuk menampilkan grafik yang jelas.</p>
            </div>
            <div class="instruction-card">
                <div style='font-size: 70px;'>🎞️</div>
                <h3>3. Ekspor</h3>
                <p>Download hasil ke PowerPoint untuk bahan presentasi rapat.</p>
            </div>
        </div>
    </div>
    """, unsafe_allow_html=True)
