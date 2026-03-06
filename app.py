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
    layout="wide", # Tetap wide, tapi kita kunci lebarnya di CSS
    initial_sidebar_state="collapsed"
)

# 2. CUSTOM CSS (ULTRA-CENTERING)
st.markdown("""
    <style>
    /* Paksa Background Gelap Total */
    .stApp {
        background-color: #0E1117 !important;
        color: #FFFFFF !important;
    }

    /* KUNCI UTAMA: Memaksa Container Streamlit ke Tengah Viewport */
    [data-testid="stAppViewBlockContainer"] {
        max-width: 1100px !important; /* Kunci lebar maksimal agar tidak melar ke kanan */
        margin-left: auto !important;
        margin-right: auto !important;
        padding-top: 5rem !important;
        padding-left: 1rem !important;
        padding-right: 1rem !important;
    }

    /* Judul & Subtitle Rata Tengah */
    h1 {
        font-size: 70px !important;
        font-weight: 850 !important;
        color: #60A5FA !important;
        text-align: center !important;
        width: 100% !important;
    }
    h2 {
        font-size: 32px !important;
        color: #CBD5E1 !important;
        text-align: center !important;
        width: 100% !important;
        border-bottom: none !important;
        margin-bottom: 50px !important;
    }

    /* Kartu Instruksi Simetris */
    .instruction-container {
        display: flex;
        flex-direction: row;
        justify-content: center;
        gap: 25px;
        width: 100%;
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
        min-height: 350px;
        box-shadow: 0 15px 30px rgba(0,0,0,0.5);
    }
    .instruction-card h3 {
        font-size: 28px !important;
        margin-top: 20px !important;
        color: #F8FAFC !important;
    }
    .instruction-card p {
        font-size: 20px !important;
        color: #94A3B8 !important;
    }

    /* Sidebar Styling */
    section[data-testid="stSidebar"] {
        background-color: #1E293B !important;
        min-width: 380px !important;
    }

    /* Memaksa elemen UI Streamlit ke Tengah */
    .stSelectbox, .stButton, [data-testid="stMetric"] {
        display: flex;
        justify-content: center;
        width: 100%;
    }

    /* Responsive Mobile */
    @media (max-width: 768px) {
        h1 { font-size: 45px !important; }
        h2 { font-size: 24px !important; }
        .instruction-container { flex-direction: column; align-items: center; }
        .instruction-card { width: 100%; min-height: auto; }
    }
    </style>
    """, unsafe_allow_html=True)

# --- SIDEBAR ---
with st.sidebar:
    st.markdown("## 📥 Panel Kontrol")
    uploaded_file = st.file_uploader("Upload File Laporan Papa", type=['xlsx', 'csv'])
    st.markdown("---")
    st.info("Pilih file Excel/CSV di sidebar untuk memulai.")

# --- LOGIKA DASHBOARD ---
if uploaded_file:
    try:
        # BACA & BERSIHKAN DATA
        if uploaded_file.name.endswith('.csv'):
            df_raw = pd.read_csv(uploaded_file, header=None)
        else:
            excel = pd.ExcelFile(uploaded_file)
            sheet = st.sidebar.selectbox("Pilih Sheet:", excel.sheet_names) if len(excel.sheet_names) > 1 else excel.sheet_names[0]
            df_raw = pd.read_excel(uploaded_file, sheet_name=sheet, header=None)

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

        # UI DASHBOARD
        st.markdown(f"<h1>Laporan: {uploaded_file.name}</h1>", unsafe_allow_html=True)
        
        # Centering Metrics
        m_col1, m_col2 = st.columns(2)
        with m_col1: st.metric("Periode Terdeteksi", f"{len(df_final)} Kolom")
        with m_col2: st.metric("Jenis Data", f"{len(df_final.columns)-1} Baris")

        st.markdown("<br>", unsafe_allow_html=True)
        
        pilihan = st.selectbox("🎯 Pilih Metrik Penjualan:", [c for c in df_final.columns if c != 'Kategori'])
        
        # Grafik Tengah
        fig = px.bar(df_final, x='Kategori', y=pilihan, text_auto='.2s', 
                     color_discrete_sequence=['#60A5FA'], title=f"Visualisasi {pilihan}")
        fig.update_layout(template="plotly_dark", font=dict(size=18), height=600, title_x=0.5)
        st.plotly_chart(fig, use_container_width=True)

        # Download PPT
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
            
        with st.expander("🔍 Lihat Detail Tabel Data"):
            st.dataframe(df_final, use_container_width=True)

    except Exception as e:
        st.error(f"Error: {e}")

else:
    # --- TAMPILAN AWAL (ULTRA CENTERED) ---
    st.markdown("<h1>Portal Analisis Data Anda</h1>", unsafe_allow_html=True)
    st.markdown("<h2>Dashboard eksekutif untuk monitoring laporan mingguan secara real-time.</h2>", unsafe_allow_html=True)
    
    
    
    st.markdown("""
    <div class="instruction-container">
        <div class="instruction-card">
            <div style='font-size: 70px;'>📁</div>
            <h3>1. Unggah</h3>
            <p>Buka panel kontrol di kiri atas, masukkan file laporan Papa.</p>
        </div>
        <div class="instruction-card">
            <div style='font-size: 70px;'>📊</div>
            <h3>2. Pantau</h3>
            <p>Lihat tren data melalui grafik interaktif yang bersih.</p>
        </div>
        <div class="instruction-card">
            <div style='font-size: 70px;'>🎞️</div>
            <h3>3. Ekspor</h3>
            <p>Download hasil ke PowerPoint untuk bahan presentasi rapat.</p>
        </div>
    </div>
    """, unsafe_allow_html=True)
