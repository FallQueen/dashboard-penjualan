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
    initial_sidebar_state="expanded" # Sekarang otomatis terbuka terus
)

# 2. CUSTOM CSS (Jurus Paksa Tengah & Dark Mode)
st.markdown("""
    <style>
    /* 1. Background Utama */
    .stApp {
        background-color: #0E1117 !important;
        color: #FFFFFF !important;
    }

    /* 2. Memaksa Kontainer Streamlit agar Benar-benar di Tengah */
    .main .block-container {
        max-width: 1200px !important;
        padding-left: 2rem !important;
        padding-right: 2rem !important;
        margin: auto !important;
        display: flex;
        flex-direction: column;
        align-items: center;
    }

    /* 3. Judul & Subtitle (Pasti Tengah) */
    h1 {
        font-size: 70px !important;
        font-weight: 850 !important;
        color: #60A5FA !important;
        text-align: center !important;
        width: 100% !important;
        margin-top: 20px !important;
    }
    h2 {
        font-size: 32px !important;
        font-weight: 400 !important;
        color: #CBD5E1 !important;
        text-align: center !important;
        width: 100% !important;
        border-bottom: none !important;
        margin-bottom: 50px !important;
    }

    /* 4. Kartu Instruksi (Flexbox Centering) */
    .centered-wrapper {
        display: flex;
        justify-content: center;
        width: 100%;
    }
    .instruction-container {
        display: flex;
        flex-direction: row;
        gap: 25px;
        width: 100%;
        max-width: 1000px;
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

    /* 5. Sidebar Styling */
    [data-testid="stSidebar"] {
        background-color: #1E293B !important;
        min-width: 380px !important;
    }
    [data-testid="stSidebar"] h2 {
        font-size: 35px !important;
        text-align: left !important;
    }

    /* 6. Memperbesar Elemen UI */
    .stSelectbox, .stButton {
        max-width: 600px !important;
        margin: auto !important;
    }
    
    [data-testid="stMetric"] {
        text-align: center !important;
        background: rgba(255, 255, 255, 0.05);
        padding: 20px;
        border-radius: 15px;
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
    st.info("Pilih file Excel/CSV di atas untuk menampilkan dashboard.")

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

        # UI DASHBOARD (SETELAH UPLOAD)
        st.markdown(f"<h1>Laporan: {uploaded_file.name}</h1>", unsafe_allow_html=True)
        
        # Centering Metrics
        c_m1, c_m2, c_m3 = st.columns([1, 2, 1])
        with c_m2:
            m1, m2 = st.columns(2)
            m1.metric("Periode Terdeteksi", f"{len(df_final)} Kolom")
            m2.metric("Jenis Data", f"{len(df_final.columns)-1} Baris")

        st.markdown("<br>", unsafe_allow_html=True)
        
        # Dropdown Tengah
        pilihan = st.selectbox("🎯 Pilih Metrik Penjualan:", [c for c in df_final.columns if c != 'Kategori'])
        
        # Grafik Tengah
        fig = px.bar(df_final, x='Kategori', y=pilihan, text_auto='.2s', 
                     color_discrete_sequence=['#60A5FA'], title=f"Visualisasi {pilihan}")
        fig.update_layout(template="plotly_dark", font=dict(size=18), height=650, title_x=0.5)
        st.plotly_chart(fig, use_container_width=True)

        # Tombol Tengah
        st.markdown("---")
        c_b1, c_b2, c_b3 = st.columns([1, 1, 1])
        with c_b2:
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
    # --- TAMPILAN AWAL (LANDING PAGE) ---
    st.markdown("<h1>Portal Analisis Data Anda</h1>", unsafe_allow_html=True)
    st.markdown("<h2>Dashboard eksekutif untuk monitoring laporan mingguan.</h2>", unsafe_allow_html=True)
    
    st.markdown("""
    <div style="display: flex; justify-content: center; width: 100%;">
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
    </div>
    """, unsafe_allow_html=True)
