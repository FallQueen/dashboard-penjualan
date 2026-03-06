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
    initial_sidebar_state="collapsed"
)

# 2. CUSTOM CSS (Full Dark Mode & Perfect Centering)
st.markdown("""
    <style>
    .stApp {
        background-color: #0E1117 !important;
        color: #FFFFFF !important;
    }

    /* Memaksa kontainer utama ke tengah */
    .main .block-container {
        display: flex;
        flex-direction: column;
        align-items: center;
        justify-content: center;
    }

    /* Judul Utama Tengah */
    h1 {
        font-size: 65px !important;
        font-weight: 800 !important;
        color: #60A5FA !important;
        text-align: center !important;
        width: 100%;
        margin-bottom: 5px !important;
    }

    /* Sub-judul Tengah */
    h2 {
        font-size: 28px !important;
        font-weight: 400 !important;
        color: #CBD5E1 !important;
        text-align: center !important;
        width: 100%;
        margin-top: 0px !important;
        margin-bottom: 40px !important;
        border-bottom: none !important;
    }

    /* Sidebar Styling */
    [data-testid="stSidebar"] {
        background-color: #1E293B !important;
        min-width: 350px !important;
    }

    /* Kontainer Instruksi - Centering Flexbox */
    .instruction-container {
        display: flex;
        flex-direction: row;
        justify-content: center;
        align-items: stretch;
        gap: 20px;
        width: 100%;
        max-width: 1100px;
        margin: 0 auto;
    }

    .instruction-card {
        background-color: #1E293B;
        padding: 30px;
        border-radius: 20px;
        border: 2px solid #3B82F6;
        text-align: center;
        flex: 1;
        display: flex;
        flex-direction: column;
        justify-content: center;
        align-items: center;
        box-shadow: 0 10px 25px rgba(0,0,0,0.4);
    }

    .instruction-card h3 {
        font-size: 26px !important;
        margin: 15px 0 !important;
        color: #F8FAFC !important;
    }

    .instruction-card p {
        font-size: 19px !important;
        color: #94A3B8 !important;
    }

    /* Merapikan posisi Selectbox & Metrics agar ke tengah */
    [data-testid="stMetric"] {
        display: flex;
        justify-content: center;
        text-align: center;
    }

    div[data-baseweb="select"] {
        max-width: 600px;
        margin: 0 auto;
    }

    /* Responsive Mobile */
    @media (max-width: 768px) {
        h1 { font-size: 40px !important; }
        h2 { font-size: 22px !important; }
        .instruction-container { flex-direction: column; align-items: center; }
        .instruction-card { width: 90%; min-height: auto; }
    }
    </style>
    """, unsafe_allow_html=True)

# --- SIDEBAR ---
with st.sidebar:
    st.markdown("## 📥 Panel Kontrol")
    uploaded_file = st.file_uploader("Upload File Laporan Papa", type=['xlsx', 'csv'])
    st.markdown("---")
    st.info("💡 Pastikan format data sesuai (Minggu di baris 1, Produk di baris 2).")

# --- LOGIKA DASHBOARD ---
if uploaded_file:
    try:
        # Baca Data
        if uploaded_file.name.endswith('.csv'):
            df_raw = pd.read_csv(uploaded_file, header=None)
        else:
            excel = pd.ExcelFile(uploaded_file)
            sheet = st.sidebar.selectbox("Pilih Sheet:", excel.sheet_names) if len(excel.sheet_names) > 1 else excel.sheet_names[0]
            df_raw = pd.read_excel(uploaded_file, sheet_name=sheet, header=None)

        # Pembersihan
        df_raw = df_raw.dropna(how='all', axis=0).dropna(how='all', axis=1).reset_index(drop=True)
        weeks = df_raw.iloc[0].ffill()
        prods = df_raw.iloc[1]
        data_body = df_raw.iloc[2:].copy()
        
        headers = []
        for w, p in zip(weeks[1:], prods[1:]):
            headers.append(f"{str(w).replace('nan','')} - {str(p).replace('nan','')}".strip(" -"))
            
        df_temp = pd.DataFrame(data_body.iloc[:, 1:].values, columns=headers)
        df_temp['Metrik'] = data_body.iloc[:, 0].values
        mask = df_temp.drop('Metrik', axis=1).apply(lambda r: pd.to_numeric(r, errors='coerce').notnull().any(), axis=1)
        df_temp = df_temp[mask].reset_index(drop=True)
        
        df_melted = df_temp.melt(id_vars=['Metrik'], var_name='Kategori', value_name='Nilai')
        df_final = df_melted.pivot_table(index='Kategori', columns='Metrik', values='Nilai', aggfunc='first').reset_index()
        
        for c in df_final.columns:
            if c != 'Kategori': df_final[c] = pd.to_numeric(df_final[c], errors='coerce').fillna(0)

        # UI Dashboard (Tengah)
        st.markdown(f"<h1>Laporan: {uploaded_file.name}</h1>", unsafe_allow_html=True)
        
        c_m1, c_m2, c_m3 = st.columns([1, 2, 1]) # Column kosong di kiri-kanan untuk centering
        with c_m2:
            m1, m2 = st.columns(2)
            m1.metric("Periode Data", f"{len(df_final)} Kolom")
            m2.metric("Jenis Item", f"{len(df_final.columns)-1} Baris")

        st.markdown("<br>", unsafe_allow_html=True)
        
        # Pilihan Data (Dropdown Tengah)
        pilihan = st.selectbox("🎯 Pilih Metrik Penjualan:", [c for c in df_final.columns if c != 'Kategori'])
        
        # Grafik
        fig = px.bar(df_final, x='Kategori', y=pilihan, text_auto='.2s', 
                     color_discrete_sequence=['#60A5FA'], title=f"Visualisasi {pilihan}")
        fig.update_layout(template="plotly_dark", font=dict(size=16), height=600)
        st.plotly_chart(fig, use_container_width=True)

        # Bagian Bawah (Centering Buttons)
        st.markdown("---")
        b_col1, b_col2, b_col3 = st.columns([1, 1, 1])
        
        with b_col2:
            if st.button("🚀 Buat Slide PowerPoint"):
                prs = Presentation()
                slide = prs.slides.add_slide(prs.slide_layouts[5])
                slide.shapes.title.text = f"Analisis {pilihan}"
                img_io = io.BytesIO(fig.to_image(format="png", width=1200, height=700))
                slide.shapes.add_picture(img_io, Inches(0.5), Inches(1.5), width=inches(9))
                out = io.BytesIO()
                prs.save(out)
                st.download_button("📥 Download .pptx", out.getvalue(), f"Laporan_{pilihan}.pptx")
            
            with st.expander("🔍 Lihat Detail Tabel"):
                st.dataframe(df_final, use_container_width=True)

    except Exception as e:
        st.error(f"Gagal memproses data: {e}")

else:
    # --- TAMPILAN AWAL (LANDING PAGE) ---
    st.markdown("<h1>Portal Analisis Data anda</h1>", unsafe_allow_html=True)
    st.markdown("<h2>Dashboard eksekutif untuk monitoring laporan mingguan.</h2>", unsafe_allow_html=True)
    
    
    
    st.markdown("""
    <div class="instruction-container">
        <div class="instruction-card">
            <div style='font-size: 65px;'>📁</div>
            <h3>1. Unggah</h3>
            <p>Buka panel kontrol di kiri atas, lalu masukkan file laporan Papa.</p>
        </div>
        <div class="instruction-card">
            <div style='font-size: 65px;'>📊</div>
            <h3>2. Pantau</h3>
            <p>Lihat tren data melalui grafik interaktif yang bersih.</p>
        </div>
        <div class="instruction-card">
            <div style='font-size: 65px;'>🎞️</div>
            <h3>3. Ekspor</h3>
            <p>Download hasil ke PowerPoint untuk bahan presentasi rapat.</p>
        </div>
    </div>
    """, unsafe_allow_html=True)
