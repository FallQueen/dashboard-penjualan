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

# 2. CUSTOM CSS (Full Dark Mode, Font Besar & Rata)
st.markdown("""
    <style>
    .stApp {
        background-color: #0E1117 !important;
        color: #FFFFFF !important;
    }

    /* Judul Utama */
    h1 {
        font-size: 65px !important;
        font-weight: 800 !important;
        color: #60A5FA !important;
        text-align: center !important;
        margin-bottom: 5px !important;
        line-height: 1.2 !important;
    }

    /* Sub-judul */
    h2 {
        font-size: 28px !important;
        font-weight: 400 !important;
        color: #CBD5E1 !important;
        text-align: center !important;
        margin-top: 0px !important;
        margin-bottom: 40px !important;
        border-bottom: none !important;
    }

    /* Sidebar */
    [data-testid="stSidebar"] {
        background-color: #1E293B !important;
        min-width: 350px !important;
    }
    [data-testid="stSidebar"] h2 {
        font-size: 30px !important;
        font-weight: bold !important;
    }

    /* Kartu Instruksi Simetris */
    .instruction-container {
        display: flex;
        flex-direction: row;
        justify-content: center;
        gap: 20px;
        width: 100%;
    }
    .instruction-card {
        background-color: #1E293B;
        padding: 25px;
        border-radius: 15px;
        border: 2px solid #3B82F6;
        text-align: center;
        flex: 1;
        min-height: 300px;
        display: flex;
        flex-direction: column;
        justify-content: center;
        align-items: center;
        box-shadow: 0 4px 15px rgba(0,0,0,0.3);
    }
    .instruction-card h3 {
        font-size: 24px !important;
        margin: 15px 0 10px 0 !important;
        color: #F8FAFC !important;
    }
    .instruction-card p {
        font-size: 18px !important;
        color: #94A3B8 !important;
        line-height: 1.4;
    }

    /* Responsive Mobile */
    @media (max-width: 768px) {
        h1 { font-size: 38px !important; }
        h2 { font-size: 20px !important; }
        .instruction-container { flex-direction: column; }
    }

    /* Tombol & Metric */
    .stButton>button {
        font-size: 18px !important;
        font-weight: bold !important;
        height: 3em !important;
        background-color: #3B82F6 !important;
    }
    [data-testid="stMetricValue"] {
        font-size: 42px !important;
    }
    </style>
    """, unsafe_allow_html=True)

# --- SIDEBAR ---
with st.sidebar:
    st.markdown("## 📥 Panel Kontrol")
    uploaded_file = st.file_uploader("Upload File Excel/CSV Papa", type=['xlsx', 'csv'])
    st.markdown("---")
    if uploaded_file:
        st.success("✅ File Terdeteksi")
    else:
        st.info("💡 Silakan masukkan file laporan.")

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

        # Proses Tabel Khusus Papa
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

        # Dashboard UI
        st.markdown(f"<h1>Laporan: {uploaded_file.name}</h1>", unsafe_allow_html=True)
        
        col_m1, col_m2 = st.columns(2)
        col_m1.metric("Periode Data", f"{len(df_final)} Kolom")
        col_m2.metric("Jenis Item", f"{len(df_final.columns)-1} Baris")

        st.markdown("---")
        pilihan = st.selectbox("🎯 Pilih Metrik Penjualan:", [c for c in df_final.columns if c != 'Kategori'])
        
        fig = px.bar(df_final, x='Kategori', y=pilihan, text_auto='.2s', 
                     color_discrete_sequence=['#60A5FA'], title=f"Analisis {pilihan}")
        fig.update_layout(template="plotly_dark", font=dict(size=16), height=600)
        st.plotly_chart(fig, use_container_width=True)

        # PPT Export
        st.markdown("### 💾 Export")
        col_p, col_t = st.columns(2)
        with col_p:
            if st.button("🚀 Buat Slide PowerPoint"):
                prs = Presentation()
                slide = prs.slides.add_slide(prs.slide_layouts[5])
                slide.shapes.title.text = f"Analisis {pilihan}"
                img_io = io.BytesIO(fig.to_image(format="png", width=1200, height=700))
                slide.shapes.add_picture(img_io, Inches(0.5), Inches(1.5), width=Inches(9))
                out = io.BytesIO()
                prs.save(out)
                st.download_button("📥 Download .pptx", out.getvalue(), f"Laporan_{pilihan}.pptx")
        with col_t:
            with st.expander("🔍 Lihat Tabel Detail"):
                st.dataframe(df_final, use_container_width=True)

    except Exception as e:
        st.error(f"Error: {e}")

else:
    # --- LANDING PAGE (AWAL) ---
    st.markdown("<h1>Portal Analisis Data Anda</h1>", unsafe_allow_html=True)
    st.markdown("<h2>Dashboard eksekutif monitoring laporan mingguan real-time.</h2>", unsafe_allow_html=True)
    
    st.markdown("""
    <div class="instruction-container">
        <div class="instruction-card">
            <div style='font-size: 60px;'>📁</div>
            <h3>1. Unggah</h3>
            <p>Buka panel kontrol di kiri atas, masukkan file laporan Papa.</p>
        </div>
        <div class="instruction-card">
            <div style='font-size: 60px;'>📊</div>
            <h3>2. Pantau</h3>
            <p>Lihat tren penjualan otomatis melalui grafik interaktif.</p>
        </div>
        <div class="instruction-card">
            <div style='font-size: 60px;'>📽️</div>
            <h3>3. Presentasi</h3>
            <p>Download hasil ke PowerPoint untuk langsung dipresentasikan.</p>
        </div>
    </div>
    """, unsafe_allow_html=True)
