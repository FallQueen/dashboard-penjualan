import streamlit as st
import pandas as pd
import plotly.express as px
from pptx import Presentation
from pptx.util import Inches
import io

# 1. Konfigurasi UI (Dark Mode Tetap & Sidebar Tertutup)
st.set_page_config(
    page_title="Sales Analytics Pro - Mobile Ready",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="collapsed" # Panel Kontrol otomatis tertutup (Full Screen)
)

# Custom CSS untuk Dark Mode Abadi & Responsif Mobile
st.markdown("""
    <style>
    /* Paksa Background Gelap */
    .stApp {
        background-color: #0E1117 !important;
        color: #FFFFFF !important;
    }
    
    /* JUDUL UTAMA (h1) - Responsif */
    h1 {
        font-size: 80px !important;
        font-weight: 900 !important;
        color: #60A5FA !important;
        margin-bottom: 15px !important;
    }
    
    /* SUB-JUDUL (h2) - Responsif */
    h2 {
        font-size: 40px !important;
        font-weight: 400 !important;
        color: #CBD5E1 !important;
        border-bottom: none !important;
    }

    /* Pengaturan untuk Layar HP (Mobile) */
    @media (max-width: 768px) {
        h1 { font-size: 45px !important; } /* Kecilkan judul di HP */
        h2 { font-size: 24px !important; } /* Kecilkan sub-judul di HP */
        [data-testid="stMetricValue"] { font-size: 30px !important; }
        .instruction-card h3 { font-size: 22px !important; }
    }

    /* Sidebar Styling */
    section[data-testid="stSidebar"] {
        background-color: #1E293B !important;
    }
    section[data-testid="stSidebar"] h2 {
        font-size: 30px !important;
    }

    .instruction-card {
        background-color: #1E293B;
        padding: 25px;
        border-radius: 15px;
        border: 2px solid #3B82F6;
        margin-bottom: 20px;
    }
    </style>
    """, unsafe_allow_html=True)

# --- SIDEBAR (Panel Kontrol) ---
with st.sidebar:
    st.markdown("## 📥 Panel Kontrol")
    uploaded_file = st.file_uploader("Upload File Excel/CSV", type=['xlsx', 'csv'])
    st.markdown("---")
    st.caption("Klik tanda panah (>) di pojok kiri atas untuk menutup kembali.")

# --- LOGIKA DASHBOARD ---
if uploaded_file:
    try:
        # Proses Baca Data (Logika Papa)
        if uploaded_file.name.endswith('.csv'):
            df_mentah = pd.read_csv(uploaded_file, header=None)
        else:
            excel = pd.ExcelFile(uploaded_file)
            sheet_pilihan = st.sidebar.selectbox("Pilih Sheet:", excel.sheet_names) if len(excel.sheet_names) > 1 else excel.sheet_names[0]
            df_mentah = pd.read_excel(uploaded_file, sheet_name=sheet_pilihan, header=None)

        df_mentah = df_mentah.dropna(how='all', axis=0).dropna(how='all', axis=1).reset_index(drop=True)
        baris_minggu = df_mentah.iloc[0].ffill()
        baris_produk = df_mentah.iloc[1]
        data_isi = df_mentah.iloc[2:].copy()
        
        judul_kolom_baru = []
        for minggu, produk in zip(baris_minggu[1:], baris_produk[1:]):
            judul_kolom_baru.append(f"{str(minggu).replace('nan','')} - {str(produk).replace('nan','')}".strip(" -"))
            
        df_temp = pd.DataFrame(data_isi.iloc[:, 1:].values, columns=judul_kolom_baru)
        df_temp['Nama_Metrik'] = data_isi.iloc[:, 0].values
        filter_baris = df_temp.drop('Nama_Metrik', axis=1).apply(lambda r: pd.to_numeric(r, errors='coerce').notnull().any(), axis=1)
        df_temp = df_temp[filter_baris].reset_index(drop=True)
        df_panjang = df_temp.melt(id_vars=['Nama_Metrik'], var_name='Kategori', value_name='Nilai')
        df_hasil = df_panjang.pivot_table(index='Kategori', columns='Nama_Metrik', values='Nilai', aggfunc='first').reset_index()
        
        for kolom in df_hasil.columns:
            if kolom != 'Kategori':
                df_hasil[kolom] = pd.to_numeric(df_hasil[kolom], errors='coerce').fillna(0)

        # --- TAMPILAN DASHBOARD ---
        st.markdown(f"<h1>Laporan: {uploaded_file.name}</h1>", unsafe_allow_html=True)
        
        # Grid Metric yang rapi di HP
        m1, m2 = st.columns(2)
        m1.metric("Data Minggu", f"{len(df_hasil)} Kolom")
        m2.metric("Jumlah Item", f"{len(df_hasil.columns)-1} Baris")

        st.markdown("---")

        pilihan_data = st.selectbox("🎯 Pilih Data Penjualan:", [c for c in df_hasil.columns if c != 'Kategori'])
        
        fig = px.bar(df_hasil, x='Kategori', y=pilihan_data, text_auto='.2s', 
                     color_discrete_sequence=['#60A5FA'], title=f"Analisis {pilihan_data}")
        
        fig.update_layout(template="plotly_dark", font=dict(size=16), height=600)
        st.plotly_chart(fig, use_container_width=True)

        # Download PPT
        if st.button("🚀 Buat Slide PowerPoint"):
            prs = Presentation()
            slide = prs.slides.add_slide(prs.slide_layouts[5])
            slide.shapes.title.text = f"Analisis {pilihan_data}"
            img_io = io.BytesIO(fig.to_image(format="png", width=1200, height=700))
            slide.shapes.add_picture(img_io, Inches(0.5), Inches(1.5), width=Inches(9))
            out = io.BytesIO()
            prs.save(out)
            st.download_button("📥 Download PPT", out.getvalue(), f"Laporan_{pilihan_data}.pptx")

        with st.expander("🔍 Lihat Detail Tabel Data"):
            st.dataframe(df_hasil, use_container_width=True)

    except Exception as e:
        st.error(f"Gagal memproses data: {e}")

else:
    # --- TAMPILAN AWAL (LANDING PAGE) ---
    st.markdown("<h1>Portal Analisis Data anda</h1>", unsafe_allow_html=True)
    st.markdown("<h2>Dashboard eksekutif untuk monitoring laporan mingguan.</h2>", unsafe_allow_html=True)
    
    # Grid Instruksi Responsif
    col1, col2, col3 = st.columns([1,1,1])
    
    steps = [
        ("📂", "1. Unggah", "Buka panel kontrol di kiri atas, lalu masukkan file Excel Papa."),
        ("📊", "2. Pantau", "Lihat tren penjualan otomatis dengan grafik yang besar."),
        ("🎞️", "3. PPT", "Simpan hasil ke PowerPoint untuk bahan rapat presentasi.")
    ]
    
    cols = [col1, col2, col3]
    for i, step in enumerate(steps):
        with cols[i]:
            st.markdown(f"""
            <div class='instruction-card'>
            <h1 style='font-size: 50px; margin:0;'>{step[0]}</h1>
            <h3>{step[1]}</h3>
            <p>{step[2]}</p>
            </div>
            """, unsafe_allow_html=True)
