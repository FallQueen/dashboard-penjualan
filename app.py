import streamlit as st
import pandas as pd
import plotly.express as px
from pptx import Presentation
from pptx.util import Inches
import io

# 1. Konfigurasi UI (Dark Theme & High Contrast)
st.set_page_config(
    page_title="Sales Analytics Pro - XL",
    page_icon="📊",
    layout="wide"
)

# Custom CSS untuk Ukuran Font Raksasa & Kontras Tinggi
st.markdown("""
    <style>
    /* Background utama */
    .stApp {
        background-color: #0E1117;
        color: #FFFFFF;
    }
    
    /* JUDUL UTAMA (h1) - EXTRA BESAR */
    h1 {
        font-size: 60px !important;
        font-weight: 900 !important;
        color: #60A5FA !important;
        margin-bottom: 15px !important;
        line-height: 1.1 !important;
    }
    
    /* SUB-JUDUL (h2) - BESAR & TEGAS */
    h2 {
        font-size: 30px !important;
        font-weight: 500 !important;
        color: #CBD5E1 !important;
        margin-top: 0px !important;
        margin-bottom: 50px !important;
        border-bottom: none !important;
    }

    /* PANEL KONTROL SIDEBAR */
    section[data-testid="stSidebar"] h2 {
        font-size: 35px !important;
        font-weight: 700 !important;
        color: #F8FAFC !important;
    }

    /* Teks di dalam Kartu Instruksi */
    .instruction-card {
        background-color: #1E293B;
        padding: 30px;
        border-radius: 15px;
        border: 3px solid #3B82F6; 
        box-shadow: 0 10px 20px rgba(0,0,0,0.5);
    }
    
    .instruction-card h3 {
        font-size: 30px !important;
        color: #F8FAFC !important;
    }
    
    .instruction-card p {
        font-size: 22px !important;
        color: #94A3B8 !important;
    }

    /* Ukuran Sidebar */
    [data-testid="stSidebar"] {
        background-color: #1E293B !important;
        min-width: 400px !important;
    }
    
    /* Memperbesar teks tombol */
    .stButton>button {
        font-size: 22px !important;
        font-weight: bold !important;
        height: 3.5em !important;
        background-color: #3B82F6 !important;
    }
    
    /* Memperbesar teks metrik */
    [data-testid="stMetricValue"] {
        font-size: 45px !important;
    }
    </style>
    """, unsafe_allow_html=True)

# --- SIDEBAR (PANEL KONTROL) ---
with st.sidebar:
    st.markdown("## 📥 Panel Kontrol")
    uploaded_file = st.file_uploader("Upload File Excel/CSV", type=['xlsx', 'csv'])
    st.markdown("---")
    st.markdown("### 🛠️ Status")
    if uploaded_file:
        st.success("✅ File Berhasil Dimasukkan")
    else:
        st.info("💡 Menunggu Upload...")

# --- LOGIKA DASHBOARD ---
if uploaded_file:
    try:
        # Proses Baca Data
        if uploaded_file.name.endswith('.csv'):
            df_mentah = pd.read_csv(uploaded_file, header=None)
        else:
            excel = pd.ExcelFile(uploaded_file)
            sheet_pilihan = st.sidebar.selectbox("Pilih Sheet:", excel.sheet_names) if len(excel.sheet_names) > 1 else excel.sheet_names[0]
            df_mentah = pd.read_excel(uploaded_file, sheet_name=sheet_pilihan, header=None)

        # Pembersihan Data (Logika Khusus)
        df_mentah = df_mentah.dropna(how='all', axis=0).dropna(how='all', axis=1).reset_index(drop=True)
        baris_minggu = df_mentah.iloc[0].ffill()
        baris_produk = df_mentah.iloc[1]
        data_isi = df_mentah.iloc[2:].copy()
        
        judul_kolom_baru = []
        for minggu, produk in zip(baris_minggu[1:], baris_produk[1:]):
            m_bersih = str(minggu).replace('nan', '').strip()
            p_bersih = str(produk).replace('nan', '').strip()
            judul_kolom_baru.append(f"{m_bersih} - {p_bersih}".strip(" -"))
            
        df_temp = pd.DataFrame(data_isi.iloc[:, 1:].values, columns=judul_kolom_baru)
        df_temp['Nama_Metrik'] = data_isi.iloc[:, 0].values
        
        filter_baris = df_temp.drop('Nama_Metrik', axis=1).apply(lambda r: pd.to_numeric(r, errors='coerce').notnull().any(), axis=1)
        df_temp = df_temp[filter_baris].reset_index(drop=True)
        df_panjang = df_temp.melt(id_vars=['Nama_Metrik'], var_name='Kategori', value_name='Nilai')
        df_hasil = df_panjang.pivot_table(index='Kategori', columns='Nama_Metrik', values='Nilai', aggfunc='first').reset_index()
        
        for kolom in df_hasil.columns:
            if kolom != 'Kategori':
                df_hasil[kolom] = pd.to_numeric(df_hasil[kolom], errors='coerce').fillna(0)

        # --- TAMPILAN DASHBOARD SETELAH UPLOAD ---
        st.markdown(f"<h1>Laporan: {uploaded_file.name}</h1>", unsafe_allow_html=True)
        
        m1, m2 = st.columns(2)
        m1.metric("Data Minggu", f"{len(df_hasil)} Kolom")
        m2.metric("Jumlah Item", f"{len(df_hasil.columns)-1} Baris")

        st.markdown("---")

        # Kontrol Grafik
        pilihan_data = st.selectbox("🎯 Pilih Data Penjualan:", [c for c in df_hasil.columns if c != 'Kategori'])
        
        fig = px.bar(df_hasil, x='Kategori', y=pilihan_data, text_auto='.2s', 
                     color_discrete_sequence=['#60A5FA'], title=f"Analisis {pilihan_data}")
        
        fig.update_layout(
            template="plotly_dark",
            font=dict(size=20),
            height=650
        )
        st.plotly_chart(fig, use_container_width=True)

        # Bagian PPT
        st.markdown("### 💾 Opsi Laporan")
        def buat_ppt(grafik, judul):
            prs = Presentation()
            slide = prs.slides.add_slide(prs.slide_layouts[5])
            slide.shapes.title.text = judul
            img_io = io.BytesIO(grafik.to_image(format="png", width=1200, height=700))
            slide.shapes.add_picture(img_io, Inches(0.5), Inches(1.5), width=Inches(9))
            out = io.BytesIO()
            prs.save(out)
            return out.getvalue()

        if st.button("🚀 Buat Slide PowerPoint"):
            data_ppt = buat_ppt(fig, f"Laporan {pilihan_data}")
            st.download_button("📥 Download PPT", data_ppt, f"Laporan_{pilihan_data}.pptx")

        with st.expander("🔍 Lihat Detail Tabel Data (Pengecekan)"):
            st.dataframe(df_hasil, use_container_width=True)

    except Exception as e:
        st.error(f"Gagal memproses data: {e}")

else:
    # --- TAMPILAN AWAL (LANDING PAGE) ---
    st.markdown("<h1>Portal Analisis Data Anda</h1>", unsafe_allow_html=True)
    st.markdown("<h2>Dashboard eksekutif untuk monitoring laporan mingguan secara real-time.</h2>", unsafe_allow_html=True)
    
    col_1, col_2, col_3 = st.columns(3)
    
    with col_1:
        st.markdown("""
        <div class='instruction-card'>
        <h1 style='font-size: 60px;'>📂</h1>
        <h3>1. Unggah</h3>
        <p>Pilih file Excel laporan melalui menu <b>Panel Kontrol</b> di sebelah kiri.</p>
        </div>
        """, unsafe_allow_html=True)
        
    with col_2:
        st.markdown("""
        <div class='instruction-card'>
        <h1 style='font-size: 60px;'>📊</h1>
        <h3>2. Pantau</h3>
        <p>Lihat grafik tren penjualan yang besar dan jelas secara otomatis.</p>
        </div>
        """, unsafe_allow_html=True)
        
    with col_3:
        st.markdown("""
        <div class='instruction-card'>
        <h1 style='font-size: 60px;'>🎞️</h1>
        <h3>3. Presentasi</h3>
        <p>Simpan grafik ke PowerPoint untuk bahan rapat presentasi.</p>
        </div>
        """, unsafe_allow_html=True)