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

# 2. CUSTOM CSS (Jurus Perbaikan & Centering)
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
        gap: 25px;
        width: 100%;
        margin-bottom: 30px;
        flex-wrap: wrap;
        text-align: center;
    }
    .header-container h1 {
        font-size: 50px !important;
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

    /* Sidebar Styling */
    [data-testid="stSidebar"] {
        background-color: #1E293B !important;
    }
    
    /* Hilangkan Tombol Panah Minimize di Awal */
    [data-testid="stSidebarCollapseButton"] {
        display: none !important;
    }

    /* Card Instruksi */
    .instruction-card {
        background-color: #1E293B;
        padding: 30px;
        border-radius: 20px;
        border: 2px solid #3B82F6;
        text-align: center;
        min-height: 300px;
        display: flex;
        flex-direction: column;
        justify-content: center;
        align-items: center;
    }

    /* Memperbesar Tab */
    button[data-baseweb="tab"] {
        font-size: 20px !important;
        font-weight: bold !important;
    }
    </style>
    """, unsafe_allow_html=True)

# --- SIDEBAR ---
with st.sidebar:
    st.markdown("## 📥 Panel Kontrol")
    uploaded_file = st.file_uploader("Upload File Laporan Papa", type=['xlsx', 'csv'])
    st.markdown("---")
    if uploaded_file:
        # TOMBOL UNTUK GANTI FILE (Reset)
        if st.button("🔄 Ganti File / Upload Ulang"):
            st.rerun()
    st.info("💡 Masukkan file laporan untuk melihat Dashboard.")

# --- LOGIKA DASHBOARD ---
if uploaded_file:
    try:
        # SEMBUNYIKAN SIDEBAR OTOMATIS SAAT ADA DATA (Pake CSS)
        st.markdown("""
            <style>
            [data-testid="stSidebar"] { display: none !important; }
            .main .block-container { max-width: 95% !important; }
            </style>
            """, unsafe_allow_html=True)

        # 1. PROSES BACA DATA
        if uploaded_file.name.endswith('.csv'):
            df_raw = pd.read_csv(uploaded_file, header=None)
        else:
            excel = pd.ExcelFile(uploaded_file)
            sheet = st.selectbox("Pilih Sheet Data:", excel.sheet_names) if len(excel.sheet_names) > 1 else excel.sheet_names[0]
            df_raw = pd.read_excel(uploaded_file, sheet_name=sheet, header=None)

        # 2. PEMBERSIHAN DATA (FORMAT KHUSUS PAPA)
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

        # --- TAMPILAN SETELAH UPLOAD ---
        st.markdown(f"<h1>📊 Laporan: {uploaded_file.name}</h1>", unsafe_allow_html=True)
        
        col_btn1, col_btn2 = st.columns([1, 5])
        with col_btn1:
            if st.button("⬅️ Ganti File"):
                st.rerun()

        st.markdown("---")

        # TABS: Grafik & Tabel Biar Rapi
        tab1, tab2 = st.tabs(["📈 Visualisasi Tren", "📋 Tabel Data Hasil Konversi"])

        with tab1:
            pilihan = st.selectbox("🎯 Pilih Metrik Penjualan:", [c for c in df_final.columns if c != 'Kategori'])
            fig = px.bar(df_final, x='Kategori', y=pilihan, text_auto='.2s', color_discrete_sequence=['#60A5FA'])
            fig.update_layout(template="plotly_dark", font=dict(size=16), height=650, title=f"Analisis {pilihan}", title_x=0.5)
            st.plotly_chart(fig, use_container_width=True)
            
            # PowerPoint Export
            if st.button("🚀 Ekspor ke PowerPoint"):
                prs = Presentation()
                slide = prs.slides.add_slide(prs.slide_layouts[5])
                slide.shapes.title.text = f"Analisis {pilihan}"
                img_io = io.BytesIO(fig.to_image(format="png", width=1200, height=700))
                slide.shapes.add_picture(img_io, Inches(0.5), Inches(1.5), width=Inches(9))
                out = io.BytesIO()
                prs.save(out)
                st.download_button("📥 Download .pptx", out.getvalue(), f"Laporan_{pilihan}.pptx")

        with tab2:
            st.markdown("### 📋 Hasil Konversi Data Papa")
            st.write("Data di bawah ini adalah hasil pembersihan otomatis dari format Excel Papa:")
            st.dataframe(df_final, use_container_width=True, height=500)

    except Exception as e:
        st.error(f"Gagal memproses data: {e}")
        if st.button("Coba Lagi"): st.rerun()

else:
    # --- TAMPILAN AWAL (SEJAJAR) ---
    st.markdown("""
        <div class="header-container">
            <h1>Portal Analisis Data Anda</h1>
            <h2>Dashboard eksekutif monitoring laporan mingguan</h2>
        </div>
    """, unsafe_allow_html=True)
    
    
    
    col1, col2, col3 = st.columns(3)
    steps = [
        ("📁", "1. Unggah", "Gunakan menu di sebelah kiri untuk memasukkan file Excel Papa."),
        ("📊", "2. Pantau", "Dashboard akan otomatis Full Screen untuk grafik yang besar."),
        ("🎞️", "3. Ekspor", "Download hasil ke PowerPoint untuk bahan presentasi rapat.")
    ]
    
    for i, (icon, title, desc) in enumerate(steps):
        with [col1, col2, col3][i]:
            st.markdown(f"""
                <div class="instruction-card">
                    <div style='font-size: 70px;'>{icon}</div>
                    <h3>{title}</h3>
                    <p>{desc}</p>
                </div>
            """, unsafe_allow_html=True)
