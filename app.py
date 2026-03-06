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

# 2. CUSTOM CSS (Full Dark Mode, Sejajar & Proportional Layout)
st.markdown("""
    <style>
    /* Background Utama */
    .stApp {
        background-color: #0E1117 !important;
        color: #FFFFFF !important;
    }

    /* Judul & Subtitle Sejajar (Horizontal) */
    .header-row {
        display: flex;
        align-items: center;
        justify-content: center;
        gap: 25px;
        width: 100%;
        margin-bottom: 30px;
        flex-wrap: wrap;
    }
    .header-row h1 {
        font-size: 42px !important;
        font-weight: 850 !important;
        color: #60A5FA !important;
        margin: 0 !important;
    }
    .header-row h2 {
        font-size: 18px !important;
        font-weight: 400 !important;
        color: #CBD5E1 !important;
        margin: 0 !important;
        border-bottom: none !important;
    }

    /* Pengaturan Lebar Sidebar */
    [data-testid="stSidebar"] {
        min-width: 350px !important;
        max-width: 350px !important;
        background-color: #1E293B !important;
    }

    /* Mengatur area konten utama agar proporsional */
    .main .block-container {
        max-width: 95% !important;
        padding-top: 2rem !important;
        margin: 0 auto !important;
    }

    /* Sembunyikan tombol minimize sidebar agar tetap "kaku" */
    [data-testid="stSidebarCollapseButton"] {
        display: none !important;
    }

    /* Kartu Instruksi */
    .instruction-card {
        background-color: #1E293B;
        padding: 25px;
        border-radius: 15px;
        border: 2px solid #3B82F6;
        text-align: center;
        min-height: 250px;
        display: flex;
        flex-direction: column;
        justify-content: center;
        align-items: center;
    }

    /* Merapikan Tabel & Grafik */
    .stDataFrame, .js-plotly-plot {
        border-radius: 10px;
        overflow: hidden;
    }
    </style>
    """, unsafe_allow_html=True)

# --- SIDEBAR (PANEL KONTROL SELALU MUNCUL) ---
with st.sidebar:
    st.markdown("## 📥 Panel Kontrol")
    uploaded_file = st.file_uploader("Upload File Laporan Papa", type=['xlsx', 'csv'])
    st.markdown("---")
    
    # Fitur Deteksi Sheet (Tetap di Sidebar agar rapi)
    selected_sheet = None
    if uploaded_file and uploaded_file.name.endswith(('.xlsx', '.xls')):
        excel_file = pd.ExcelFile(uploaded_file)
        if len(excel_file.sheet_names) > 1:
            st.markdown("### 📂 Pilih Halaman")
            selected_sheet = st.selectbox("Pilih Sheet Data:", excel_file.sheet_names)
        else:
            selected_sheet = excel_file.sheet_names[0]
    
    st.info("💡 Panel ini akan tetap muncul untuk memudahkan ganti file atau halaman.")

# --- LOGIKA DASHBOARD ---
if uploaded_file:
    try:
        # 1. BACA DATA
        if uploaded_file.name.endswith('.csv'):
            df_raw = pd.read_csv(uploaded_file, header=None)
        else:
            df_raw = pd.read_excel(uploaded_file, sheet_name=selected_sheet, header=None)

        # 2. PEMBERSIHAN DATA (FORMAT PAPA)
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
        st.markdown(f"<h1 style='text-align: center; color: #60A5FA;'>📊 Laporan: {uploaded_file.name}</h1>", unsafe_allow_html=True)
        st.markdown("---")

        # Layout Utama: Grafik (Atas) dan Tabel (Bawah)
        pilihan = st.selectbox("🎯 Pilih Metrik Penjualan untuk Grafik:", [c for c in df_final.columns if c != 'Kategori'])
        
        fig = px.bar(df_final, x='Kategori', y=pilihan, text_auto='.2s', color_discrete_sequence=['#60A5FA'])
        fig.update_layout(template="plotly_dark", height=550, title=f"Tren {pilihan}", title_x=0.5)
        st.plotly_chart(fig, use_container_width=True)

        st.markdown("---")

        # Bagian Tabel & Ekspor
        col_tab, col_down = st.columns([2, 1])
        
        with col_tab:
            st.markdown("### 📋 Tabel Data Konversi")
            st.dataframe(df_final, use_container_width=True, height=400)

        with col_down:
            st.markdown("### 📽️ Menu Ekspor")
            st.write("Klik tombol di bawah untuk mengunduh hasil analisis ke format PowerPoint.")
            if st.button("🚀 Buat Slide PowerPoint"):
                try:
                    prs = Presentation()
                    slide = prs.slides.add_slide(prs.slide_layouts[5])
                    slide.shapes.title.text = f"Analisis {pilihan}"
                    img_io = io.BytesIO(fig.to_image(format="png", width=1200, height=700))
                    slide.shapes.add_picture(img_io, Inches(0.5), Inches(1.5), width=Inches(9))
                    out = io.BytesIO()
                    prs.save(out)
                    st.download_button("📥 Download .pptx", out.getvalue(), f"Laporan_{pilihan}.pptx")
                except:
                    st.error("Gagal membuat PPT. Pastikan Kaleido terinstal.")

    except Exception as e:
        st.error(f"Gagal memproses data: {e}")

else:
    # --- TAMPILAN AWAL (JUDUL SEJAJAR) ---
    st.markdown("""
        <div class="header-row">
            <h1>Portal Analisis Data Anda</h1>
            <h2>Dashboard eksekutif monitoring laporan mingguan</h2>
        </div>
    """, unsafe_allow_html=True)
    
    

    c1, c2, c3 = st.columns(3)
    steps = [
        ("📁", "1. Unggah", "Gunakan Panel Kontrol di kiri untuk memasukkan file laporan Papa."),
        ("📊", "2. Pantau", "Lihat tren data melalui grafik interaktif dan tabel konversi."),
        ("🎞️", "3. Ekspor", "Download hasil ke PowerPoint untuk bahan presentasi rapat.")
    ]
    for i, (icon, title, desc) in enumerate(steps):
        with [c1, c2, c3][i]:
            st.markdown(f"""
                <div class="instruction-card">
                    <div style='font-size: 60px;'>{icon}</div>
                    <h3>{title}</h3>
                    <p>{desc}</p>
                </div>
            """, unsafe_allow_html=True)
