import streamlit as st
import pandas as pd
import plotly.express as px
import matplotlib.pyplot as plt
from fpdf import FPDF
import io

# 1. KONFIGURASI HALAMAN
st.set_page_config(
    page_title="Sales Analytics Pro",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded" 
)

# 2. CUSTOM CSS (Layout Bertumpuk & Sidebar Kaku)
st.markdown("""
    <style>
    .stApp {
        background-color: #0E1117 !important;
        color: #FFFFFF !important;
    }

    /* Layout Judul Bertumpuk (Stack) */
    .header-stack {
        text-align: center;
        width: 100%;
        margin-bottom: 40px;
    }
    .header-stack h1 {
        font-size: 55px !important;
        font-weight: 850 !important;
        color: #60A5FA !important;
        margin-bottom: 5px !important;
    }
    .header-stack h2 {
        font-size: 24px !important;
        font-weight: 400 !important;
        color: #CBD5E1 !important;
        border-bottom: none !important;
        margin-top: 0 !important;
    }

    /* Sidebar Kaku & Lebar Pas */
    [data-testid="stSidebar"] {
        min-width: 350px !important;
        background-color: #1E293B !important;
    }
    [data-testid="stSidebarCollapseButton"] {
        display: none !important;
    }

    /* Container Tengah */
    .main .block-container {
        max-width: 95% !important;
        padding-top: 2rem !important;
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
    </style>
    """, unsafe_allow_html=True)

# --- SIDEBAR (PANEL KONTROL) ---
with st.sidebar:
    st.markdown("## 📥 Panel Kontrol")
    uploaded_file = st.file_uploader("Upload File Laporan Papa", type=['xlsx', 'csv'])
    st.markdown("---")
    
    selected_sheet = None
    if uploaded_file and uploaded_file.name.endswith(('.xlsx', '.xls')):
        excel_file = pd.ExcelFile(uploaded_file)
        if len(excel_file.sheet_names) > 1:
            st.markdown("### 📂 Pilih Halaman")
            selected_sheet = st.selectbox("Pilih Sheet Data:", excel_file.sheet_names)
        else:
            selected_sheet = excel_file.sheet_names[0]
    st.info("💡 Panel kontrol selalu aktif untuk memudahkan ganti data.")

# --- LOGIKA DASHBOARD ---
if uploaded_file:
    try:
        # 1. BACA & PROSES DATA
        if uploaded_file.name.endswith('.csv'):
            df_raw = pd.read_csv(uploaded_file, header=None)
        else:
            df_raw = pd.read_excel(uploaded_file, sheet_name=selected_sheet, header=None)

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

        # --- UI DASHBOARD ---
        st.markdown(f"<h1 style='text-align: center; color: #60A5FA; font-size: 50px;'>📊 Laporan: {uploaded_file.name}</h1>", unsafe_allow_html=True)
        st.markdown("---")

        # Grafik Interaktif (Web)
        pilihan = st.selectbox("🎯 Pilih Metrik Penjualan:", [c for c in df_final.columns if c != 'Kategori'])
        fig = px.bar(df_final, x='Kategori', y=pilihan, text_auto='.2s', color_discrete_sequence=['#60A5FA'])
        fig.update_layout(template="plotly_dark", height=550, title=f"Tren {pilihan}", title_x=0.5)
        st.plotly_chart(fig, use_container_width=True)

        st.markdown("---")

        # Tabel Konversi (Di Atas)
        st.markdown("### 📋 Tabel Hasil Konversi Data Papa")
        st.dataframe(df_final, use_container_width=True, height=400)

        # Menu Ekspor PDF (Anti Error)
        st.markdown("<br>", unsafe_allow_html=True)
        st.markdown("### 📄 Menu Ekspor Laporan (PDF)")
        st.write("Simpan laporan grafik dan data ini ke format PDF untuk dicetak atau dikirim.")
        
        if st.button("🚀 Buat File PDF"):
            try:
                # 1. Buat Gambar Grafik Pakai Matplotlib
                fig_pdf, ax = plt.subplots(figsize=(10, 5))
                ax.bar(df_final['Kategori'], df_final[pilihan], color='#60A5FA')
                ax.set_title(f"Analisis Penjualan: {pilihan}", fontsize=14, fontweight='bold')
                ax.set_xlabel('Kategori Minggu - Produk')
                ax.set_ylabel(pilihan)
                plt.xticks(rotation=45, ha='right')
                plt.tight_layout()

                # Simpan gambar ke memori sementara
                img_buf = io.BytesIO()
                fig_pdf.savefig(img_buf, format='png')
                img_buf.seek(0)
                plt.close(fig_pdf)

                # 2. Buat Dokumen PDF
                pdf = FPDF()
                pdf.add_page()
                
                # Tambah Judul PDF
                pdf.set_font("helvetica", "B", 18)
                pdf.cell(0, 10, f"Laporan Eksekutif: {pilihan}", new_x="LMARGIN", new_y="NEXT", align="C")
                pdf.ln(5)
                
                # Masukkan Gambar Grafik ke PDF
                pdf.image(img_buf, x=10, w=190)
                pdf.ln(5)
                
                # Output PDF ke tombol Download (PERBAIKAN DI SINI: convert ke bytes)
                pdf_bytes = bytes(pdf.output())
                
                st.download_button(
                    label="📥 Download PDF Sekarang",
                    data=pdf_bytes,
                    file_name=f"Laporan_{pilihan}.pdf",
                    mime="application/pdf"
                )
                st.success("✅ File PDF berhasil dibuat! Silakan klik tombol Download di atas.")
                
            except Exception as e:
                st.error(f"Gagal membuat PDF. Detail error: {e}")

    except Exception as e:
        st.error(f"Gagal memproses data: {e}")

else:
    # --- TAMPILAN AWAL (JUDUL BERTUMPUK) ---
    st.markdown("""
        <div class="header-stack">
            <h1>Portal Analisis Data Anda</h1>
            <h2>Dashboard eksekutif monitoring laporan mingguan</h2>
        </div>
    """, unsafe_allow_html=True)

    c1, c2, c3 = st.columns(3)
    steps = [
        ("📁", "1. Unggah", "Gunakan Panel Kontrol di kiri untuk memasukkan file laporan Papa."),
        ("📊", "2. Pantau", "Lihat tren data melalui grafik interaktif dan tabel konversi."),
        ("📄", "3. Ekspor", "Download hasil ke PDF yang rapi dan siap untuk dibagikan.")
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
