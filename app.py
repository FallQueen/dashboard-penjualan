import streamlit as st
import pandas as pd
import plotly.express as px
import matplotlib.pyplot as plt
from fpdf import FPDF
import io
import re

# 1. KONFIGURASI HALAMAN
st.set_page_config(
    page_title="Sales Analytics Pro",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded" 
)

# 2. CUSTOM CSS
st.markdown("""
    <style>
    .stApp { background-color: #0E1117 !important; color: #FFFFFF !important; }
    .header-stack { text-align: center; width: 100%; margin-bottom: 40px; }
    .header-stack h1 { font-size: 55px !important; font-weight: 850 !important; color: #60A5FA !important; margin-bottom: 5px !important; }
    .header-stack h2 { font-size: 24px !important; font-weight: 400 !important; color: #CBD5E1 !important; margin-top: 0 !important; }
    [data-testid="stSidebar"] { min-width: 350px !important; background-color: #1E293B !important; }
    [data-testid="stSidebarCollapseButton"] { display: none !important; }
    .main .block-container { max-width: 95% !important; padding-top: 2rem !important; }
    .instruction-card { background-color: #1E293B; padding: 30px; border-radius: 20px; border: 2px solid #3B82F6; text-align: center; min-height: 280px; display: flex; flex-direction: column; justify-content: center; align-items: center; }
    [data-testid="stMetric"] { background: rgba(255, 255, 255, 0.05); padding: 15px; border-radius: 10px; text-align: center; }
    </style>
    """, unsafe_allow_html=True)

# --- SIDEBAR (PANEL KONTROL) ---
with st.sidebar:
    st.markdown("## 📥 Panel Kontrol")
    uploaded_file = st.file_uploader("Upload File Laporan Papa", type=['xlsx', 'csv', 'xls'])
    st.markdown("---")
    
    selected_sheet = None
    if uploaded_file and uploaded_file.name.endswith(('.xlsx', '.xls')):
        excel_file = pd.ExcelFile(uploaded_file)
        if len(excel_file.sheet_names) > 1:
            st.markdown("### 📂 Pilih Halaman")
            selected_sheet = st.selectbox("Pilih Sheet Data (Misal: Jember/Banyuwangi):", excel_file.sheet_names)
        else:
            selected_sheet = excel_file.sheet_names[0]
    st.info("💡 Panel kontrol selalu aktif untuk memudahkan ganti data wilayah.")

# --- LOGIKA DASHBOARD ---
if uploaded_file:
    try:
        # 1. BACA & PROSES DATA
        if uploaded_file.name.endswith('.csv'):
            df_raw = pd.read_csv(uploaded_file, header=None)
        else:
            df_raw = pd.read_excel(uploaded_file, sheet_name=selected_sheet, header=None)

        # Buang baris & kolom yang benar-benar tidak ada isinya sama sekali
        df_raw = df_raw.dropna(how='all', axis=0).dropna(how='all', axis=1).reset_index(drop=True)
        
        # Perbaiki typo "Weeek" jadi "Week"
        df_raw.iloc[0] = df_raw.iloc[0].astype(str).str.replace('Weeek', 'Week', regex=False)
        
        # Ekstrak Baris Minggu dan Produk (Mengatasi Merge Cells dengan ffill)
        weeks = df_raw.iloc[0].replace('nan', pd.NA).ffill()
        prods = df_raw.iloc[1]
        
        w_list = [str(w).replace('nan','').strip() for w in weeks[1:]]
        p_list = [str(p).replace('nan','').strip() for p in prods[1:]]
        
        kategori = [f"{w} ||| {p}" for w, p in zip(w_list, p_list)]
        
        df_temp = pd.DataFrame(df_raw.iloc[2:, 1:].values, columns=kategori)
        df_temp['Metrik'] = df_raw.iloc[2:, 0].values
        
        # Saring Metrik: Pertahankan semua baris yang ada namanya (seperti BPJ, OOS, dll)
        df_temp = df_temp[df_temp['Metrik'].notna() & (df_temp['Metrik'].astype(str).str.strip() != '')].reset_index(drop=True)
        
        df_melted = df_temp.melt(id_vars=['Metrik'], var_name='Kategori', value_name='Nilai')
        df_melted[['Minggu', 'Produk']] = df_melted['Kategori'].str.split(' \|\|\| ', expand=True)
        
        # Pembersih teks Minggu (Menyeragamkan W 39, Week 2, dll jadi Week X)
        def clean_week(s):
            match = re.search(r'(?:Week|Weeek|W)\s*(\d+)', str(s), flags=re.IGNORECASE)
            return f"Week {match.group(1)}" if match else str(s)
        
        df_melted['Minggu'] = df_melted['Minggu'].apply(clean_week)
        
        # Pembersih Angka & Anti-Error Excel (#DIV/0!)
        def clean_numeric(val):
            if isinstance(val, str):
                val = val.replace('%', '').replace(',', '').strip()
                if val in ['#DIV/0!', '-', '']: 
                    return 0
            return pd.to_numeric(val, errors='coerce')
            
        df_melted['Nilai'] = df_melted['Nilai'].apply(clean_numeric)
        
        # Pivot Data kembali ke bentuk rapi
        df_final = df_melted.pivot_table(index=['Minggu', 'Produk'], columns='Metrik', values='Nilai', aggfunc='first').reset_index()
        df_final = df_final.fillna(0)
        
        # Urutkan Minggu secara otomatis berdasarkan angkanya
        def get_week_num(w):
            match = re.search(r'\d+', str(w))
            return int(match.group(0)) if match else 0
        df_final['WeekNum'] = df_final['Minggu'].apply(get_week_num)
        df_final = df_final.sort_values(['WeekNum', 'Produk']).drop(columns=['WeekNum']).reset_index(drop=True)

        # --- UI DASHBOARD ---
        st.markdown(f"<h1 style='text-align: center; color: #60A5FA; font-size: 50px;'>📊 Laporan: {selected_sheet if selected_sheet else uploaded_file.name}</h1>", unsafe_allow_html=True)
        st.markdown("---")

        col_opt1, col_opt2 = st.columns(2)
        with col_opt1:
            # Pilihan Metrik yang otomatis menyesuaikan apa saja yang diketik Papa di Excel
            pilihan = st.selectbox("🎯 Pilih Metrik Penjualan:", [c for c in df_final.columns if c not in ['Minggu', 'Produk']])
        with col_opt2:
            jenis_grafik = st.selectbox("📈 Pilih Bentuk Diagram:", ["Diagram Batang Berdampingan (Grouped Bar)", "Diagram Garis Tren (Line)", "Diagram Lingkaran (Pie)"])

        # MENGHITUNG AVERAGE, TOTAL, DAN MAX
        rata_rata = df_final[pilihan].mean()
        total_semua = df_final[pilihan].sum()
        nilai_tertinggi = df_final[pilihan].max()

        st.markdown("### 💡 Ringkasan Eksekutif")
        c_sum1, c_sum2, c_sum3 = st.columns(3)
        satuan = "%" if "%" in pilihan else ""
        c_sum1.metric(label="Rata-rata (Average)", value=f"{rata_rata:,.2f}{satuan}")
        c_sum2.metric(label="Total Keseluruhan", value=f"{total_semua:,.0f}{satuan}")
        c_sum3.metric(label="Nilai Tertinggi", value=f"{nilai_tertinggi:,.0f}{satuan}")
        
        st.markdown("<br>", unsafe_allow_html=True)

        # GRAFIK INTERAKTIF (WEB)
        if jenis_grafik == "Diagram Batang Berdampingan (Grouped Bar)":
            fig = px.bar(df_final, x='Minggu', y=pilihan, color='Produk', barmode='group', text_auto='.2s', color_discrete_sequence=px.colors.qualitative.Pastel)
        elif jenis_grafik == "Diagram Garis Tren (Line)":
            fig = px.line(df_final, x='Minggu', y=pilihan, color='Produk', markers=True, color_discrete_sequence=px.colors.qualitative.Pastel)
        else:
            df_pie = df_final.groupby('Produk')[pilihan].sum().reset_index()
            fig = px.pie(df_pie, names='Produk', values=pilihan, hole=0.3, color_discrete_sequence=px.colors.qualitative.Pastel)
            
        fig.update_layout(template="plotly_dark", height=550, title=f"Visualisasi {pilihan} ({selected_sheet})", title_x=0.5)
        st.plotly_chart(fig, use_container_width=True)

        st.markdown("---")
        st.markdown("### 📋 Tabel Hasil Konversi Data Papa")
        st.dataframe(df_final, use_container_width=True, height=400, hide_index=True)

        # MENU EKSPOR PDF
        st.markdown("<br>", unsafe_allow_html=True)
        st.markdown("### 📄 Menu Ekspor Laporan (PDF)")
        if st.button("🚀 Buat File PDF"):
            try:
                fig_pdf, ax = plt.subplots(figsize=(10, 5))
                
                if jenis_grafik != "Diagram Lingkaran (Pie)":
                    df_pivot = df_final.pivot(index='Minggu', columns='Produk', values=pilihan)
                
                if jenis_grafik == "Diagram Batang Berdampingan (Grouped Bar)":
                    df_pivot.plot(kind='bar', ax=ax, colormap='Paired')
                    ax.set_ylabel(pilihan)
                    plt.xticks(rotation=0)
                elif jenis_grafik == "Diagram Garis Tren (Line)":
                    df_pivot.plot(kind='line', marker='o', ax=ax, linewidth=2, colormap='Paired')
                    ax.set_ylabel(pilihan)
                    ax.grid(True, linestyle='--', alpha=0.6)
                else: 
                    df_pie = df_final.groupby('Produk')[pilihan].sum()
                    ax.pie(df_pie, labels=df_pie.index, autopct='%1.1f%%', startangle=90, colors=plt.cm.Paired.colors)
                    ax.axis('equal')

                ax.set_title(f"Analisis {pilihan} - {selected_sheet}", fontsize=14, fontweight='bold')
                plt.tight_layout()

                img_buf = io.BytesIO()
                fig_pdf.savefig(img_buf, format='png')
                img_buf.seek(0)
                plt.close(fig_pdf)

                pdf = FPDF()
                pdf.add_page()
                pdf.set_font("helvetica", "B", 18)
                pdf.cell(0, 10, f"Laporan Eksekutif: {pilihan} ({selected_sheet})", new_x="LMARGIN", new_y="NEXT", align="C")
                pdf.ln(5)
                pdf.image(img_buf, x=10, w=190)
                pdf.ln(10)

                pdf.set_font("helvetica", "", 12)
                pdf.cell(0, 8, f"Rata-rata (Average): {rata_rata:,.2f}{satuan}", new_x="LMARGIN", new_y="NEXT")
                pdf.cell(0, 8, f"Total Keseluruhan: {total_semua:,.0f}{satuan}", new_x="LMARGIN", new_y="NEXT")
                pdf.cell(0, 8, f"Nilai Tertinggi: {nilai_tertinggi:,.0f}{satuan}", new_x="LMARGIN", new_y="NEXT")
                
                pdf_bytes = bytes(pdf.output())
                st.download_button(label="📥 Download PDF Sekarang", data=pdf_bytes, file_name=f"Laporan_{selected_sheet}_{pilihan.replace(' ', '_')}.pdf", mime="application/pdf")
                st.success("✅ File PDF berhasil dibuat! Silakan klik tombol Download di atas.")
                
            except Exception as e:
                st.error(f"Gagal membuat PDF. Detail error: {e}")

    except Exception as e:
        st.error(f"Gagal memproses data: Pastikan format file sesuai dengan laporan Papa. Detail: {e}")

else:
    st.markdown("""
        <div class="header-stack">
            <h1>Portal Analisis Data Anda</h1>
            <h2>Dashboard eksekutif monitoring laporan mingguan</h2>
        </div>
    """, unsafe_allow_html=True)
    c1, c2, c3 = st.columns(3)
    steps = [("📁", "1. Unggah", "Gunakan Panel Kontrol di kiri."), ("📊", "2. Pantau", "Lihat tren data via grafik."), ("📄", "3. Ekspor", "Download hasil ke PDF.")]
    for i, (icon, title, desc) in enumerate(steps):
        with [c1, c2, c3][i]:
            st.markdown(f"<div class='instruction-card'><div style='font-size: 60px;'>{icon}</div><h3>{title}</h3><p>{desc}</p></div>", unsafe_allow_html=True)
