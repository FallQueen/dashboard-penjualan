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

st.markdown("""
    <style>
    .stApp { background-color: #0E1117 !important; color: #FFFFFF !important; }
    .header-stack { text-align: center; width: 100%; margin-bottom: 40px; margin-top: 20px;}
    .header-stack h1 { font-size: 55px !important; font-weight: 850 !important; color: #60A5FA !important; margin-bottom: 5px !important; }
    .header-stack h2 { font-size: 24px !important; font-weight: 400 !important; color: #CBD5E1 !important; margin-top: 0 !important; }
    [data-testid="stSidebar"] { min-width: 350px !important; background-color: #1E293B !important; }
    [data-testid="stSidebarCollapseButton"] { display: none !important; }
    .main .block-container { max-width: 95% !important; padding-top: 2rem !important; }
    .instruction-card { 
        background-color: #1E293B; padding: 30px; border-radius: 20px; 
        border: 2px solid #3B82F6; text-align: center; min-height: 280px; 
        display: flex; flex-direction: column; justify-content: center; 
        align-items: center; box-shadow: 0 4px 6px rgba(0, 0, 0, 0.3);
    }
    .instruction-card h3 { color: #60A5FA; font-size: 22px; margin-top: 15px;}
    .instruction-card p { color: #CBD5E1; font-size: 16px;}
    [data-testid="stMetric"] { background: rgba(255, 255, 255, 0.05); padding: 15px; border-radius: 10px; text-align: center; }
    </style>
    """, unsafe_allow_html=True)

# --- SIDEBAR ---
with st.sidebar:
    st.markdown("## 📥 Panel Kontrol")
    uploaded_file = st.file_uploader("Upload File Laporan", type=['xlsx', 'csv', 'xls'])
    st.markdown("---")
    
    selected_sheet = None
    if uploaded_file and uploaded_file.name.endswith(('.xlsx', '.xls')):
        excel_file = pd.ExcelFile(uploaded_file)
        if len(excel_file.sheet_names) > 1:
            st.markdown("### 📂 Pilih Halaman Excel")
            selected_sheet = st.selectbox("Pilih Wilayah (Misal: Jember/Banyuwangi):", excel_file.sheet_names)
        else:
            selected_sheet = excel_file.sheet_names[0]

# --- LOGIKA DASHBOARD ---
if uploaded_file:
    try:
        # 1. BACA EXCEL APA ADANYA (Tanpa dipotong/dibuang barisnya agar koordinat aman)
        if uploaded_file.name.endswith('.csv'):
            df_raw = pd.read_csv(uploaded_file, header=None)
        else:
            df_raw = pd.read_excel(uploaded_file, sheet_name=selected_sheet, header=None)

        # ✨ ALGORITMA SUPER FLEKSIBEL: "CELL-BY-CELL HARVESTER" ✨
        data_list = []
        
        # A. Cari baris dan kolom mana saja yang punya tulisan "Week" atau "W"
        anchor_rows = []
        for r in range(len(df_raw)):
            for c in range(len(df_raw.columns)):
                val = str(df_raw.iloc[r, c]).strip()
                if re.search(r'^(Week|Weeek|W)\s*\d+', val, re.IGNORECASE):
                    anchor_rows.append((r, c))
                    break # Ketemu 1 di baris ini, langsung lanjut ke baris berikutnya
                    
        if not anchor_rows:
            st.error("Gagal mendeteksi data. Pastikan ada tulisan 'Week' atau 'W' di dalam tabel.")
            st.stop()
            
        # B. Ekstrak data berdasarkan koordinat jangkar (Anchor)
        for i, (r, c) in enumerate(anchor_rows):
            metric_col = c - 1
            if metric_col < 0: continue # Jaga-jaga kalau error format
            
            # Batas tabel ini adalah tabel berikutnya (atau akhir file)
            end_r = anchor_rows[i+1][0] if i+1 < len(anchor_rows) else len(df_raw)
            
            # Tarik baris Minggu dan Produk, ratakan sel yang di-merge (ffill)
            week_row = df_raw.iloc[r].copy()
            week_row = week_row.replace(['nan', 'NaN', 'None', ''], pd.NA).ffill()
            prod_row = df_raw.iloc[r+1].copy().fillna('')
            
            # Panen angka per sel!
            for data_r in range(r+2, end_r):
                metric_name = str(df_raw.iloc[data_r, metric_col]).strip()
                
                # Abaikan baris kalau nama metriknya kosong
                if metric_name in ['nan', 'NaN', 'None', '', '0']: 
                    continue
                    
                for data_c in range(c, len(df_raw.columns)):
                    week_val = str(week_row[data_c]).strip()
                    prod_val = str(prod_row[data_c]).strip()
                    
                    # Cek apakah kolom ini benar-benar kolom "Week"
                    if not re.search(r'^(Week|Weeek|W)\s*\d+', week_val, re.IGNORECASE):
                        continue
                        
                    val = df_raw.iloc[data_r, data_c]
                    
                    # Simpan ke database internal
                    data_list.append({
                        'Minggu': week_val,
                        'Produk': prod_val,
                        'Metrik': metric_name,
                        'Nilai': val
                    })
                    
        if not data_list:
            st.error("Gagal memproses metrik. Format tidak sesuai.")
            st.stop()

        # 2. RAKIT DATABASE JADI TABEL RAPI
        df_melted = pd.DataFrame(data_list)
        
        # Bersihkan nama minggu ("W 06" jadi "Week 6")
        def clean_week(s):
            match = re.search(r'(?:Week|Weeek|W)\s*(\d+)', str(s), flags=re.IGNORECASE)
            return f"Week {int(match.group(1))}" if match else str(s)
        df_melted['Minggu'] = df_melted['Minggu'].apply(clean_week)
        
        # Bersihkan angka dan isi kosong dengan 0
        def clean_numeric(val):
            if isinstance(val, str):
                val = val.replace('%', '').replace(',', '').strip()
                if val in ['#DIV/0!', '-', '']: 
                    return 0
            return pd.to_numeric(val, errors='coerce')
        df_melted['Nilai'] = df_melted['Nilai'].apply(clean_numeric).fillna(0)
        
        # 3. PIVOT (Buat tabel siap pakai untuk aplikasi)
        df_final = df_melted.pivot_table(index=['Minggu', 'Produk'], columns='Metrik', values='Nilai', aggfunc='first').reset_index()
        
        def get_week_num(w):
            match = re.search(r'\d+', str(w))
            return int(match.group(0)) if match else 0
        df_final['WeekNum'] = df_final['Minggu'].apply(get_week_num)
        df_final = df_final.sort_values(['WeekNum', 'Produk']).drop(columns=['WeekNum']).reset_index(drop=True)

        # --- UI DASHBOARD ---
        st.markdown(f"<h1 style='text-align: center; color: #60A5FA; font-size: 50px;'>📊 Laporan: {selected_sheet}</h1>", unsafe_allow_html=True)
        st.markdown("---")

        col_opt1, col_opt2 = st.columns(2)
        with col_opt1:
            pilihan = st.selectbox("🎯 Pilih Metrik Penjualan:", [c for c in df_final.columns if c not in ['Minggu', 'Produk']])
        with col_opt2:
            jenis_grafik = st.selectbox("📈 Pilih Bentuk Diagram:", ["Diagram Batang Berdampingan (Grouped Bar)", "Diagram Garis Tren (Line)", "Diagram Lingkaran (Pie)"])

        rata_rata = df_final[pilihan].mean()
        total_semua = df_final[pilihan].sum()
        nilai_tertinggi = df_final[pilihan].max()

        st.markdown(f"### 💡 Ringkasan Eksekutif - {selected_sheet}")
        c_sum1, c_sum2, c_sum3 = st.columns(3)
        satuan = "%" if "%" in pilihan else ""
        c_sum1.metric(label="Rata-rata (Average)", value=f"{rata_rata:,.2f}{satuan}")
        c_sum2.metric(label="Total Keseluruhan", value=f"{total_semua:,.0f}{satuan}")
        c_sum3.metric(label="Nilai Tertinggi", value=f"{nilai_tertinggi:,.0f}{satuan}")
        
        st.markdown("<br>", unsafe_allow_html=True)

        # GRAFIK INTERAKTIF
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
        st.markdown(f"### 📋 Tabel Hasil Konversi - {selected_sheet}")
        st.dataframe(df_final, use_container_width=True, height=400, hide_index=True)

        # EKSPOR PDF
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
                pdf.cell(0, 8, f"Rata-rata: {rata_rata:,.2f}{satuan}", new_x="LMARGIN", new_y="NEXT")
                pdf.cell(0, 8, f"Total: {total_semua:,.0f}{satuan}", new_x="LMARGIN", new_y="NEXT")
                pdf.cell(0, 8, f"Tertinggi: {nilai_tertinggi:,.0f}{satuan}", new_x="LMARGIN", new_y="NEXT")
                
                pdf_bytes = bytes(pdf.output())
                st.download_button(label="📥 Download PDF Sekarang", data=pdf_bytes, file_name=f"Laporan_{selected_sheet}_{pilihan.replace(' ', '_')}.pdf", mime="application/pdf")
                st.success("✅ File PDF berhasil dibuat!")
                
            except Exception as e:
                st.error(f"Gagal membuat PDF. Detail error: {e}")

    except Exception as e:
        st.error(f"Terjadi kesalahan sistem: {e}. Pastikan file Excel sesuai standar.")

# --- TAMPILAN AWAL SEBELUM UPLOAD ---
else:
    st.markdown("""
        <div class="header-stack">
            <h1>Portal Analisis Data Anda</h1>
            <h2>Dashboard eksekutif monitoring laporan mingguan</h2>
        </div>
    """, unsafe_allow_html=True)
    
    c1, c2, c3 = st.columns(3)
    steps = [
        ("📁", "1. Unggah", "Gunakan Panel Kontrol di kiri untuk memasukkan file laporan."),
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
