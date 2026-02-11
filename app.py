import streamlit as st
import pandas as pd
import pdfplumber
import tabula
import PyPDF2
import numpy as np
from io import BytesIO
import tempfile
import os
import re

st.set_page_config(
    page_title="PDF to Excel Converter",
    page_icon="📄",
    layout="wide"
)

# Judul aplikasi
st.title("📄 PDF to Excel/CSV Converter")
st.markdown("Unggah file PDF berisi tabel dan konversi ke format Excel atau CSV")

# Sidebar untuk pengaturan
with st.sidebar:
    st.header("⚙️ Pengaturan Konversi")
    
    # Metode ekstraksi
    extraction_method = st.selectbox(
        "Pilih metode ekstraksi:",
        ["pdfplumber (recommended)", "tabula", "PyPDF2"]
    )
    
    # Format output
    output_format = st.selectbox(
        "Format output:",
        ["Excel (.xlsx)", "CSV (.csv)"]
    )
    
    # Halaman yang akan diekstrak
    page_option = st.radio(
        "Halaman PDF:",
        ["Semua halaman", "Halaman tertentu"]
    )
    
    if page_option == "Halaman tertentu":
        page_number = st.number_input(
            "Nomor halaman:",
            min_value=1,
            value=1,
            step=1
        )
    else:
        page_number = "all"
    
    # Opsi untuk membersihkan kolom
    st.markdown("---")
    st.header("🧹 Opsi Pembersihan")
    clean_columns = st.checkbox("Bersihkan nama kolom", value=True)
    remove_empty_columns = st.checkbox("Hapus kolom kosong", value=True)
    fill_na_values = st.checkbox("Isi nilai kosong dengan string kosong", value=True)
    
    st.markdown("---")
    st.markdown("### Cara Penggunaan:")
    st.markdown("""
    1. Unggah file PDF
    2. Pilih metode ekstraksi
    3. Atur pengaturan konversi
    4. Download hasil konversi
    """)

# Fungsi untuk membersihkan nama kolom
def clean_column_names(columns):
    cleaned_columns = []
    seen = {}
    
    for i, col in enumerate(columns):
        if col is None or pd.isna(col):
            # Untuk kolom None, beri nama generic
            base_name = f"Column_{i+1}"
            col_name = base_name
            counter = 1
            while col_name in seen:
                col_name = f"{base_name}_{counter}"
                counter += 1
        else:
            # Bersihkan string
            col_str = str(col).strip()
            # Hapus karakter khusus
            col_str = re.sub(r'[^\w\s]', '_', col_str)
            # Ganti spasi dengan underscore
            col_str = re.sub(r'\s+', '_', col_str)
            # Pastikan tidak kosong
            if not col_str:
                col_str = f"Column_{i+1}"
            
            col_name = col_str
            counter = 1
            original_name = col_name
            while col_name in seen:
                col_name = f"{original_name}_{counter}"
                counter += 1
        
        cleaned_columns.append(col_name)
        seen[col_name] = True
    
    return cleaned_columns

# Fungsi untuk membersihkan DataFrame
def clean_dataframe(df, clean_columns=True, remove_empty=True, fill_na=True):
    if df.empty:
        return df
    
    # Buat copy
    df_clean = df.copy()
    
    # 1. Bersihkan nama kolom jika ada duplikat atau None
    if clean_columns:
        df_clean.columns = clean_column_names(df_clean.columns)
    
    # 2. Hapus kolom yang sepenuhnya kosong
    if remove_empty:
        # Hapus kolom yang semua nilainya NaN atau string kosong
        cols_to_drop = []
        for col in df_clean.columns:
            if df_clean[col].dropna().empty:
                cols_to_drop.append(col)
            elif df_clean[col].astype(str).str.strip().eq('').all():
                cols_to_drop.append(col)
        
        if cols_to_drop:
            df_clean = df_clean.drop(columns=cols_to_drop)
    
    # 3. Isi nilai NaN dengan string kosong
    if fill_na:
        df_clean = df_clean.fillna('')
    
    # 4. Hapus baris yang sepenuhnya kosong
    mask = df_clean.astype(str).apply(lambda x: x.str.strip()).ne('').any(axis=1)
    df_clean = df_clean[mask].reset_index(drop=True)
    
    return df_clean

# Fungsi untuk ekstraksi tabel dengan pdfplumber (diperbaiki)
def extract_with_pdfplumber(pdf_file, page_num):
    tables = []
    with pdfplumber.open(pdf_file) as pdf:
        if page_num == "all":
            pages = pdf.pages
        else:
            pages = [pdf.pages[page_num-1]]
        
        for page_idx, page in enumerate(pages):
            page_tables = page.extract_tables()
            
            for table_idx, table in enumerate(page_tables):
                if table and len(table) > 0:  # Hanya tambahkan tabel yang tidak kosong
                    # Ambil header (baris pertama)
                    headers = table[0] if table[0] else []
                    
                    # Data (baris setelah header)
                    data_rows = table[1:] if len(table) > 1 else []
                    
                    # Buat DataFrame
                    if headers:
                        try:
                            df = pd.DataFrame(data_rows, columns=headers)
                        except ValueError as e:
                            # Jika ada masalah dengan kolom, buat kolom generic
                            st.warning(f"Masalah dengan header di tabel {table_idx+1}, halaman {page_idx+1}: {str(e)}")
                            num_cols = len(data_rows[0]) if data_rows else len(headers)
                            generic_headers = [f"Col_{i+1}" for i in range(num_cols)]
                            df = pd.DataFrame(data_rows, columns=generic_headers)
                    else:
                        # Jika tidak ada header, gunakan kolom generic
                        if data_rows:
                            num_cols = len(data_rows[0])
                            generic_headers = [f"Col_{i+1}" for i in range(num_cols)]
                            df = pd.DataFrame(data_rows, columns=generic_headers)
                        else:
                            continue  # Skip tabel kosong
                    
                    # Bersihkan DataFrame
                    df = clean_dataframe(df)
                    
                    if not df.empty:
                        # Tambahkan informasi halaman
                        df.insert(0, 'PDF_Halaman', page_idx + 1)
                        df.insert(1, 'PDF_Tabel_Index', table_idx + 1)
                        tables.append(df)
    
    return tables

# Fungsi untuk ekstraksi tabel dengan tabula (diperbaiki)
def extract_with_tabula(pdf_path, page_num):
    if page_num == "all":
        pages = "all"
    else:
        pages = page_num
    
    try:
        dfs = tabula.read_pdf(
            pdf_path, 
            pages=pages,
            multiple_tables=True,
            lattice=True,  # Untuk tabel dengan garis
            stream=True,   # Untuk tabel tanpa garis
            pandas_options={'header': None}  # Baca semua baris sebagai data
        )
        
        cleaned_dfs = []
        for idx, df in enumerate(dfs):
            if not df.empty:
                # Coba identifikasi header (baris pertama yang tidak kosong)
                df_clean = clean_dataframe(df)
                if not df_clean.empty:
                    cleaned_dfs.append(df_clean)
        
        return cleaned_dfs
    except Exception as e:
        st.error(f"Error dengan tabula: {str(e)}")
        return []

# Fungsi untuk ekstraksi dengan PyPDF2
def extract_with_pypdf2(pdf_file):
    pdf_reader = PyPDF2.PdfReader(pdf_file)
    text_content = []
    
    for page_num, page in enumerate(pdf_reader.pages):
        text = page.extract_text()
        if text.strip():
            # Split menjadi baris dan buat DataFrame
            lines = text.split('\n')
            df = pd.DataFrame(lines, columns=[f"Line"])
            df.insert(0, 'Halaman', page_num + 1)
            df.insert(1, 'Baris', range(1, len(df) + 1))
            text_content.append(df)
    
    return text_content

# Area upload file
uploaded_file = st.file_uploader(
    "📤 Unggah file PDF", 
    type=['pdf'],
    help="Maksimal ukuran file: 200MB"
)

if uploaded_file is not None:
    # Tampilkan informasi file
    file_size = uploaded_file.size / (1024 * 1024)  # Konversi ke MB
    st.info(f"📁 File: {uploaded_file.name} | Ukuran: {file_size:.2f} MB")
    
    # Buat tab untuk preview dan konversi
    tab1, tab2 = st.tabs(["👁️ Preview PDF", "🔄 Konversi"])
    
    with tab1:
        st.subheader("Preview Konten PDF")
        
        # Tampilkan halaman pertama sebagai preview
        try:
            with pdfplumber.open(uploaded_file) as pdf:
                first_page = pdf.pages[0]
                text_preview = first_page.extract_text()[:1000]  # Batasi preview
                
                col1, col2 = st.columns(2)
                with col1:
                    st.text_area(
                        "Preview Teks (Halaman 1):",
                        text_preview,
                        height=300,
                        disabled=True
                    )
                
                # Hitung jumlah halaman
                with col2:
                    st.metric("Jumlah Halaman", len(pdf.pages))
                    st.metric("Ukuran File", f"{file_size:.2f} MB")
                    
                    # Deteksi tabel sederhana
                    tables_count = 0
                    for page in pdf.pages[:3]:  # Cek 3 halaman pertama saja
                        tables = page.extract_tables()
                        tables_count += len([t for t in tables if t and len(t) > 1])
                    
                    st.metric("Perkiraan Jumlah Tabel", tables_count)
        
        except Exception as e:
            st.error(f"Error membaca PDF: {str(e)}")
    
    with tab2:
        st.subheader("Konversi ke Excel/CSV")
        
        # Tombol untuk memproses
        if st.button("🚀 Proses Konversi", type="primary", key="process_button"):
            with st.spinner("Memproses PDF... Ini mungkin memerlukan beberapa saat"):
                try:
                    # Simpan file PDF sementara
                    with tempfile.NamedTemporaryFile(delete=False, suffix='.pdf') as tmp_file:
                        tmp_file.write(uploaded_file.getvalue())
                        tmp_path = tmp_file.name
                    
                    # Ekstraksi berdasarkan metode yang dipilih
                    progress_bar = st.progress(0)
                    
                    if extraction_method == "pdfplumber (recommended)":
                        st.info("Menggunakan pdfplumber...")
                        tables = extract_with_pdfplumber(uploaded_file, page_number)
                        method_name = "pdfplumber"
                        progress_bar.progress(50)
                    
                    elif extraction_method == "tabula":
                        st.info("Menggunakan tabula... (Mungkin perlu waktu)")
                        tables = extract_with_tabula(tmp_path, page_number)
                        method_name = "tabula"
                        progress_bar.progress(50)
                    
                    else:  # PyPDF2
                        st.info("Menggunakan PyPDF2...")
                        tables = extract_with_pypdf2(uploaded_file)
                        method_name = "PyPDF2"
                        progress_bar.progress(50)
                    
                    # Hapus file temporary
                    if os.path.exists(tmp_path):
                        os.unlink(tmp_path)
                    
                    progress_bar.progress(80)
                    
                    if not tables:
                        st.warning("⚠️ Tidak ada tabel yang ditemukan dalam PDF.")
                    else:
                        progress_bar.progress(100)
                        
                        # Bersihkan semua tabel
                        cleaned_tables = []
                        for table_df in tables:
                            cleaned_df = clean_dataframe(
                                table_df, 
                                clean_columns=clean_columns,
                                remove_empty=remove_empty_columns,
                                fill_na=fill_na_values
                            )
                            if not cleaned_df.empty:
                                cleaned_tables.append(cleaned_df)
                        
                        st.success(f"✅ Berhasil mengekstrak {len(cleaned_tables)} tabel menggunakan {method_name}")
                        
                        # Tampilkan preview tabel
                        st.subheader("📊 Preview Data")
                        
                        for i, table_df in enumerate(cleaned_tables[:5]):  # Batasi preview ke 5 tabel pertama
                            with st.expander(f"Tabel {i+1} - {len(table_df)} baris × {len(table_df.columns)} kolom"):
                                # Tampilkan nama kolom
                                st.write("**Kolom:**", ", ".join(table_df.columns.tolist()))
                                
                                # Tampilkan dataframe dengan error handling
                                try:
                                    st.dataframe(
                                        table_df.head(10), 
                                        use_container_width=True,
                                        hide_index=True
                                    )
                                except Exception as e:
                                    st.error(f"Error menampilkan tabel: {str(e)}")
                                    # Tampilkan sebagai teks sebagai fallback
                                    st.write(table_df.head(10).to_string())
                                
                                # Statistik tabel
                                col1, col2, col3, col4 = st.columns(4)
                                with col1:
                                    st.metric("Baris", len(table_df))
                                with col2:
                                    st.metric("Kolom", len(table_df.columns))
                                with col3:
                                    missing_values = (table_df == '').sum().sum() + table_df.isna().sum().sum()
                                    st.metric("Nilai Kosong", missing_values)
                                with col4:
                                    duplicate_cols = len(table_df.columns) - len(set(table_df.columns))
                                    st.metric("Duplikat Kolom", duplicate_cols)
                        
                        if len(cleaned_tables) > 5:
                            st.info(f"📝 ...dan {len(cleaned_tables) - 5} tabel lainnya")
                        
                        # Konversi ke Excel atau CSV
                        st.subheader("💾 Download Hasil")
                        
                        # Pilihan untuk merge semua tabel
                        if len(cleaned_tables) > 1:
                            merge_option = st.checkbox("Gabungkan semua tabel menjadi satu sheet", value=False)
                        else:
                            merge_option = False
                        
                        if output_format == "Excel (.xlsx)":
                            # Simpan ke Excel
                            output = BytesIO()
                            
                            with pd.ExcelWriter(output, engine='openpyxl') as writer:
                                if merge_option and len(cleaned_tables) > 1:
                                    # Gabungkan semua tabel
                                    merged_df = pd.concat(cleaned_tables, ignore_index=True)
                                    merged_df.to_excel(writer, sheet_name="Merged_Data", index=False)
                                    st.info(f"📊 Semua tabel digabungkan: {len(merged_df)} baris")
                                else:
                                    # Simpan setiap tabel di sheet terpisah
                                    for i, table_df in enumerate(cleaned_tables):
                                        # Bersihkan nama sheet
                                        sheet_name = f"Table_{i+1}"
                                        sheet_name = re.sub(r'[^\w\s]', '_', sheet_name)
                                        sheet_name = sheet_name[:31]  # Excel limit
                                        
                                        table_df.to_excel(writer, sheet_name=sheet_name, index=False)
                            
                            output.seek(0)
                            
                            # Tombol download
                            filename_base = uploaded_file.name.replace('.pdf', '').replace('.PDF', '')
                            filename = f"{filename_base}_converted.xlsx"
                            
                            st.download_button(
                                label="📥 Download Excel File",
                                data=output,
                                file_name=filename,
                                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                                type="primary"
                            )
                        
                        else:  # CSV
                            # Untuk CSV, kita berikan opsi download per tabel
                            filename_base = uploaded_file.name.replace('.pdf', '').replace('.PDF', '')
                            
                            if merge_option and len(cleaned_tables) > 1:
                                # Gabungkan dan download sebagai satu CSV
                                merged_df = pd.concat(cleaned_tables, ignore_index=True)
                                csv_data = merged_df.to_csv(index=False).encode('utf-8')
                                
                                st.download_button(
                                    label="📥 Download Semua Data sebagai CSV",
                                    data=csv_data,
                                    file_name=f"{filename_base}_merged.csv",
                                    mime="text/csv",
                                    type="primary"
                                )
                            else:
                                # Download per tabel
                                for i, table_df in enumerate(cleaned_tables):
                                    csv_data = table_df.to_csv(index=False).encode('utf-8')
                                    
                                    st.download_button(
                                        label=f"📥 Download Tabel {i+1} sebagai CSV",
                                        data=csv_data,
                                        file_name=f"{filename_base}_table_{i+1}.csv",
                                        mime="text/csv"
                                    )
                        
                        # Tampilkan informasi konversi
                        st.info(f"""
                        **📋 Informasi Konversi:**
                        - Metode: {extraction_method}
                        - Jumlah tabel: {len(cleaned_tables)}
                        - Format output: {output_format}
                        - Total baris data: {sum(len(df) for df in cleaned_tables):,}
                        - Total kolom: {sum(len(df.columns) for df in cleaned_tables):,}
                        - Nama file: {uploaded_file.name}
                        """)
                
                except Exception as e:
                    st.error(f"❌ Error saat konversi: {str(e)}")
                    # Tampilkan error detail untuk debugging
                    with st.expander("Detail Error"):
                        st.exception(e)

else:
    # Tampilkan contoh UI ketika belum ada file
    st.markdown("---")
    col1, col2, col3 = st.columns(3)
    
    with col1:
        st.markdown("### 📤 Unggah PDF")
        st.markdown("Seret file PDF Anda ke area upload di atas")
    
    with col2:
        st.markdown("### ⚙️ Atur Konfigurasi")
        st.markdown("Pilih metode ekstraksi dan format output di sidebar")
    
    with col3:
        st.markdown("### 📥 Download Hasil")
        st.markdown("Download file Excel/CSV hasil konversi")

# Footer
st.markdown("---")
st.markdown(
    "<div style='text-align: center; color: gray;'>"
    "Dibuat dengan Streamlit • PDF to Excel Converter v2.0 • Perbaikan duplikat kolom"
    "</div>",
    unsafe_allow_html=True
)
