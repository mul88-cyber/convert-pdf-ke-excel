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
from typing import List, Dict, Tuple
import time

st.set_page_config(
    page_title="PDF to Excel Converter - Deteksi Tabel",
    page_icon="📊",
    layout="wide"
)

# Judul aplikasi
st.title("📊 PDF to Excel/CSV Converter dengan Deteksi Tabel")
st.markdown("Unggah file PDF, deteksi halaman yang berisi tabel, dan konversi hanya tabelnya saja")

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
    
    # Mode pemilihan halaman
    page_mode = st.radio(
        "Mode pemilihan halaman:",
        ["Otomatis deteksi tabel", "Manual pilih halaman", "Semua halaman"]
    )
    
    # Opsi untuk membersihkan kolom
    st.markdown("---")
    st.header("🧹 Opsi Pembersihan")
    clean_columns = st.checkbox("Bersihkan nama kolom", value=True)
    remove_empty_columns = st.checkbox("Hapus kolom kosong", value=True)
    fill_na_values = st.checkbox("Isi nilai kosong dengan string kosong", value=True)
    
    # Threshold untuk deteksi tabel
    st.markdown("---")
    st.header("🔍 Pengaturan Deteksi")
    table_threshold = st.slider(
        "Sensitivitas deteksi tabel:", 
        min_value=1, 
        max_value=10, 
        value=3,
        help="Nilai lebih tinggi = hanya deteksi tabel yang lebih jelas"
    )
    
    st.markdown("---")
    st.markdown("### Cara Penggunaan:")
    st.markdown("""
    1. Unggah file PDF
    2. Sistem akan otomatis deteksi halaman berisi tabel
    3. Pilih halaman yang ingin dikonversi
    4. Download hasil konversi
    """)

# Fungsi untuk mendeteksi halaman yang mengandung tabel
def detect_tables_in_pdf(pdf_file, threshold: int = 3) -> Dict[int, List[Dict]]:
    """
    Mendeteksi halaman yang mengandung tabel dalam PDF
    Returns: Dictionary {page_number: [table_info]}
    """
    tables_by_page = {}
    
    with pdfplumber.open(pdf_file) as pdf:
        total_pages = len(pdf.pages)
        
        # Progress bar
        progress_bar = st.progress(0)
        status_text = st.empty()
        
        for page_num in range(total_pages):
            status_text.text(f"Menganalisis halaman {page_num + 1} dari {total_pages}...")
            progress_bar.progress((page_num + 1) / total_pages)
            
            page = pdf.pages[page_num]
            
            # Ekstrak tabel dari halaman
            tables = page.extract_tables()
            
            # Filter tabel yang valid (memiliki baris dan kolom)
            valid_tables = []
            for table_idx, table in enumerate(tables):
                if table and len(table) > 1:  # Minimal ada header dan satu baris data
                    num_rows = len(table)
                    num_cols = max(len(row) for row in table) if table else 0
                    
                    # Hitung rasio sel yang terisi (sebagai indikator kualitas tabel)
                    filled_cells = sum(1 for row in table for cell in row if cell and str(cell).strip())
                    total_cells = num_rows * num_cols if num_cols > 0 else 0
                    fill_ratio = filled_cells / total_cells if total_cells > 0 else 0
                    
                    # Gunakan threshold untuk menentukan apakah ini tabel yang valid
                    if (num_rows >= threshold and 
                        num_cols >= 2 and 
                        fill_ratio > 0.3):  # Minimal 30% sel terisi
                        
                        table_info = {
                            'index': table_idx,
                            'rows': num_rows,
                            'cols': num_cols,
                            'fill_ratio': fill_ratio,
                            'preview_data': table[:3]  # Preview 3 baris pertama
                        }
                        valid_tables.append(table_info)
            
            if valid_tables:
                tables_by_page[page_num + 1] = valid_tables
        
        progress_bar.empty()
        status_text.empty()
    
    return tables_by_page

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
            # Hapus underscore berlebih di awal/akhir
            col_str = col_str.strip('_')
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

# Fungsi untuk ekstraksi tabel dari halaman tertentu
def extract_tables_from_pages(pdf_file, pages_to_extract, extraction_method):
    all_tables = []
    
    if extraction_method == "pdfplumber (recommended)":
        with pdfplumber.open(pdf_file) as pdf:
            for page_num in pages_to_extract:
                page_idx = page_num - 1
                if page_idx < len(pdf.pages):
                    page = pdf.pages[page_idx]
                    tables = page.extract_tables()
                    
                    for table_idx, table in enumerate(tables):
                        if table and len(table) > 0:
                            # Ambil header (baris pertama)
                            headers = table[0] if table[0] else []
                            data_rows = table[1:] if len(table) > 1 else []
                            
                            if data_rows:
                                try:
                                    if headers:
                                        df = pd.DataFrame(data_rows, columns=headers)
                                    else:
                                        num_cols = len(data_rows[0])
                                        generic_headers = [f"Col_{i+1}" for i in range(num_cols)]
                                        df = pd.DataFrame(data_rows, columns=generic_headers)
                                    
                                    df = clean_dataframe(df)
                                    if not df.empty:
                                        df.insert(0, 'PDF_Halaman', page_num)
                                        df.insert(1, 'PDF_Tabel_Index', table_idx + 1)
                                        all_tables.append(df)
                                except Exception as e:
                                    st.warning(f"Error di halaman {page_num}, tabel {table_idx+1}: {str(e)}")
    
    elif extraction_method == "tabula":
        # Simpan file sementara untuk tabula
        with tempfile.NamedTemporaryFile(delete=False, suffix='.pdf') as tmp_file:
            tmp_file.write(pdf_file.getvalue())
            tmp_path = tmp_file.name
        
        try:
            for page_num in pages_to_extract:
                dfs = tabula.read_pdf(
                    tmp_path,
                    pages=page_num,
                    multiple_tables=True,
                    lattice=True,
                    stream=True,
                    pandas_options={'header': None}
                )
                
                for idx, df in enumerate(dfs):
                    if not df.empty:
                        df_clean = clean_dataframe(df)
                        if not df_clean.empty:
                            df_clean.insert(0, 'PDF_Halaman', page_num)
                            df_clean.insert(1, 'PDF_Tabel_Index', idx + 1)
                            all_tables.append(df_clean)
        
        finally:
            if os.path.exists(tmp_path):
                os.unlink(tmp_path)
    
    else:  # PyPDF2
        pdf_reader = PyPDF2.PdfReader(pdf_file)
        for page_num in pages_to_extract:
            page_idx = page_num - 1
            if page_idx < len(pdf_reader.pages):
                page = pdf_reader.pages[page_idx]
                text = page.extract_text()
                if text.strip():
                    lines = text.split('\n')
                    df = pd.DataFrame(lines, columns=['Konten'])
                    df.insert(0, 'PDF_Halaman', page_num)
                    all_tables.append(df)
    
    return all_tables

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
    
    # Tab untuk navigasi
    tab1, tab2, tab3 = st.tabs(["🔍 Deteksi Tabel", "👁️ Preview PDF", "🔄 Konversi"])
    
    with tab1:
        st.subheader("Deteksi Halaman Berisi Tabel")
        
        if st.button("🔎 Mulai Deteksi Tabel", type="primary"):
            with st.spinner("Mendeteksi tabel dalam PDF..."):
                try:
                    # Deteksi tabel
                    tables_by_page = detect_tables_in_pdf(uploaded_file, table_threshold)
                    
                    # Simpan ke session state
                    st.session_state['tables_by_page'] = tables_by_page
                    st.session_state['total_pages'] = len(pdfplumber.open(uploaded_file).pages)
                    
                    if not tables_by_page:
                        st.warning("❌ Tidak ada tabel yang terdeteksi dalam PDF.")
                    else:
                        st.success(f"✅ Ditemukan tabel di {len(tables_by_page)} halaman dari total {st.session_state['total_pages']} halaman")
                        
                        # Tampilkan hasil deteksi
                        st.subheader("📋 Hasil Deteksi Tabel")
                        
                        for page_num, tables in tables_by_page.items():
                            with st.expander(f"Halaman {page_num} - {len(tables)} tabel ditemukan"):
                                for table_info in tables:
                                    st.write(f"**Tabel {table_info['index'] + 1}:**")
                                    col1, col2, col3 = st.columns(3)
                                    with col1:
                                        st.metric("Baris", table_info['rows'])
                                    with col2:
                                        st.metric("Kolom", table_info['cols'])
                                    with col3:
                                        st.metric("Kepadatan", f"{table_info['fill_ratio']*100:.1f}%")
                                    
                                    # Tampilkan preview kecil
                                    if table_info['preview_data']:
                                        preview_df = pd.DataFrame(table_info['preview_data'][1:], 
                                                                  columns=table_info['preview_data'][0] if table_info['preview_data'][0] else [])
                                        st.dataframe(preview_df, height=120, hide_index=True)
                        
                        # Pilihan halaman untuk konversi
                        st.subheader("🎯 Pilih Halaman untuk Konversi")
                        
                        if 'tables_by_page' in st.session_state:
                            # Default pilih semua halaman dengan tabel
                            default_pages = list(tables_by_page.keys())
                            
                            selected_pages = st.multiselect(
                                "Pilih halaman yang akan dikonversi:",
                                options=range(1, st.session_state['total_pages'] + 1),
                                default=default_pages,
                                format_func=lambda x: f"Halaman {x} {'📊' if x in tables_by_page else '📄'}"
                            )
                            
                            # Simpan ke session state
                            st.session_state['selected_pages'] = selected_pages
                            
                            # Tampilkan statistik
                            col1, col2 = st.columns(2)
                            with col1:
                                st.metric("Halaman dipilih", len(selected_pages))
                            with col2:
                                tables_count = sum(len(tables_by_page.get(p, [])) for p in selected_pages)
                                st.metric("Total tabel", tables_count)
                            
                            st.success(f"✅ Siap mengkonversi {len(selected_pages)} halaman yang dipilih")
                
                except Exception as e:
                    st.error(f"Error saat mendeteksi tabel: {str(e)}")
        
        elif 'tables_by_page' in st.session_state:
            # Tampilkan hasil deteksi yang sudah ada
            st.success(f"✅ Hasil deteksi tersedia: {len(st.session_state['tables_by_page'])} halaman berisi tabel")
            
            # Pilihan halaman untuk konversi
            st.subheader("🎯 Pilih Halaman untuk Konversi")
            
            tables_by_page = st.session_state['tables_by_page']
            total_pages = st.session_state['total_pages']
            
            # Default pilih semua halaman dengan tabel
            default_pages = list(tables_by_page.keys())
            
            selected_pages = st.multiselect(
                "Pilih halaman yang akan dikonversi:",
                options=range(1, total_pages + 1),
                default=default_pages,
                format_func=lambda x: f"Halaman {x} {'📊' if x in tables_by_page else '📄'}"
            )
            
            # Simpan ke session state
            st.session_state['selected_pages'] = selected_pages
    
    with tab2:
        st.subheader("Preview Konten PDF")
        
        # Pilih halaman untuk preview
        if 'total_pages' in st.session_state:
            page_to_preview = st.selectbox(
                "Pilih halaman untuk preview:",
                options=range(1, st.session_state['total_pages'] + 1),
                format_func=lambda x: f"Halaman {x}"
            )
        else:
            page_to_preview = 1
        
        # Tampilkan preview halaman
        try:
            with pdfplumber.open(uploaded_file) as pdf:
                if page_to_preview <= len(pdf.pages):
                    page = pdf.pages[page_to_preview - 1]
                    text_preview = page.extract_text()[:1500]  # Batasi preview
                    
                    col1, col2 = st.columns([2, 1])
                    with col1:
                        st.text_area(
                            f"Preview Teks Halaman {page_to_preview}:",
                            text_preview,
                            height=400,
                            disabled=True
                        )
                    
                    with col2:
                        st.metric("Halaman", page_to_preview)
                        
                        # Cek apakah halaman ini ada tabel
                        if 'tables_by_page' in st.session_state:
                            tables_in_page = st.session_state['tables_by_page'].get(page_to_preview, [])
                            st.metric("Tabel terdeteksi", len(tables_in_page))
                        
                        # Statistik halaman
                        words = len(text_preview.split())
                        lines = len(text_preview.split('\n'))
                        
                        col2.metric("Kata", words)
                        col2.metric("Baris", lines)
        
        except Exception as e:
            st.error(f"Error membaca PDF: {str(e)}")
    
    with tab3:
        st.subheader("Konversi ke Excel/CSV")
        
        # Cek apakah sudah memilih halaman
        if 'selected_pages' not in st.session_state or not st.session_state['selected_pages']:
            st.warning("⚠️ Silakan pilih halaman terlebih dahulu di tab 'Deteksi Tabel'")
            st.info("Klik tab '🔍 Deteksi Tabel' untuk mendeteksi dan memilih halaman berisi tabel")
        else:
            st.success(f"✅ {len(st.session_state['selected_pages'])} halaman dipilih untuk konversi")
            
            # Tampilkan halaman yang dipilih
            with st.expander("📋 Halaman yang akan dikonversi"):
                selected_pages = st.session_state['selected_pages']
                tables_by_page = st.session_state.get('tables_by_page', {})
                
                for page_num in selected_pages:
                    tables_count = len(tables_by_page.get(page_num, []))
                    st.write(f"**Halaman {page_num}:** {tables_count} tabel")
            
            # Tombol konversi
            if st.button("🚀 Mulai Konversi", type="primary"):
                with st.spinner(f"Mengkonversi {len(st.session_state['selected_pages'])} halaman..."):
                    try:
                        # Ekstrak tabel dari halaman yang dipilih
                        tables = extract_tables_from_pages(
                            uploaded_file,
                            st.session_state['selected_pages'],
                            extraction_method
                        )
                        
                        if not tables:
                            st.warning("⚠️ Tidak ada tabel yang berhasil diekstrak dari halaman yang dipilih")
                        else:
                            st.success(f"✅ Berhasil mengekstrak {len(tables)} tabel")
                            
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
                            
                            # Tampilkan preview tabel
                            st.subheader("📊 Preview Data Hasil Konversi")
                            
                            for i, table_df in enumerate(cleaned_tables[:3]):  # Batasi preview ke 3 tabel pertama
                                with st.expander(f"Tabel {i+1} - Halaman {table_df.iloc[0]['PDF_Halaman'] if 'PDF_Halaman' in table_df.columns else 'N/A'}"):
                                    # Tampilkan informasi tabel
                                    col1, col2, col3 = st.columns(3)
                                    with col1:
                                        st.metric("Baris", len(table_df))
                                    with col2:
                                        st.metric("Kolom", len(table_df.columns))
                                    with col3:
                                        halaman = table_df.iloc[0]['PDF_Halaman'] if 'PDF_Halaman' in table_df.columns else 'N/A'
                                        st.metric("Halaman PDF", halaman)
                                    
                                    # Tampilkan dataframe
                                    st.dataframe(
                                        table_df.head(10), 
                                        use_container_width=True,
                                        hide_index=True
                                    )
                            
                            if len(cleaned_tables) > 3:
                                st.info(f"📝 ...dan {len(cleaned_tables) - 3} tabel lainnya")
                            
                            # Konversi ke Excel atau CSV
                            st.subheader("💾 Download Hasil")
                            
                            # Pilihan untuk merge semua tabel
                            if len(cleaned_tables) > 1:
                                merge_option = st.checkbox("Gabungkan semua tabel menjadi satu sheet", value=True)
                            else:
                                merge_option = False
                            
                            filename_base = uploaded_file.name.replace('.pdf', '').replace('.PDF', '')
                            
                            if output_format == "Excel (.xlsx)":
                                # Simpan ke Excel
                                output = BytesIO()
                                
                                with pd.ExcelWriter(output, engine='openpyxl') as writer:
                                    if merge_option and len(cleaned_tables) > 1:
                                        # Gabungkan semua tabel
                                        merged_df = pd.concat(cleaned_tables, ignore_index=True)
                                        merged_df.to_excel(writer, sheet_name="Data_Terpisah", index=False)
                                        sheet_info = f"1 sheet (tergabung)"
                                    else:
                                        # Simpan setiap tabel di sheet terpisah
                                        for i, table_df in enumerate(cleaned_tables):
                                            # Cari halaman untuk nama sheet
                                            halaman = table_df.iloc[0]['PDF_Halaman'] if 'PDF_Halaman' in table_df.columns else i+1
                                            sheet_name = f"H{halaman}_T{i+1}"
                                            sheet_name = re.sub(r'[^\w\s]', '_', sheet_name)
                                            sheet_name = sheet_name[:31]  # Excel limit
                                            
                                            table_df.to_excel(writer, sheet_name=sheet_name, index=False)
                                        sheet_info = f"{len(cleaned_tables)} sheets"
                                
                                output.seek(0)
                                
                                # Tombol download
                                filename = f"{filename_base}_tables_only.xlsx"
                                
                                st.download_button(
                                    label=f"📥 Download Excel File ({sheet_info})",
                                    data=output,
                                    file_name=filename,
                                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                                    type="primary"
                                )
                            
                            else:  # CSV
                                if merge_option and len(cleaned_tables) > 1:
                                    # Gabungkan dan download sebagai satu CSV
                                    merged_df = pd.concat(cleaned_tables, ignore_index=True)
                                    csv_data = merged_df.to_csv(index=False).encode('utf-8')
                                    
                                    st.download_button(
                                        label="📥 Download Semua Data sebagai CSV",
                                        data=csv_data,
                                        file_name=f"{filename_base}_all_tables.csv",
                                        mime="text/csv",
                                        type="primary"
                                    )
                                else:
                                    # Download per tabel
                                    for i, table_df in enumerate(cleaned_tables):
                                        csv_data = table_df.to_csv(index=False).encode('utf-8')
                                        halaman = table_df.iloc[0]['PDF_Halaman'] if 'PDF_Halaman' in table_df.columns else i+1
                                        
                                        st.download_button(
                                            label=f"📥 Download Tabel {i+1} (Halaman {halaman})",
                                            data=csv_data,
                                            file_name=f"{filename_base}_halaman_{halaman}_tabel_{i+1}.csv",
                                            mime="text/csv"
                                        )
                            
                            # Statistik konversi
                            st.info(f"""
                            **📋 Statistik Konversi:**
                            - Halaman diproses: {len(st.session_state['selected_pages'])}
                            - Total tabel: {len(cleaned_tables)}
                            - Total baris: {sum(len(df) for df in cleaned_tables):,}
                            - Total kolom: {sum(len(df.columns) for df in cleaned_tables):,}
                            - Metode: {extraction_method}
                            """)
                    
                    except Exception as e:
                        st.error(f"❌ Error saat konversi: {str(e)}")
                        with st.expander("Detail Error"):
                            st.exception(e)

else:
    # Tampilkan petunjuk penggunaan
    st.markdown("---")
    col1, col2, col3 = st.columns(3)
    
    with col1:
        st.markdown("### 📤 Unggah PDF")
        st.markdown("Unggah file PDF yang ingin dikonversi")
    
    with col2:
        st.markdown("### 🔍 Deteksi Tabel")
        st.markdown("Sistem akan otomatis mendeteksi halaman berisi tabel")
    
    with col3:
        st.markdown("### 🎯 Pilih & Konversi")
        st.markdown("Pilih hanya halaman dengan tabel untuk dikonversi")

# Footer
st.markdown("---")
st.markdown(
    "<div style='text-align: center; color: gray;'>"
    "Dibuat dengan Streamlit • PDF to Excel Converter v3.0 • Deteksi Tabel Otomatis"
    "</div>",
    unsafe_allow_html=True
)
