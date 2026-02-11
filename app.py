import streamlit as st
import pandas as pd
import pdfplumber
import tabula
import PyPDF2
from io import BytesIO
import tempfile
import os

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
    
    st.markdown("---")
    st.markdown("### Cara Penggunaan:")
    st.markdown("""
    1. Unggah file PDF
    2. Pilih metode ekstraksi
    3. Atur pengaturan konversi
    4. Download hasil konversi
    """)

# Fungsi untuk ekstraksi tabel dengan pdfplumber
def extract_with_pdfplumber(pdf_file, page_num):
    tables = []
    with pdfplumber.open(pdf_file) as pdf:
        if page_num == "all":
            pages = pdf.pages
        else:
            pages = [pdf.pages[page_num-1]]
        
        for page in pages:
            page_tables = page.extract_tables()
            for table in page_tables:
                if table:  # Hanya tambahkan tabel yang tidak kosong
                    df = pd.DataFrame(table[1:], columns=table[0])
                    tables.append(df)
    return tables

# Fungsi untuk ekstraksi tabel dengan tabula
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
            lattice=True  # Untuk tabel dengan garis
        )
        return dfs
    except Exception as e:
        st.error(f"Error dengan tabula: {str(e)}")
        return []

# Fungsi untuk ekstraksi dengan PyPDF2 (lebih sederhana)
def extract_with_pypdf2(pdf_file):
    pdf_reader = PyPDF2.PdfReader(pdf_file)
    text_content = []
    
    for page in pdf_reader.pages:
        text = page.extract_text()
        text_content.append(text)
    
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
                    
                    # Ekstrak dan tampilkan jumlah tabel
                    tables_count = 0
                    for page in pdf.pages:
                        tables = page.extract_tables()
                        tables_count += len([t for t in tables if t])
                    
                    st.metric("Jumlah Tabel Terdeteksi", tables_count)
        
        except Exception as e:
            st.error(f"Error membaca PDF: {str(e)}")
    
    with tab2:
        st.subheader("Konversi ke Excel/CSV")
        
        if st.button("🚀 Proses Konversi", type="primary"):
            with st.spinner("Memproses PDF..."):
                try:
                    # Simpan file PDF sementara
                    with tempfile.NamedTemporaryFile(delete=False, suffix='.pdf') as tmp_file:
                        tmp_file.write(uploaded_file.getvalue())
                        tmp_path = tmp_file.name
                    
                    # Ekstraksi berdasarkan metode yang dipilih
                    if extraction_method == "pdfplumber (recommended)":
                        tables = extract_with_pdfplumber(uploaded_file, page_number)
                        method_name = "pdfplumber"
                    
                    elif extraction_method == "tabula":
                        tables = extract_with_tabula(tmp_path, page_number)
                        method_name = "tabula"
                    
                    else:  # PyPDF2
                        text_content = extract_with_pypdf2(uploaded_file)
                        # Untuk PyPDF2, kita buat dataframe dari teks
                        tables = []
                        for i, text in enumerate(text_content):
                            lines = text.split('\n')
                            df = pd.DataFrame(lines, columns=[f"Line_{i+1}"])
                            tables.append(df)
                        method_name = "PyPDF2"
                    
                    # Hapus file temporary
                    os.unlink(tmp_path)
                    
                    if not tables:
                        st.warning("⚠️ Tidak ada tabel yang ditemukan dalam PDF.")
                    else:
                        st.success(f"✅ Berhasil mengekstrak {len(tables)} tabel menggunakan {method_name}")
                        
                        # Tampilkan preview tabel
                        st.subheader("📊 Preview Data")
                        
                        for i, table_df in enumerate(tables):
                            with st.expander(f"Tabel {i+1} - {len(table_df)} baris × {len(table_df.columns)} kolom"):
                                st.dataframe(table_df.head(10), use_container_width=True)
                                
                                # Statistik tabel
                                col1, col2, col3 = st.columns(3)
                                with col1:
                                    st.metric("Baris", len(table_df))
                                with col2:
                                    st.metric("Kolom", len(table_df.columns))
                                with col3:
                                    missing_values = table_df.isnull().sum().sum()
                                    st.metric("Missing Values", missing_values)
                        
                        # Konversi ke Excel atau CSV
                        st.subheader("💾 Download Hasil")
                        
                        if output_format == "Excel (.xlsx)":
                            # Simpan semua tabel ke satu file Excel dengan sheet berbeda
                            output = BytesIO()
                            with pd.ExcelWriter(output, engine='openpyxl') as writer:
                                for i, table_df in enumerate(tables):
                                    # Bersihkan nama sheet (max 31 karakter, tidak boleh ada karakter khusus)
                                    sheet_name = f"Table_{i+1}"[:31]
                                    table_df.to_excel(writer, sheet_name=sheet_name, index=False)
                            
                            output.seek(0)
                            
                            # Tombol download
                            st.download_button(
                                label="📥 Download Excel File",
                                data=output,
                                file_name=f"{uploaded_file.name.replace('.pdf', '')}_converted.xlsx",
                                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                                type="primary"
                            )
                        
                        else:  # CSV
                            # Untuk CSV, kita berikan opsi download per tabel
                            for i, table_df in enumerate(tables):
                                csv_data = table_df.to_csv(index=False).encode('utf-8')
                                
                                st.download_button(
                                    label=f"📥 Download Tabel {i+1} sebagai CSV",
                                    data=csv_data,
                                    file_name=f"{uploaded_file.name.replace('.pdf', '')}_table_{i+1}.csv",
                                    mime="text/csv"
                                )
                        
                        # Tampilkan informasi konversi
                        st.info(f"""
                        **Informasi Konversi:**
                        - Metode: {extraction_method}
                        - Jumlah tabel: {len(tables)}
                        - Format output: {output_format}
                        - Total baris data: {sum(len(df) for df in tables):,}
                        """)
                
                except Exception as e:
                    st.error(f"❌ Error saat konversi: {str(e)}")
                    st.exception(e)

    # Tips untuk hasil terbaik
    with st.expander("💡 Tips untuk hasil konversi terbaik"):
        st.markdown("""
        1. **PDF dengan garis tabel**: Gunakan metode **pdfplumber** atau **tabula** dengan opsi `lattice=True`
        2. **PDF tanpa garis tabel**: Gunakan metode **tabula** dengan opsi `stream=True`
        3. **PDF hasil scan**: Konversi ke PDF yang bisa dibaca teks terlebih dahulu
        4. **Struktur kompleks**: Coba beberapa metode untuk hasil terbaik
        5. **Periksa hasil**: Selalu periksa hasil konversi karena struktur PDF bisa bervariasi
        """)

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
    "Dibuat dengan Streamlit • PDF to Excel Converter v1.0"
    "</div>",
    unsafe_allow_html=True
)
