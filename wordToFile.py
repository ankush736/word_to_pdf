import streamlit as st
import tempfile
import os
import zipfile
import subprocess

st.set_page_config(page_title="Word to PDF Converter", layout="wide")

st.title("📄 Word to PDF Converter")

if "files" not in st.session_state:
    st.session_state.files = []

uploaded_files = st.file_uploader(
    "Upload Word files",
    type=["docx"],
    accept_multiple_files=True
)

# Add files
if uploaded_files:
    for file in uploaded_files:
        if file.name not in [f.name for f in st.session_state.files]:
            st.session_state.files.append(file)

st.subheader("📂 Selected Files")

# Remove option
files_to_keep = []
for i, file in enumerate(st.session_state.files):
    col1, col2 = st.columns([4, 1])

    with col1:
        st.write(f"📄 {file.name}")

    with col2:
        if st.button("❌ Remove", key=f"remove_{i}"):
            continue
        else:
            files_to_keep.append(file)

st.session_state.files = files_to_keep

# Convert
if st.session_state.files:
    if st.button("🚀 Convert to PDF & Download ZIP"):
        with tempfile.TemporaryDirectory() as tmpdir:
            pdf_paths = []

            for file in st.session_state.files:
                input_path = os.path.join(tmpdir, file.name)

                with open(input_path, "wb") as f:
                    f.write(file.getbuffer())

                # LibreOffice conversion
                try:
                    subprocess.run([
                        "libreoffice",
                        "--headless",
                        "--convert-to", "pdf",
                        "--outdir", tmpdir,
                        input_path
                    ], check=True)

                    pdf_name = file.name.replace(".docx", ".pdf")
                    pdf_path = os.path.join(tmpdir, pdf_name)

                    if os.path.exists(pdf_path):
                        pdf_paths.append(pdf_path)
                    else:
                        st.error(f"Conversion failed: {file.name}")

                except Exception as e:
                    st.error(f"Error converting {file.name}: {e}")

            # Create ZIP
            zip_path = os.path.join(tmpdir, "converted_pdfs.zip")
            with zipfile.ZipFile(zip_path, 'w') as zipf:
                for pdf in pdf_paths:
                    zipf.write(pdf, os.path.basename(pdf))

            with open(zip_path, "rb") as f:
                st.download_button(
                    "📥 Download ZIP",
                    f,
                    "converted_pdfs.zip",
                    "application/zip"
                )
else:
    st.info("Upload at least one Word file.")
