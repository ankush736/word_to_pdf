import streamlit as st
from docx2pdf import convert
import tempfile
import os
import zipfile

st.set_page_config(page_title="Word to PDF Converter", layout="wide")

st.title("📄 Word to PDF Converter")

# Session state for uploaded files
if "files" not in st.session_state:
    st.session_state.files = []

uploaded_files = st.file_uploader(
    "Upload Word files",
    type=["docx"],
    accept_multiple_files=True
)

# Add new files
if uploaded_files:
    for file in uploaded_files:
        if file.name not in [f.name for f in st.session_state.files]:
            st.session_state.files.append(file)

st.subheader("📂 Selected Files")

# Show files with remove option
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

# Convert button
if st.session_state.files:
    if st.button("🚀 Convert to PDF & Download ZIP"):
        with tempfile.TemporaryDirectory() as tmpdir:
            pdf_paths = []

            for file in st.session_state.files:
                # Save uploaded file temporarily
                input_path = os.path.join(tmpdir, file.name)
                with open(input_path, "wb") as f:
                    f.write(file.getbuffer())

                # Output PDF path
                pdf_name = file.name.replace(".docx", ".pdf")
                output_path = os.path.join(tmpdir, pdf_name)

                # Convert
                convert(input_path, output_path)
                pdf_paths.append(output_path)

            # Create ZIP
            zip_path = os.path.join(tmpdir, "converted_pdfs.zip")
            with zipfile.ZipFile(zip_path, 'w') as zipf:
                for pdf in pdf_paths:
                    zipf.write(pdf, os.path.basename(pdf))

            # Download button
            with open(zip_path, "rb") as f:
                st.download_button(
                    label="📥 Download ZIP",
                    data=f,
                    file_name="converted_pdfs.zip",
                    mime="application/zip"
                )
else:
    st.info("Upload at least one Word file to proceed.")
