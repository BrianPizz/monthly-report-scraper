import streamlit as st
import pandas as pd
import io
import re
import sys
from datetime import datetime
from pathlib import Path
import tempfile
import os
import zipfile
import random

import pytesseract
from PIL import Image
from pymupdf import pymupdf
from openpyxl import Workbook, load_workbook

# Import functions from main.py
from main import (
    School, COLUMNS, es_and_hs_schools, index_labels,
    _blank_if_zero, _get, apply_ocr_corrections, extract_data_from_ocr,
    school_from_pdf, build_rows_for_school, get_versioned_output_path
)

# Page configuration
st.set_page_config(
    page_title="OH SPED Monthly Report Scraper",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded"
)

# Custom CSS for better styling
st.markdown("""
<style>
    .main-header {
        font-size: 2.5rem;
        font-weight: bold;
        color: #1f77b4;
        text-align: center;
        margin-bottom: 2rem;
    }
    .metric-card {
        background-color: #f0f2f6;
        padding: 1rem;
        border-radius: 0.5rem;
        border-left: 4px solid #1f77b4;
    }
    .success-message {
        background-color: #d4edda;
        color: #155724;
        padding: 1rem;
        border-radius: 0.5rem;
        border: 1px solid #c3e6cb;
    }
    .error-message {
        background-color: #f8d7da;
        color: #721c24;
        padding: 1rem;
        border-radius: 0.5rem;
        border: 1px solid #f5c6cb;
    }
</style>
""", unsafe_allow_html=True)

def main():
    # Header
    st.markdown('<h1 class="main-header">Ohio SPED Monthly Report Scraper</h1>', unsafe_allow_html=True)
    st.markdown("---")
    
    # Sidebar
    with st.sidebar:
        st.header("⚙️ Configuration")
        
        # Processing options
        st.subheader("Processing Options")
        verbose_mode = st.checkbox("Verbose Mode", help="Show detailed processing information")
        use_ocr = st.checkbox("Use OCR Processing", value=True, help="Use OCR for better text extraction")
    
    # Main content area - Drag and Drop Upload (Primary Method)
    st.header("📤 Upload Files")
    st.markdown("**Drag and drop your PDF files or ZIP folder here, or click to browse**")
    
    # Large, prominent file uploader in main area
    uploaded_files = st.file_uploader(
        "Upload PDF files or ZIP folders",
        type=['pdf', 'zip'],
        accept_multiple_files=True,
        help="Upload one or more PDF files or a ZIP folder containing PDF files",
        label_visibility="collapsed"
    )
    
    # Expand ZIPs into PDFs, keep PDFs as-is
    expanded_files = []
    
    if uploaded_files:
        for uf in uploaded_files:
            name_lower = uf.name.lower()
            if name_lower.endswith('.zip'):
                try:
                    # Read ZIP file content
                    # Convert memoryview to bytes, then to BytesIO
                    zip_bytes = bytes(uf.getbuffer())
                    zip_buffer = io.BytesIO(zip_bytes)
                    zf = zipfile.ZipFile(zip_buffer)
                    
                    # Count PDFs in ZIP, filtering out directories and duplicates
                    # Use a set to track filenames we've already processed to avoid duplicates
                    seen_filenames = set()
                    pdf_members = []
                    
                    for member in zf.namelist():
                        # Skip directories
                        if member.endswith('/'):
                            continue
                        
                        # Get just the filename (not the full path)
                        filename = Path(member).name
                        
                        # Skip macOS resource fork files (start with ._)
                        if filename.startswith('._'):
                            continue
                        
                        # Skip other macOS system files
                        if filename == '.DS_Store':
                            continue
                        
                        # Only process PDF files
                        if member.lower().endswith('.pdf'):
                            # Only add if we haven't seen this filename before
                            if filename not in seen_filenames:
                                pdf_members.append(member)
                                seen_filenames.add(filename)
                    
                    if not pdf_members:
                        st.warning(f"⚠️ {uf.name}: No PDF files found in ZIP")
                    else:
                        extracted_count = 0
                        for member in pdf_members:
                            try:
                                # Check if this is actually a file (not a directory)
                                member_info = zf.getinfo(member)
                                if member_info.is_dir():
                                    continue
                                
                                pdf_bytes = zf.read(member)
                                if len(pdf_bytes) == 0:
                                    continue
                                
                                # Mimic Streamlit UploadedFile
                                class MockUploadedFile:
                                    def __init__(self, name, data_bytes):
                                        self.name = name
                                        self._data = data_bytes
                                    def getbuffer(self):
                                        return io.BytesIO(self._data)
                                
                                filename = Path(member).name
                                expanded_files.append(
                                    MockUploadedFile(filename, pdf_bytes)
                                )
                                extracted_count += 1
                            except Exception as e:
                                st.error(f"❌ Error extracting {member} from {uf.name}: {str(e)}")
                        
                        if extracted_count > 0:
                            st.success(f"✅ {uf.name}: Extracted {extracted_count} PDF file(s)")
                        else:
                            st.warning(f"⚠️ {uf.name}: Could not extract any PDF files")
                    
                    zf.close()
                    
                except zipfile.BadZipFile:
                    st.error(f"❌ {uf.name}: Invalid or corrupted ZIP file")
                except Exception as e:
                    st.error(f"❌ Error processing ZIP file {uf.name}: {str(e)}")
                    if verbose_mode:
                        st.exception(e)
            elif name_lower.endswith('.pdf'):
                expanded_files.append(uf)
    
    # Store in session state
    if expanded_files:
        st.session_state.uploaded_files = expanded_files
        st.success(f"✅ Ready to process {len(expanded_files)} file(s)")
    else:
        st.session_state.uploaded_files = []
        if uploaded_files:
            st.warning("⚠️ No valid PDF files found. Please check your uploads.")
    
    # Folder Upload - Secondary Option (Collapsed by default)
    with st.expander("📁 Alternative: Select from Folder (Advanced)", expanded=False):
        st.info("If you prefer to select files from a folder on your computer instead of drag & drop, use this option.")
        
        # Multiple ways to select folders
        st.subheader("🔍 Choose Folder Selection Method")
        
        folder_method = st.radio(
            "How would you like to select your folder?",
            ["📂 Quick Access (Recommended)", "📁 Browse Common Locations", "⌨️ Manual Path Entry"],
            horizontal=True
        )
        
        folder_path = None
        
        if folder_method == "📂 Quick Access (Recommended)":
            st.markdown("**Quick access to common PDF folders:**")
            
            # Quick access buttons for common locations
            col1, col2, col3 = st.columns(3)
            
            with col1:
                if st.button("📁 Current Project Folder", use_container_width=True):
                    folder_path = Path("SeptemberMR2025complete")
            
            with col2:
                if st.button("📁 Files Directory", use_container_width=True):
                    folder_path = Path("files")
            
            with col3:
                if st.button("📁 Desktop", use_container_width=True):
                    folder_path = Path.home() / "Desktop"
            
            # Show current working directory
            st.info(f"💡 **Current directory:** `{Path.cwd()}`")
            
            # Allow user to enter a relative path from current directory
            relative_path = st.text_input(
                "Or enter a relative path from current directory:",
                placeholder="e.g., 'SeptemberMR2025complete' or 'files'",
                help="Enter a folder name or path relative to the current directory"
            )
            
            if relative_path:
                folder_path = Path(relative_path)
        
        elif folder_method == "📁 Browse Common Locations":
            st.markdown("**Browse common system locations:**")
            
            # Common system paths
            common_paths = {
                "🏠 Home Directory": str(Path.home()),
                "📁 Desktop": str(Path.home() / "Desktop"),
                "📁 Documents": str(Path.home() / "Documents"),
                "📁 Downloads": str(Path.home() / "Downloads"),
                "📁 Current Project": str(Path.cwd()),
                "📁 Project Files": str(Path.cwd() / "files"),
                "📁 September Reports": str(Path.cwd() / "SeptemberMR2025complete")
            }
            
            selected_path = st.selectbox(
                "Choose a common location:",
                list(common_paths.keys()),
                help="Select from common system locations"
            )
            
            if selected_path:
                base_path = Path(common_paths[selected_path])
                
                if base_path.exists():
                    # Show subdirectories if any
                    try:
                        subdirs = [d for d in base_path.iterdir() if d.is_dir()]
                        if subdirs:
                            st.write("**Available subdirectories:**")
                            subdir_names = [d.name for d in subdirs]
                            selected_subdir = st.selectbox(
                                "Choose subdirectory (optional):",
                                ["None"] + subdir_names
                            )
                            
                            if selected_subdir != "None":
                                folder_path = base_path / selected_subdir
                            else:
                                folder_path = base_path
                        else:
                            folder_path = base_path
                    except PermissionError:
                        st.error("❌ Permission denied to access this location")
                        folder_path = None
                else:
                    st.error(f"❌ Path does not exist: {base_path}")
                    folder_path = None
        else:  # Manual Path Entry
            st.markdown("**Manual path entry:**")
            folder_path = st.text_input(
                "Enter full folder path:",
                placeholder="/Users/username/Desktop/MyPDFs",
                help="Enter the complete path to your PDF folder"
            )
            
            if folder_path:
                folder_path = Path(folder_path)
        
        # Process the selected folder
        folder_uploaded_files = []  # Initialize empty list for folder upload
        
        if folder_path:
            folder_path = Path(folder_path)
            if folder_path.exists() and folder_path.is_dir():
                # Find all PDF files in the folder
                pdf_files = list(folder_path.glob("*.pdf"))
                if pdf_files:
                    st.success(f"✅ Found {len(pdf_files)} PDF files in: `{folder_path}`")
                    
                    # Show file list preview
                    with st.expander(f"📋 Preview {len(pdf_files)} files found"):
                        for i, pdf_file in enumerate(pdf_files[:10]):  # Show first 10
                            st.write(f"{i+1}. {pdf_file.name}")
                        if len(pdf_files) > 10:
                            st.write(f"... and {len(pdf_files) - 10} more files")
                    
                    # Create file-like objects for each PDF
                    for pdf_file in pdf_files:
                        try:
                            with open(pdf_file, 'rb') as f:
                                file_data = f.read()
                            
                            # Create a file-like object that mimics uploaded file
                            class MockUploadedFile:
                                def __init__(self, name, data):
                                    self.name = name
                                    self._data = data
                                
                                def getbuffer(self):
                                    return io.BytesIO(self._data)
                            
                            folder_uploaded_files.append(MockUploadedFile(pdf_file.name, file_data))
                            
                        except Exception as e:
                            st.error(f"Error reading {pdf_file.name}: {e}")
                    
                    # Store in session state (only if folder upload was used)
                    if folder_uploaded_files:
                        st.session_state.uploaded_files = folder_uploaded_files
                        st.success(f"✅ Ready to process {len(folder_uploaded_files)} file(s) from folder")
                else:
                    st.warning(f"⚠️ No PDF files found in: `{folder_path}`")
                    st.info("💡 Make sure the folder contains PDF files with `.pdf` extension")
            else:
                st.error(f"❌ Invalid path or folder does not exist: `{folder_path}`")
    
    # Main content area
    col1, col2 = st.columns([2, 1])
    
    with col1:
        st.header("📋 Processing Status")
        
        # Get uploaded files from session state
        uploaded_files = st.session_state.get('uploaded_files', [])
        
        if uploaded_files:
            st.success(f"✅ {len(uploaded_files)} file(s) ready for processing")
            
            # Show file names
            with st.expander(f"📋 View {len(uploaded_files)} files to process"):
                for i, file in enumerate(uploaded_files, 1):
                    st.write(f"{i}. {file.name}")
            
            # Process files button
            if st.button("🚀 Process Files", type="primary", use_container_width=True):
                if uploaded_files:
                    with st.spinner("Processing files..."):
                        process_files(uploaded_files, verbose_mode, use_ocr)
                else:
                    st.error("❌ No files to process. Please upload files first.")
        else:
            st.info("👆 Please drag and drop PDF files or a ZIP folder above to get started")
    
    with col2:
        st.header("📊 Quick Stats")
        
        # Display stats if we have processed data
        if 'processed_data' in st.session_state:
            data = st.session_state.processed_data
            st.metric("Schools Processed", len(data))
            st.metric("Total Rows", len(data))
            
            # Show sample data
            if data:
                st.subheader("📋 Sample Data")
                sample_df = pd.DataFrame(data[:5])
                st.dataframe(sample_df, use_container_width=True)
        else:
            st.info("No data processed yet")
    
    # Debug section (can be removed in production)
    if st.sidebar.checkbox("🐛 Show Debug Info"):
        st.sidebar.subheader("Debug Information")
        st.sidebar.write(f"Uploaded files: {len(st.session_state.get('uploaded_files', []))}")
        st.sidebar.write(f"Session state keys: {list(st.session_state.keys())}")
        if 'uploaded_files' in st.session_state:
            st.sidebar.write("File names:")
            for file in st.session_state.uploaded_files:
                st.sidebar.write(f"- {file.name}")
        
        # Test processing with a single file
        if st.sidebar.button("🧪 Test Single File Processing"):
            uploaded_files = st.session_state.get('uploaded_files', [])
            if uploaded_files:
                st.sidebar.write("Testing with first file...")
                try:
                    test_file = uploaded_files[0]
                    st.sidebar.write(f"Testing: {test_file.name}")
                    
                    # Test if we can read the file
                    file_data = bytes(test_file.getbuffer())
                    st.sidebar.write(f"File size: {len(file_data)} bytes")
                    
                    # Test if main.py functions are importable
                    try:
                        from main import school_from_pdf
                        st.sidebar.write("✅ main.py functions imported successfully")
                    except Exception as import_error:
                        st.sidebar.error(f"❌ Import error: {import_error}")
                    
                except Exception as e:
                    st.sidebar.error(f"Test failed: {e}")
                    st.sidebar.exception(e)
            else:
                st.sidebar.error("No files to test")
        
        # Test with existing PDF file
        if st.sidebar.button("🧪 Test with Existing PDF"):
            try:
                test_pdf_path = Path("files") / "BRIAN School Leader Monthly Report 2025-26.pdf"
                if test_pdf_path.exists():
                    st.sidebar.write(f"Testing with: {test_pdf_path}")
                    school = school_from_pdf(test_pdf_path)
                    if school:
                        st.sidebar.write(f"✅ Success: {school.name}")
                    else:
                        st.sidebar.write("❌ Failed to extract school data")
                else:
                    st.sidebar.error(f"Test PDF not found: {test_pdf_path}")
            except Exception as e:
                st.sidebar.error(f"Test failed: {e}")
                st.sidebar.exception(e)

def process_files(uploaded_files, verbose_mode, use_ocr):
    """Process uploaded PDF files and extract school data."""
    
    # Add error handling for empty file list
    if not uploaded_files:
        st.error("❌ No files to process!")
        return
    
    # Display loading GIF
    gif_container = st.empty()
    
    # Use a direct GIF URL (must end in .gif)
    loading_gif_url_list = ["https://media1.tenor.com/m/pR1FSngr_3sAAAAd/were-here-john-david-washington.gif", "https://giffiles.alphacoders.com/221/221681.gif", "https://media0.giphy.com/media/v1.Y2lkPTc5MGI3NjExNDcxY2gwMnY3d29wOWM0dm1pNDVpZ2ZkY2NyM3FyNm4zcDFkbW41ZSZlcD12MV9pbnRlcm5hbF9naWZfYnlfaWQmY3Q9Zw/mCqOIwWNnqwdJUn4FW/giphy.gif", "https://media1.tenor.com/m/6tVFJnLL3q0AAAAd/matrix-fight.gif", "https://media1.tenor.com/m/FDkNs-UIqxkAAAAC/bane-batman.gif", "https://media1.giphy.com/media/v1.Y2lkPTc5MGI3NjExa3M5eXU3M2JpdWtpZGc3eGM3aW9hbTd3aGc1cGF1NGxoejZuY3piMiZlcD12MV9pbnRlcm5hbF9naWZfYnlfaWQmY3Q9Zw/RLnTaVfBJ7Sxv5mIVg/giphy.gif"]
    
    # Or use a local file (place loading.gif in project root)
    loading_gif_path = Path("loading.gif")

    loading_gif_url = random.choice(loading_gif_url_list)
    
    try:
        if loading_gif_path.exists():
            gif_container.image(str(loading_gif_path), width=600)
        elif loading_gif_url:
            gif_container.image(loading_gif_url, width=600)
        else:
            gif_container.markdown("🔄 **Processing...**")
    except Exception as e:
        gif_container.markdown("🔄 **Processing...**")
        if verbose_mode:
            st.warning(f"Could not load GIF: {e}")
    
    progress_bar = st.progress(0)
    status_text = st.empty()
    
    all_rows = []
    successful_files = 0
    failed_files = []
    
    # Add initial status
    st.info(f"🚀 Starting to process {len(uploaded_files)} file(s)...")
    
    try:
        # Create temporary directory for processing
        with tempfile.TemporaryDirectory() as temp_dir:
            temp_path = Path(temp_dir)
            
            for i, uploaded_file in enumerate(uploaded_files):
                # Update progress
                progress = (i + 1) / len(uploaded_files)
                progress_bar.progress(progress)
                status_text.text(f"Processing {uploaded_file.name}...")
                
                try:
                    # Save uploaded file to temporary location
                    temp_pdf_path = temp_path / uploaded_file.name
                    with open(temp_pdf_path, "wb") as f:
                        # Handle both real UploadedFile (memoryview) and MockUploadedFile (BytesIO)
                        buffer = uploaded_file.getbuffer()
                        if isinstance(buffer, io.BytesIO):
                            # MockUploadedFile returns BytesIO
                            f.write(buffer.getvalue())
                        else:
                            # Real UploadedFile returns memoryview
                            f.write(bytes(buffer))
                    
                    # Check if file was saved correctly
                    if not temp_pdf_path.exists():
                        raise Exception(f"Failed to save file: {uploaded_file.name}")
                    
                    if verbose_mode:
                        st.write(f"📄 Processing: {uploaded_file.name}")
                    
                    # Process the PDF
                    school = school_from_pdf(temp_pdf_path)
                    
                    if school:
                        rows = build_rows_for_school(school)
                        all_rows.extend(rows)
                        successful_files += 1
                        
                        if verbose_mode:
                            st.success(f"✅ {school.name}: {len(rows)} row(s) extracted")
                        else:
                            st.success(f"✅ {uploaded_file.name}: {len(rows)} row(s)")
                    else:
                        failed_files.append(uploaded_file.name)
                        st.error(f"❌ Failed to extract data from {uploaded_file.name}")
                        
                except Exception as e:
                    failed_files.append(f"{uploaded_file.name} ({str(e)})")
                    st.error(f"❌ Error processing {uploaded_file.name}: {e}")
                    if verbose_mode:
                        st.exception(e)  # Show full traceback in verbose mode
                    else:
                        st.info("💡 Enable 'Verbose Mode' in sidebar to see detailed error information")
    
    except Exception as e:
        gif_container.empty()
        st.error(f"❌ Critical error during processing: {e}")
        if verbose_mode:
            st.exception(e)
        return
    
    # Update progress to complete
    progress_bar.progress(1.0)
    status_text.text("Processing complete!")
    
    # Clear loading GIF
    gif_container.empty()
    
    # Store results in session state
    st.session_state.processed_data = all_rows
    st.session_state.successful_files = successful_files
    st.session_state.failed_files = failed_files
    
    # Display results
    display_results(all_rows, successful_files, failed_files, len(uploaded_files))

def display_results(all_rows, successful_files, failed_files, total_files):
    """Display processing results and provide download options."""
    
    st.markdown("---")
    st.header("📈 Processing Results")
    
    # Results summary
    col1, col2, col3, col4 = st.columns(4)
    
    with col1:
        st.metric("Total Files", total_files)
    with col2:
        st.metric("Successful", successful_files)
    with col3:
        st.metric("Failed", len(failed_files))
    with col4:
        st.metric("Total Rows", len(all_rows))
    
    # Display data table
    if all_rows:
        st.subheader("📋 Extracted Data")
        
        # Create DataFrame
        df = pd.DataFrame(all_rows)
        df = df.sort_values('School')
        
        # Display with search and filtering
        search_term = st.text_input("🔍 Search schools:", placeholder="Enter school name...")
        if search_term:
            df = df[df['School'].str.contains(search_term, case=False, na=False)]
        
        # Show the data
        st.dataframe(df, use_container_width=True)
        
        # Download options
        st.subheader("💾 Download Options")
        
        col1, col2 = st.columns(2)
        
        with col1:
            # Excel download
            excel_buffer = io.BytesIO()
            with pd.ExcelWriter(excel_buffer, engine='openpyxl') as writer:
                df.to_excel(writer, sheet_name='School Data', index=False)
            excel_buffer.seek(0)
            
            st.download_button(
                label="📊 Download Excel File",
                data=excel_buffer.getvalue(),
                file_name=f"monthly_report_data_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
        
        with col2:
            # CSV download
            csv_buffer = io.StringIO()
            df.to_csv(csv_buffer, index=False)
            csv_buffer.seek(0)
            
            st.download_button(
                label="📄 Download CSV File",
                data=csv_buffer.getvalue(),
                file_name=f"monthly_report_data_{datetime.now().strftime('%Y%m%d_%H%M%S')}.csv",
                mime="text/csv"
            )
        
        # Data visualization
        st.subheader("📊 Data Visualization")
        
        # Create visualizations
        viz_col1, viz_col2 = st.columns(2)
        
        with viz_col1:
            # Students by school
            if 'Students' in df.columns:
                students_data = df[df['Students'].notna() & (df['Students'] > 0)]
                if not students_data.empty:
                    st.bar_chart(students_data.set_index('School')['Students'])
                    st.caption("Students by School")
        
        with viz_col2:
            # Teachers by school
            if 'Teachers' in df.columns:
                teachers_data = df[df['Teachers'].notna() & (df['Teachers'] > 0)]
                if not teachers_data.empty:
                    st.bar_chart(teachers_data.set_index('School')['Teachers'])
                    st.caption("Teachers by School")
    
    # Show failed files if any
    if failed_files:
        st.subheader("❌ Failed Files")
        for failed_file in failed_files:
            st.error(f"• {failed_file}")

def show_help():
    """Display help information."""
    st.header("❓ Help & Instructions")
    
    st.markdown("""
    ### How to Use This Application
    
    1. **Upload Files**: Use the sidebar to upload one or more PDF files containing school monthly reports
    2. **Configure Options**: 
       - Enable "Verbose Mode" for detailed processing information
       - Enable "Use OCR Processing" for better text extraction
    3. **Process Files**: Click the "Process Files" button to extract data
    4. **View Results**: Review the extracted data in the table
    5. **Download Data**: Use the download buttons to save results as Excel or CSV
    
    ### Supported Data Fields
    
    The application extracts the following data from each PDF:
    - **School Name**: Automatically detected
    - **Students**: Number of students (K-8 and 9-12)
    - **Teachers**: Number of teachers (K-8 and 9-12)
    - **Sub**: Substitute teachers
    - **OSS**: Out-of-school suspensions
    - **EX**: Expulsions
    - **ER**: Emergency removals
    - **MDM**: Manifestation determination meetings
    
    ### Tips for Best Results
    
    - Ensure PDF files are clear and readable
    - Use high-quality scans for better OCR results
    - Check the verbose mode for detailed processing information
    - Review extracted data before downloading
    """)

if __name__ == "__main__":
    # Add navigation
    st.sidebar.markdown("---")
    
    # Navigation
    page = st.sidebar.selectbox("Navigate", ["🏠 Main", "❓ Help"])
    
    if page == "🏠 Main":
        main()
    elif page == "❓ Help":
        show_help()
