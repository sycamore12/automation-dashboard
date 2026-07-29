import streamlit as st
import pdfplumber
import openpyxl
import pandas as pd
import io
import os

# Set Streamlit Page Configuration
st.set_page_config(
    page_title="Rekap Automation Dashboard",
    page_icon="📊",
    layout="wide"
)

# --- CUSTOM CSS FOR BPJS KETENAGAKERJAAN THEME ---
custom_css = """
<style>
    /* Tri-color BPJS Banner */
    .bpjs-banner {
        height: 8px;
        width: 100%;
        background: linear-gradient(90deg, #008C44 33.3%, #F6EA00 33.3%, #F6EA00 66.6%, #005C9A 66.6%);
        margin-top: -10px;
        margin-bottom: 25px;
        border-radius: 4px;
    }
    
    /* Custom styling for the primary 'Process' button (BPJS Green) */
    div.stButton > button:first-child {
        background-color: #008C44 !important;
        color: white !important;
        border: none !important;
        font-weight: bold;
    }
    div.stButton > button:first-child:hover {
        background-color: #007036 !important;
        box-shadow: 0 4px 8px rgba(0,0,0,0.1);
    }
    
    /* Custom styling for the Download button (BPJS Blue) */
    div.stDownloadButton > button:first-child {
        background-color: #005C9A !important;
        color: white !important;
        border: none !important;
        font-weight: bold;
    }
    div.stDownloadButton > button:first-child:hover {
        background-color: #004777 !important;
        box-shadow: 0 4px 8px rgba(0,0,0,0.1);
    }
</style>
"""
st.markdown(custom_css, unsafe_allow_html=True)
# --------------------------------------------------

# Supported Indonesian Months
MONTHS = [
    "JANUARI", "FEBRUARI", "MARET", "APRIL", "MEI", "JUNI", 
    "JULI", "AGUSTUS", "SEPTEMBER", "OKTOBER", "NOVEMBER", "DESEMBER"
]

def get_month_from_filename(filename):
    """Detects the month name from the file name."""
    upper_filename = filename.upper()
    for month in MONTHS:
        if month in upper_filename:
            return month
    return None

def extract_data_from_pdf_stream(pdf_file_bytes):
    """Extracts Kode Perisai and NIK BARU values from PDF byte streams."""
    extracted_data = {}
    with pdfplumber.open(pdf_file_bytes) as pdf:
        for page in pdf.pages:
            tables = page.extract_tables()
            for table in tables:
                if not table:
                    continue
                for row in table:
                    if not row or len(row) < 7:
                        continue
                    
                    kode_perisai = str(row[1]).strip() if row[1] else ""
                    
                    if kode_perisai.startswith("AB"):
                        try:
                            val_jht_jkk_jkm = int(row[5]) if row[5] and str(row[5]).strip().isdigit() else 0
                            val_jkk_jkm = int(row[6]) if row[6] and str(row[6]).strip().isdigit() else 0
                            
                            extracted_data[kode_perisai] = {
                                'JHT_JKK_JKM': val_jht_jkk_jkm,
                                'JKK_JKM': val_jkk_jkm
                            }
                        except ValueError:
                            continue
    return extracted_data

def process_excel_update(excel_bytes, all_pdf_data, log_container):
    """Updates the Excel file buffer with all extracted monthly data."""
    wb = openpyxl.load_workbook(excel_bytes)
    sheet = wb.active
    
    total_updates = 0
    
    for month, data in all_pdf_data.items():
        # Find column index for current month
        month_col_start = None
        for col in range(1, sheet.max_column + 1):
            cell_value = str(sheet.cell(row=1, column=col).value).strip().upper()
            if cell_value == month:
                month_col_start = col
                break
                
        if not month_col_start:
            log_container.warning(f"⚠️ Month '{month}' was not found in Row 1 of the Excel template.")
            continue
            
        month_updates = 0
        for row in range(3, sheet.max_row + 1):
            excel_kode = str(sheet.cell(row=row, column=3).value).strip() if sheet.cell(row=row, column=3).value else ""
            
            if excel_kode in data:
                jht_val = data[excel_kode]['JHT_JKK_JKM']
                jkk_val = data[excel_kode]['JKK_JKM']
                
                sheet.cell(row=row, column=month_col_start).value = jht_val
                sheet.cell(row=row, column=month_col_start + 1).value = jkk_val
                
                month_updates += 1
                
        total_updates += month_updates
        log_container.success(f"✅ Updated {month_updates} rows for **{month}**.")
        
    output_stream = io.BytesIO()
    wb.save(output_stream)
    output_stream.seek(0)
    
    return output_stream, total_updates

# --- UI LAYOUT ---

st.title("Sistem Automasi Rekapitulasi Data")
st.markdown('<div class="bpjs-banner"></div>', unsafe_allow_html=True)
st.markdown("Upload your master Excel template and monthly PDF recaps below to process and merge data automatically.")


# Sidebar Setup
st.sidebar.header("📁 Step 1: Upload Files")

excel_file = st.sidebar.file_uploader(
    "Upload HASIL_REKAP.xlsx",
    type=["xlsx"],
    help="Select the master Excel template file."
)

pdf_files = st.sidebar.file_uploader(
    "Upload PDF Recaps (Bulk)",
    type=["pdf"],
    accept_multiple_files=True,
    help="Drag & drop multiple monthly recap PDF files here."
)

st.sidebar.divider()

# Main Body Content
col1, col2 = st.columns([1, 1])

with col1:
    st.subheader("📄 Uploaded PDFs")
    if pdf_files:
        pdf_summary = []
        for pdf in pdf_files:
            detected_month = get_month_from_filename(pdf.name)
            pdf_summary.append({
                "Filename": pdf.name,
                "Detected Month": detected_month if detected_month else "❌ Not Detected"
            })
        st.dataframe(pd.DataFrame(pdf_summary), use_container_width=True)
    else:
        st.info("No PDF files uploaded yet.")

with col2:
    st.subheader("📊 Excel Template Status")
    if excel_file:
        st.success(f"Loaded: `{excel_file.name}`")
    else:
        st.info("No Excel template uploaded yet.")

st.divider()

# Processing Trigger
if st.button("🚀 Process and Generate Updated Excel", type="primary", use_container_width=True):
    if not excel_file:
        st.error("Please upload the `HASIL_REKAP.xlsx` template file first.")
    elif not pdf_files:
        st.error("Please upload at least one recap PDF file.")
    else:
        st.subheader("⚙️ Processing Logs")
        log_box = st.container()
        
        all_pdf_data = {}
        
        # Parse PDFs
        for pdf in pdf_files:
            target_month = get_month_from_filename(pdf.name)
            if not target_month:
                log_box.error(f"Skipping `{pdf.name}`: Could not detect valid month in filename.")
                continue
                
            pdf_bytes = io.BytesIO(pdf.read())
            extracted = extract_data_from_pdf_stream(pdf_bytes)
            
            if extracted:
                all_pdf_data[target_month] = extracted
                log_box.info(f"Extracted **{len(extracted)}** records from `{pdf.name}` for **{target_month}**.")
            else:
                log_box.warning(f"No valid records found in `{pdf.name}`.")
                
        if all_pdf_data:
            # Process Excel
            excel_bytes = io.BytesIO(excel_file.read())
            updated_excel_stream, total_updates = process_excel_update(excel_bytes, all_pdf_data, log_box)
            
            st.balloons()
            st.success(f"🎉 Process completed successfully! Total rows updated across all months: **{total_updates}**")
            
            # Download Button
            st.download_button(
                label="📥 Download Updated HASIL_REKAP.xlsx",
                data=updated_excel_stream,
                file_name="HASIL_REKAP_UPDATED.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True
            )
        else:
            st.error("Could not extract any valid data from the provided PDFs.")
