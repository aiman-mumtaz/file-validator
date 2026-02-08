import io
import logging
from datetime import datetime
import streamlit as st
from report_generator import ValidationError, generate_excel_report

# loggin framework setup
@st.cache_resource
def setup_logging():
    logger = logging.getLogger("validation_app")
    logger.setLevel(logging.INFO)
    
    if not logger.handlers:
        formatter = logging.Formatter('%(asctime)s - %(name)s - %(levelname)s - %(message)s')
        
        console_handler = logging.StreamHandler()
        console_handler.setFormatter(formatter)
        logger.addHandler(console_handler)
        
        file_handler = logging.FileHandler("app_activity.log")
        file_handler.setFormatter(formatter)
        logger.addHandler(file_handler)
        
    return logger

logger = setup_logging()
# ui config
st.set_page_config(page_title="Validation Report", layout="wide")

st.title("Claims Validation Report")
st.caption("Upload Base and Validation files to generate the validation report Excel.")

with st.sidebar:
    st.header("Inputs")
    mapping_default_path = "/sample_data/mapping.csv"
    mapping_path = st.text_input("Mapping CSV path (optional)", value=mapping_default_path)
    mapping_upload = st.file_uploader("Or upload mapping CSV", type=["csv"])

base_upload = st.file_uploader("Upload Base file (Reference)", type=["txt", "dat", "base"])
validation_upload = st.file_uploader("Upload Validation file", type=["txt", "dat", "validation"])

col1, col2 = st.columns(2)

with col1:
    generate = st.button("Generate Report", type="primary")

if generate:
    logger.info("Starting report generation process...")
    
    if not base_upload or not validation_upload:
        logger.warning("Generation failed: Missing Base or Validation file upload.")
        st.error("Please upload both Base and Validation files.")
        st.stop()

    try:
        if mapping_upload is not None:
            logger.info(f"Using uploaded mapping file: {mapping_upload.name}")
            mapping_stream = io.TextIOWrapper(io.BytesIO(mapping_upload.getvalue()), encoding="utf-8", errors="ignore")
        else:
            logger.info(f"Attempting to open default mapping at: {mapping_path}")
            mapping_stream = open(mapping_path, "r", encoding="utf-8")
    except FileNotFoundError:
        logger.error(f"Mapping CSV not found at path: {mapping_path}")
        st.error("Mapping CSV not found. Upload the mapping file or update the path.")
        st.stop()

    try:
        # Log metadata about the files being processed
        logger.info(f"Processing Base File: {base_upload.name} ({base_upload.size} bytes)")
        logger.info(f"Processing Validation File: {validation_upload.name} ({validation_upload.size} bytes)")

        with mapping_stream:
            base_stream = io.TextIOWrapper(io.BytesIO(base_upload.getvalue()), encoding="utf-8", errors="ignore")
            validation_stream = io.TextIOWrapper(io.BytesIO(validation_upload.getvalue()), encoding="utf-8", errors="ignore")

            with base_stream, validation_stream:
                excel_bytes, summary = generate_excel_report(mapping_stream, base_stream, validation_stream)
        
        timestamp = summary['timestamp']
        filename = f"Validation_Report_{timestamp}.xlsx"

        logger.info(f"Report successfully generated. Filename: {filename}")
        st.success("Report generated successfully.")

        st.download_button(
            label="Download Validation Report",
            data=excel_bytes,
            file_name=filename,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )

        st.subheader("Summary")
        st.write(
            {
                "Total Base claims": summary['total_base'],
                "Total Validation claims": summary['total_validation'],
                "Total differences": summary['total_differences'],
                "Fields with differences": summary['fields_with_differences'],
                "Missing in Validation": summary['missing_in_validation'],
                "Extra in Validation": summary['extra_in_validation'],
            }
        )
        logger.info(f"Summary displayed: {summary['total_differences']} differences found.")

    except ValidationError as exc:
        logger.warning(f"Business Logic Validation Error: {str(exc)}")
        st.error(str(exc))
    except Exception as exc:
        # stack trace
        logger.exception("A critical unexpected error occurred during report generation.")
        st.error(f"Unexpected error: {exc}")
