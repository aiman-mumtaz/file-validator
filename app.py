import io
from datetime import datetime

import streamlit as st

from report_generator import ValidationError, generate_excel_report

st.set_page_config(page_title="Validation Report", layout="wide")

st.title("Claims Validation Report")
st.caption("Upload Base and Validation files to generate the validation report Excel.")

with st.sidebar:
    st.header("Inputs")
    mapping_default_path = "/Path/To/Mapping.csv"
    mapping_path = st.text_input("Mapping CSV path (optional)", value=mapping_default_path)
    mapping_upload = st.file_uploader("Or upload mapping CSV", type=["csv"])

base_upload = st.file_uploader("Upload Base file (Reference)", type=["txt", "dat", "base"])
validation_upload = st.file_uploader("Upload Validation file", type=["txt", "dat", "validation"])

col1, col2 = st.columns(2)

with col1:
    generate = st.button("Generate Report", type="primary")

with col2:
    st.write("")

if generate:
    if not base_upload or not validation_upload:
        st.error("Please upload both Base and Validation files.")
        st.stop()

    try:
        if mapping_upload is not None:
            mapping_stream = io.TextIOWrapper(io.BytesIO(mapping_upload.getvalue()), encoding="utf-8", errors="ignore")
        else:
            mapping_stream = open(mapping_path, "r", encoding="utf-8")
    except FileNotFoundError:
        st.error("Mapping CSV not found. Upload the mapping file or update the path.")
        st.stop()

    try:
        with mapping_stream:
            base_stream = io.TextIOWrapper(io.BytesIO(base_upload.getvalue()), encoding="utf-8", errors="ignore")
            validation_stream = io.TextIOWrapper(io.BytesIO(validation_upload.getvalue()), encoding="utf-8", errors="ignore")

            with base_stream, validation_stream:
                excel_bytes, summary = generate_excel_report(mapping_stream, base_stream, validation_stream)
        timestamp = summary['timestamp']
        filename = f"Validation_Report_{timestamp}.xlsx"

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
    except ValidationError as exc:
        st.error(str(exc))
    except Exception as exc:
        st.error(f"Unexpected error: {exc}")

