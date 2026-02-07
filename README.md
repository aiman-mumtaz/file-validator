# File Validator Tool
This project provides a reusable framework to compare fixed-width file outputs between a legacy sytem and a modernised application

## Features
   - ⁠Fixed-width file parsing driven by a mapping document
   - ⁠Field-level comparison using a base file as reference
   - ⁠Excel validation report with:
      1. Summary sheet
      2. Field-specific worksheets highlighting mismatches
   - ⁠Streamlit-based UI for uploading input files and mappings

### Tech Stack
•⁠  ⁠Python
•⁠  ⁠Pandas
•⁠  ⁠Streamlit
•⁠  ⁠OpenPyXL


## Run

1. Install dependencies:
   - `pip install -r requirements.txt`
2. Start the UI:
   - `streamlit run app.py`

## How to use the application

- Upload the Base and Validation files.
- Provide the mapping CSV either via upload or by updating the path in the sidebar.
- Click **Generate Report**, then download the Excel report.

## Use Cases
   - ⁠Data validation during application migration
   - ⁠Regression testing for reporting systems
   - Automated reconciliation of batch outputs


## Disclaimer
This project uses synthetic data and is a generalized implementation inspired by real-world migration scenarios. It does not contain proprietary or sensitive information.
