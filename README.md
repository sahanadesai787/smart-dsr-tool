# Smart DSR Tool 📊

Smart DSR Tool is a web-based application that automates the preparation of Daily Sales Reports (DSR) using Excel files.

## Features

- Upload daily Sales Order Reports.
- Upload Globus Cataloging and Item Master files.
- Upload Channel Master, Vendor Cost and Week Master files.
- Automatically map product, category, style, size and channel details.
- Calculate TTL Vendor Cost, TTL QTY, TTL ESP, TTL MRP and TTL RGM.
- Identify missing mappings and unmatched records.
- Generate and download the completed DSR as an Excel file.

## Technology Stack

- Python
- Streamlit
- Pandas
- OpenPyXL

## How to Run Locally

1. Install Python.
2. Download or clone this repository.
3. Install dependencies:

   `pip install -r requirements.txt`

4. Run the application:

   `streamlit run app.py`

5. Open the local URL displayed in your terminal.

## Usage

1. Upload the required Excel master files.
2. Upload the daily Sales Order Report.
3. Generate the DSR.
4. Review any missing mappings.
5. Download the generated Excel report.

## Important

Verify the generated report and calculations before using it for official reporting. Do not upload confidential company data to a public repository.

## Author

Smart DSR Tool Project
