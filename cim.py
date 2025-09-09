# -*- coding: utf-8 -*-
"""
Streamlit Web Application to generate a .cim file for QAD's icunis.p program
from an Excel file.
"""

import streamlit as st
import pandas as pd
import io
from datetime import datetime

def generate_cim_content(df):
    """
    Generates the multi-line .cim file content from a DataFrame.

    Args:
        df (pd.DataFrame): The DataFrame containing the source data.

    Returns:
        str: A string containing the full content for the .cim file.
    """
    # Using io.StringIO for efficient string building
    output = io.StringIO()

    # --- CIM FILE STRUCTURE DEFINITION ---
    # This section maps the Excel columns to the specific multi-line format
    # required by the 'icunis.p' QAD program.

    # Iterate through each row of the DataFrame from the uploaded Excel file
    for index, row in df.iterrows():
        # --- 1. Data Extraction and Cleanup ---
        # Get data from each column. Use .get() for safety and fill missing values.
        # str() ensures all data is treated as a string. .strip() removes whitespace.
        pt_part = str(row.get('pt part', '')).strip()
        qty = str(row.get('lotserial qty', '0')).strip()
        site = str(row.get('site', '')).strip()
        location = str(row.get('location', '')).strip()
        lot_ref = str(row.get('lotref', '')).strip()
        order_nbr = str(row.get('ordernbr', '')).strip()

        # Handle date formatting
        eff_date_val = row.get('eff date')
        if pd.notna(eff_date_val) and isinstance(eff_date_val, datetime):
            eff_date = eff_date_val.strftime('%d/%m/%y')
        else:
            eff_date = str(eff_date_val or '').strip()

        dr_acct1 = str(row.get('dr acct', '0')).strip()
        dr_acct2 = str(row.get('dr acct.1', '0')).strip() # Assuming the second account column is named 'dr acct.1'

        # --- 2. CIM Record Construction ---
        # Assemble the multi-line string for a single record.
        # The structure is hardcoded to match the 'icunis.p' format exactly.
        output.write("@@batchload  icunis.p\n")
        output.write(f'"{pt_part}"\n')
        # Line 3: Qty, placeholders, Site, placeholders, Location, placeholder
        output.write(f'{qty} - - {site} "" {location} ""\n')
        # Line 4: Lot Ref, placeholders, Order Number, Date, Accounts
        output.write(f'"{lot_ref}" - - "" "{order_nbr}" {eff_date} {dr_acct1} {dr_acct2}\n')
        output.write("-\n")
        output.write("-\n")
        output.write("@@end\n")

    # Get the complete string from the StringIO object
    return output.getvalue()

# --- Streamlit App UI ---
st.set_page_config(layout="wide")
st.title("`.cim` File Generator for QAD `icunis.p`")
st.markdown("Upload your 'Unplanned Issue' Excel file to generate the corresponding `.cim` file for batch loading.")

# File Uploader
uploaded_file = st.file_uploader("Choose an Excel file", type=["xlsx", "xls", "csv"])

if uploaded_file is not None:
    try:
        # Read the uploaded file into a pandas DataFrame
        # For CSV, we can use read_csv. For Excel, use read_excel.
        if uploaded_file.name.endswith('.csv'):
            df = pd.read_csv(uploaded_file)
        else:
            # You might need to specify the sheet name if it's not the first one
            df = pd.read_excel(uploaded_file)

        st.success("File successfully uploaded and read.")
        st.write("### Data Preview")
        st.dataframe(df.head())

        # Generate the .cim file content in memory
        cim_data = generate_cim_content(df)

        st.write("### Generated .cim File Preview")
        st.text_area("CIM Content", cim_data, height=300)

        # Provide a download button for the generated file
        st.download_button(
            label="Download .cim File",
            data=cim_data,
            file_name="unplanned_issue.cim",
            mime="text/plain",
        )

    except Exception as e:
        st.error(f"An error occurred: {e}")
        st.warning("Please ensure the uploaded file has the correct columns: 'pt part', 'lotserial qty', 'site', 'location', etc.")
