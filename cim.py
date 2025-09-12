# -*- coding: utf-8 -*-
"""
Streamlit Web Application to generate a .cim file for QAD's icunis.p program
from a specific Excel template.
"""

import streamlit as st
import pandas as pd
import io
from datetime import datetime

def generate_cim_content(df):
    """
    Generates the multi-line .cim file content from a DataFrame, matching the
    specific format required by the 'icunis.p' program.

    Args:
        df (pd.DataFrame): The DataFrame containing the source data.

    Returns:
        str: A string containing the full content for the .cim file.
    """
    output = io.StringIO()

    # --- CIM FILE STRUCTURE DEFINITION ---
    # This section is now fine-tuned to match the exact BCF.cim format.

    # Iterate through each row of the DataFrame
    for index, row in df.iterrows():
        try:
            # --- 1. Data Extraction and Cleanup ---
            pt_part = str(row.get('pt part', '')).strip()
            qty = str(row.get('lotserial qty', '0')).strip()
            site = str(row.get('site', '')).strip()
            location = str(row.get('location', '')).strip()
            lot_ref = str(row.get('lotref', '')).strip()
            order_nbr = str(row.get('ordernbr', '')).strip()

            # Handle date formatting
            eff_date_val = row.get('eff date')
            if pd.notna(eff_date_val):
                # Attempt to parse the date, robustly handling different formats
                if isinstance(eff_date_val, datetime):
                     eff_date = eff_date_val.strftime('%-d/%-m/%y') # Use - to avoid leading zeros
                else:
                    eff_date = pd.to_datetime(eff_date_val).strftime('%-d/%-m/%y')
            else:
                eff_date = ""

            # The template has two columns named 'dr acct'. Pandas renames the second to 'dr acct.1'
            dr_acct1 = str(row.get('dr acct', '0')).strip()
            dr_acct2 = str(row.get('dr acct.1', '0')).strip()

            # Skip rows where the part number is empty
            if not pt_part:
                continue

            # --- 2. CIM Record Construction (Corrected Format) ---
            output.write("@@batchload  icunis.p\n")
            output.write(f'"{pt_part}"\n')
            # Line 3: Qty - - "Site" "Location" "" ""
            output.write(f'{qty} - - "{site}" "{location}" "" ""\n')
            # Line 4: "LotRef" - - "" "OrderNbr" Date Acct1 Acct2
            output.write(f'"{lot_ref}" - - "" "{order_nbr}" {eff_date} {dr_acct1} {dr_acct2}\n')
            output.write("-\n")
            output.write("-\n")
            output.write("@@end\n")

        except Exception as e:
            # If a row fails, we can note it and continue
            st.warning(f"Skipping row {index + 1} due to an error: {e}. Please check the data in this row.")
            continue

    return output.getvalue()

# --- Streamlit App UI ---
st.set_page_config(layout="wide")
st.title("`.cim` File Generator for QAD `icunis.p`")
st.markdown("Upload your **Unplanned Issue Template** Excel file to generate the corresponding `.cim` file for batch loading.")

uploaded_file = st.file_uploader("Choose an Excel or CSV file", type=["xlsx", "xls", "csv"])

if uploaded_file is not None:
    try:
        # --- File Reading (Corrected Logic) ---
        # The template has several header lines that need to be skipped.
        # We target the actual header row and skip the metadata rows.
        if uploaded_file.name.endswith('.csv'):
            # The header is on the 2nd line (index 1), so we skip the junk rows around it.
            df = pd.read_csv(uploaded_file, skiprows=[0, 2, 3, 4])
        else:
            # For excel, we use the same logic, specifying the engine.
            df = pd.read_excel(uploaded_file, skiprows=[0, 2, 3, 4], engine='openpyxl')

        # Clean up the DataFrame: remove rows where 'pt part' is not present
        df.dropna(subset=['pt part'], inplace=True)
        df = df[df['pt part'].str.strip() != '']

        st.success("File successfully uploaded and parsed.")
        st.write("### Data Preview (First 5 Rows)")
        st.dataframe(df.head())

        # Generate the .cim file content in memory
        cim_data = generate_cim_content(df)

        if cim_data:
            st.write("### Generated .cim File Preview")
            st.text_area("CIM Content", cim_data, height=300)

            # Provide a download button for the generated file
            st.download_button(
                label="Download .cim File",
                data=cim_data,
                file_name="unplanned_issue.cim",
                mime="text/plain",
            )
        else:
            st.warning("No data was processed. Please check your file content and format.")

    except Exception as e:
        st.error(f"An error occurred while processing the file: {e}")
        st.warning("Please ensure you are uploading the correct 'Unplanned Issue Template' file.")

