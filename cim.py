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
    # Use newline='\r\n' to ensure Windows-style line endings (CRLF).
    # The write() calls below will now only use '\n', which this object will
    # automatically convert to the correct '\r\n'.
    output = io.StringIO(newline='\r\n')

    # --- CIM FILE STRUCTURE DEFINITION ---
    # This section is fine-tuned to match the exact BCF.cim format.

    # Iterate through each row of the DataFrame
    for index, row in df.iterrows():
        try:
            # --- 1. Data Extraction and Cleanup ---
            pt_part = str(row.get('pt part', '')).strip()

            # Skip rows where the part number is empty
            if not pt_part or pt_part.lower() == 'nan':
                continue

            # Handle quantity formatting to be integer for whole numbers
            qty_val = pd.to_numeric(row.get('lotserial qty'), errors='coerce')
            if pd.isna(qty_val):
                qty = '0'
            elif qty_val == int(qty_val): # Check if it's a whole number
                qty = str(int(qty_val))
            else:
                qty = str(qty_val)

            site = str(row.get('site', '')).strip()
            location = str(row.get('location', '')).strip()

            # --- CORRECTED MAPPING based on BCF.cim analysis ---
            # The 'lotref' in the CIM file comes from the 'ordernbr' column in the template.
            lot_ref = str(row.get('ordernbr', '')).strip()
            # The 'ordernbr' in the CIM file comes from the 'rmks' column in the template.
            order_nbr = str(row.get('rmks', '')).strip()

            # Handle date formatting
            eff_date_val = row.get('eff date')
            if pd.notna(eff_date_val):
                eff_date = pd.to_datetime(eff_date_val).strftime('%-d/%-m/%y')
            else:
                eff_date = ""

            # The template has two 'dr acct' columns. Pandas renames the second to 'dr acct.1'
            # Convert to numeric, then int, then string to remove any '.0'
            dr_acct1_val = pd.to_numeric(row.get('dr acct'), errors='coerce')
            dr_acct1 = str(int(dr_acct1_val)) if pd.notna(dr_acct1_val) else '0'

            dr_acct2_val = pd.to_numeric(row.get('dr acct.1'), errors='coerce')
            dr_acct2 = str(int(dr_acct2_val)) if pd.notna(dr_acct2_val) else '0'


            # --- 2. CIM Record Construction (Corrected to use '\n' which gets converted to '\r\n') ---
            output.write("@@batchload  icunis.p\n")
            output.write(f'"{pt_part}" \n')
            output.write(f'{qty} - - "{site}" "{location}" "" "" \n')
            output.write(f'"{lot_ref}" - - "" "{order_nbr}" {eff_date} {dr_acct1} {dr_acct2} \n')
            output.write("- \n")
            output.write("- \n")
            output.write("@@end\n")

        except Exception as e:
            # If a row fails, we can note it and continue
            st.warning(f"Skipping row {index + 2} due to an error: {e}. Please check the data in this row.")
            continue

    return output.getvalue()

# --- Streamlit App UI ---
st.set_page_config(layout="wide")
st.title("`.cim` File Generator for QAD `icunis.p`")
st.markdown("Upload your **Unplanned Issue Template** Excel file to generate the corresponding `.cim` file for batch loading.")

uploaded_file = st.file_uploader("Choose an Excel or CSV file", type=["xlsx", "xls", "csv"])

if uploaded_file is not None:
    try:
        # --- File Reading Logic ---
        # The template has header lines that need to be skipped.
        if uploaded_file.name.endswith('.csv'):
            df = pd.read_csv(uploaded_file, skiprows=[0, 2, 3, 4])
        else:
            df = pd.read_excel(uploaded_file, skiprows=[0, 2, 3, 4], engine='openpyxl')

        # Clean up the DataFrame: remove rows where 'pt part' is not present
        df.dropna(subset=['pt part'], inplace=True)
        # Further filter out any rows that might be empty strings after stripping
        df = df[df['pt part'].astype(str).str.strip() != '']

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
                data=cim_data.encode('utf-8'), # Encode for consistency
                file_name="unplanned_issue.cim",
                mime="text/plain",
            )
        else:
            st.warning("No data was processed. Please check your file content and format.")

    except Exception as e:
        st.error(f"An error occurred while processing the file: {e}")
        st.warning("Please ensure you are uploading the correct 'Unplanned Issue Template' file.")

