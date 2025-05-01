import pandas as pd
import streamlit as st
from datetime import datetime
import os

# Function to parse batch
def parse_batch(batch):
    if isinstance(batch, str) and len(batch) >= 10:
        batch_prefix = batch[:2]
        try:
            year = int("20" + batch[-2:])
            week = int(batch[-4:-2])
            batch_date = datetime.strptime(f"{year}-W{week}-1", "%Y-W%W-%w")
            return batch_prefix, batch_date
        except ValueError:
            return batch_prefix, pd.NaT
    return None, pd.NaT

# Function to check if a file is open
def is_file_open(file_path):
    try:
        with open(file_path, "a"):
            return False
    except PermissionError:
        return True
    except FileNotFoundError:
        return False

# Function to validate Excel file extension
def validate_excel_extension(filename):
    return filename.lower().endswith('.xlsx')

# Main function to process files (UNCHANGED)
def process_files(endcaps_file, open_space_file, selected_types):
    # [Rest of your existing process_files function remains exactly the same]
    # ... [all original code] ...

# Streamlit UI with modified file uploaders
st.title("Storage Type Filter")

# File uploaders with case-insensitive validation
st.header("Upload Files")
endcaps_file = st.file_uploader("Endcaps File", type=None)  # Changed to type=None
if endcaps_file and not validate_excel_extension(endcaps_file.name):
    st.error("Endcaps file must be an .xlsx file")
    endcaps_file = None

open_space_file = st.file_uploader("Open Space File", type=None)  # Changed to type=None
if open_space_file and not validate_excel_extension(open_space_file.name):
    st.error("Open Space file must be an .xlsx file")
    open_space_file = None

# [Rest of your existing UI code remains exactly the same]
# Storage type selection
st.header("Select Storage Types to Filter")
storage_types = [
    "901", "902", "910", "916", "920", "921", "922", "980", "998", "999",
    "DT1", "LT1", "LT2", "LT4", "LT5", "OE1", "OVG", "OVL", "OVP", "OVT",
    "PC1", "PC2", "PC4", "PC5", "PSA", "QAH", "RET", "TB1", "TB2", "TB4",
    "TR1", "TR2", "VIR", "VTL"
]

cols = st.columns(4)
selected_types = []

for i, stype in enumerate(storage_types):
    with cols[i % 4]:
        if st.checkbox(stype, key=stype):
            selected_types.append(stype)

# Process button (UNCHANGED)
if st.button("Run Script"):
    if not endcaps_file or not open_space_file:
        st.error("Please upload both input files!")
    elif not selected_types:
        st.error("Please select at least one storage type to filter.")
    else:
        with st.spinner("Processing files..."):
            result = process_files(endcaps_file, open_space_file, selected_types)
            
            if result is not None:
                st.success("Processing complete!")
                st.dataframe(result)
                
                # Download button
                output_file = "final_assignments.xlsx"
                if os.path.exists(output_file) and is_file_open(output_file):
                    st.error(f"The file '{output_file}' is open in another program. Please close it and try again.")
                else:
                    result.to_excel(output_file, index=False, engine="openpyxl")
                    with open(output_file, "rb") as file:
                        st.download_button(
                            label="Download Results",
                            data=file,
                            file_name=output_file,
                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                        )
            else:
                st.warning("No suitable matches were found.")
