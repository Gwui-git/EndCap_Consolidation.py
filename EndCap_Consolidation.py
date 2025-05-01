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

# Function to validate Excel file
def validate_excel_file(uploaded_file):
    if uploaded_file is None:
        return None
    if not uploaded_file.name.lower().endswith('.xlsx'):
        st.error(f"Invalid file type: {uploaded_file.name}. Please upload an .xlsx file")
        return None
    return uploaded_file

# Main function to process files
def process_files(endcaps_file, open_space_file, selected_types):
    try:
        endcaps_df = pd.read_excel(endcaps_file, sheet_name="Sheet1")
        open_space_df = pd.read_excel(open_space_file, sheet_name="Sheet1")
    except Exception as e:
        st.error(f"Failed to read Excel files: {e}")
        return None

    # Rest of your processing logic remains exactly the same...
    open_space_df = open_space_df[open_space_df["Storage Type"] != "VIR"].copy()
    endcaps_df = endcaps_df[endcaps_df["Storage Type"].isin(selected_types)].copy()
    
    # ... [keep all your existing processing code]
    
    if assignments:
        final_output = pd.DataFrame(assignments, columns=[
            "Open Space Storage Type", "Storage Bin", "Bin Moving From", "Endcap Storage Type",
            "Material", "Batch", "Original Batch", "SU Capacity", "SU Count", "Avail SU", "Storage Unit", "Total Stock"
        ])
        return final_output
    else:
        return None

# Streamlit UI
st.title("Storage Type Filter")

# File uploaders with validation
st.header("Upload Files")
endcaps_file = st.file_uploader("Endcaps File", type=None)
endcaps_file = validate_excel_file(endcaps_file)

open_space_file = st.file_uploader("Open Space File", type=None)
open_space_file = validate_excel_file(open_space_file)

# Storage type selection (unchanged)
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

# Process button
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
