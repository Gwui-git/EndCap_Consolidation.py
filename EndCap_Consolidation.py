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

# Main processing function
def process_files(endcaps_file, open_space_file, selected_types):
    # Read Excel files
    try:
        endcaps_df = pd.read_excel(endcaps_file, sheet_name="Sheet1")
        open_space_df = pd.read_excel(open_space_file, sheet_name="Sheet1")
    except Exception as e:
        st.error(f"Failed to read Excel files: {e}")
        return None

    # Initialize assignments list
    assignments = []

    # Step 1: Filter out VIR locations
    open_space_df = open_space_df[open_space_df["Storage Type"] != "VIR"].copy()

    # Step 2: Filter Endcaps by selected types
    endcaps_df = endcaps_df[endcaps_df["Storage Type"].isin(selected_types)].copy()

    # Standardize formats
    for col in ["Material", "Batch", "Storage Unit", "Storage Bin"]:
        if col in endcaps_df.columns:
            endcaps_df[col] = endcaps_df[col].astype(str).str.strip()
    
    for col in ["Material Number", "Batch Number", "Storage Bin"]:
        if col in open_space_df.columns:
            open_space_df[col] = open_space_df[col].astype(str).str.strip()

    # Parse batches
    endcaps_df[["Batch Prefix", "Batch Date"]] = endcaps_df["Batch"].apply(parse_batch).apply(pd.Series)
    open_space_df[["Batch Prefix", "Batch Date"]] = open_space_df["Batch Number"].apply(parse_batch).apply(pd.Series)
    
    # Convert to datetime
    endcaps_df["Batch Date"] = pd.to_datetime(endcaps_df["Batch Date"], errors='coerce')
    open_space_df["Batch Date"] = pd.to_datetime(open_space_df["Batch Date"], errors='coerce')

    # Filter available bins
    available_bins = open_space_df[
        ~open_space_df["Storage Type"].isin(selected_types) & 
        (open_space_df["Utilization %"] < 100)
    ].copy()
    available_bins.sort_values("Avail SU", ascending=False, inplace=True)

    # Match materials and batches
    for su_id, su_group in endcaps_df.groupby("Storage Unit"):
        material = su_group["Material"].iloc[0]
        batch_prefix = su_group["Batch Prefix"].iloc[0]
        source_bin = su_group["Storage Bin"].iloc[0]

        # Find matching bins
        matching_bins = available_bins[
            (available_bins["Material Number"] == material) & 
            (available_bins["Batch Prefix"] == batch_prefix)
        ].copy().dropna(subset=["Batch Date"])

        for _, target_bin in matching_bins.iterrows():
            # Check date compatibility
            valid_match = True
            for _, su_row in su_group.iterrows():
                su_date = su_row["Batch Date"]
                if pd.isna(su_date):
                    valid_match = False
                    break
                    
                date_diff = abs((target_bin["Batch Date"] - su_date).days)
                if date_diff > 364:
                    valid_match = False
                    break

            if valid_match and target_bin["Avail SU"] >= 1:
                new_avail = target_bin["Avail SU"] - 1
                
                assignments.append([
                    target_bin["Storage Type"], target_bin["Storage Bin"], source_bin,
                    su_group["Storage Type"].iloc[0], material, 
                    target_bin["Batch Number"], su_group["Batch"].iloc[0],
                    target_bin["SU Capacity"], 1, new_avail,
                    su_id, su_group["Total Stock"].iloc[0]
                ])
                
                available_bins.loc[available_bins["Storage Bin"] == target_bin["Storage Bin"], "Avail SU"] = new_avail
                break

    # Prepare output
    if assignments:
        return pd.DataFrame(assignments, columns=[
            "Target Storage Type", "Target Bin", "Source Bin",
            "Source Storage Type", "Material", "Target Batch",
            "Source Batch", "Bin Capacity", "SU Count", 
            "Remaining Capacity", "Storage Unit", "Total Stock"
        ])
    return None

# Streamlit UI
st.set_page_config(layout="wide", page_title="Inventory Consolidation Tool")
st.title("📦 Advanced Inventory Processor")

# File upload with validation
with st.expander("📂 STEP 1: Upload Files", expanded=True):
    col1, col2 = st.columns(2)
    with col1:
        endcaps_file = st.file_uploader("Endcaps File", type=None)
        endcaps_file = validate_excel_file(endcaps_file)
    with col2:
        open_space_file = st.file_uploader("Open Space File", type=None)
        open_space_file = validate_excel_file(open_space_file)

# Storage type selection
with st.expander("⚙️ STEP 2: Configure Filters", expanded=True):
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
            if st.checkbox(stype, key=f"stype_{i}"):
                selected_types.append(stype)

# Process files
if st.button("🚀 Process Files", type="primary"):
    if not endcaps_file or not open_space_file:
        st.error("Please upload both files!")
    elif not selected_types:
        st.error("Please select at least one storage type!")
    else:
        with st.spinner("Processing files..."):
            result = process_files(endcaps_file, open_space_file, selected_types)
            
            if result is not None:
                st.success(f"✅ Found {len(result)} assignments!")
                
                # Show results
                with st.expander("🔍 View Assignments", expanded=True):
                    st.dataframe(result)
                
                # Download
                output_file = "inventory_assignments.xlsx"
                if os.path.exists(output_file) and is_file_open(output_file):
                    st.error(f"Please close {output_file} before downloading")
                else:
                    result.to_excel(output_file, index=False)
                    with open(output_file, "rb") as f:
                        st.download_button(
                            "📥 Download Results",
                            f,
                            file_name=output_file,
                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                        )
            else:
                st.warning("⚠️ No valid assignments found")
