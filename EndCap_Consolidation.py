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
        # Try opening the file in append mode
        with open(file_path, "a"):
            return False
    except PermissionError:
        return True  # File is open in another program
    except FileNotFoundError:
        return False  # File does not exist yet

# Main function to process files
def process_files(endcaps_file, open_space_file, selected_types):
    # Read Excel files
    try:
        endcaps_df = pd.read_excel(endcaps_file, sheet_name="Sheet1")
        open_space_df = pd.read_excel(open_space_file, sheet_name="Sheet1")
    except Exception as e:
        st.error(f"Failed to read Excel files: {e}")
        return None

    # Step 1: Filter out VIR locations in Open Spaces
    open_space_df = open_space_df[open_space_df["Storage Type"] != "VIR"].copy()

    # Step 2: Filter Endcaps based on selected storage types
    endcaps_df = endcaps_df[endcaps_df["Storage Type"].isin(selected_types)].copy()

    # Standardize Material & Batch formats
    endcaps_df["Material"] = endcaps_df["Material"].astype(str).str.strip()
    open_space_df["Material Number"] = open_space_df["Material Number"].astype(str).str.strip()
    endcaps_df["Batch"] = endcaps_df["Batch"].astype(str).str.strip()
    open_space_df["Batch Number"] = open_space_df["Batch Number"].astype(str).str.strip()

    # Apply batch parsing
    endcaps_df[["Batch Prefix", "Batch Date"]] = endcaps_df["Batch"].apply(lambda x: pd.Series(parse_batch(x)))
    open_space_df[["Batch Prefix", "Batch Date"]] = open_space_df["Batch Number"].apply(lambda x: pd.Series(parse_batch(x)))

    endcaps_df["Batch Date"] = pd.to_datetime(endcaps_df["Batch Date"], errors='coerce')
    open_space_df["Batch Date"] = pd.to_datetime(open_space_df["Batch Date"], errors='coerce')

    # Step 3: Filter Open Space (Exclude selected storage types)
    available_bins = open_space_df[~open_space_df["Storage Type"].isin(selected_types)].copy()
    available_bins = available_bins[available_bins["Utilization %"] < 100].copy()
    available_bins.sort_values(by="Avail SU", ascending=False, inplace=True)

    # Step 4: Match materials & batches by Storage Unit (SU)
    assignments = []

    for su_id, su_group in endcaps_df.groupby("Storage Unit"):
        material = su_group["Material"].iloc[0]
        batch_prefix = su_group["Batch Prefix"].iloc[0]
        bin_moving_from = su_group["Storage Bin"].iloc[0]

        # Find matching bins
        matching_bins = available_bins[
            (available_bins["Material Number"] == material) & 
            (available_bins["Batch Prefix"] == batch_prefix)
        ].copy()

        matching_bins = matching_bins.dropna(subset=["Batch Date"])

        # Group matching bins by Storage Bin
        for storage_bin, bin_group in matching_bins.groupby("Storage Bin"):
            # Check if all batches in the SU are within the date range of all batches in the bin
            valid_match = True
            for _, su_batch_row in su_group.iterrows():
                su_batch_date = su_batch_row["Batch Date"]
                if pd.isna(su_batch_date):
                    valid_match = False
                    break

                # Check date differences for all batches in the bin
                date_differences = bin_group["Batch Date"].apply(lambda x: abs((x - su_batch_date).days))
                if not all(date_differences <= 364):
                    valid_match = False
                    break

            if valid_match:
                # If all batches in the SU are within range, consider this bin as a match
                avail_su = bin_group["Avail SU"].iloc[0]

                # Ensure there is enough space for the entire SU
                if avail_su >= 1:  # Each SU counts as 1 unit
                    new_avail_su = max(0, avail_su - 1)

                    # Add all batches within the SU to the assignments
                    for _, batch_row in su_group.iterrows():
                        assignments.append([
                            bin_group["Storage Type"].iloc[0], storage_bin, bin_moving_from,
                            batch_row["Storage Type"], batch_row["Material"], bin_group["Batch Number"].iloc[0],  # Batch from Open Spaces
                            batch_row["Batch"],  # Original Batch from Endcaps
                            bin_group["SU Capacity"].iloc[0], 1, new_avail_su, su_id, batch_row["Total Stock"]
                        ])

                    # Update available SU count
                    available_bins.loc[bin_group.index, "Avail SU"] = new_avail_su
                    break

    # Step 5: Return the results
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

# File uploaders
st.header("Upload Files")
endcaps_file = st.file_uploader("Endcaps File", type=["xlsx", "XLSX"])  # Added uppercase
open_space_file = st.file_uploader("Open Space File", type=["xlsx", "XLSX"])  # Added uppercase

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
