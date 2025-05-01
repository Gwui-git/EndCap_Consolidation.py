def process_files(endcaps_file, open_space_file, selected_types):
    # Read Excel files
    try:
        endcaps_df = pd.read_excel(endcaps_file, sheet_name="Sheet1")
        open_space_df = pd.read_excel(open_space_file, sheet_name="Sheet1")
    except Exception as e:
        st.error(f"Failed to read Excel files: {e}")
        return None

    # Initialize assignments list at the start of processing
    assignments = []

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
            # Check if all batches in the SU are within the date range
            valid_match = True
            for _, su_batch_row in su_group.iterrows():
                su_batch_date = su_batch_row["Batch Date"]
                if pd.isna(su_batch_date):
                    valid_match = False
                    break

                date_differences = bin_group["Batch Date"].apply(lambda x: abs((x - su_batch_date).days))
                if not all(date_differences <= 364):
                    valid_match = False
                    break

            if valid_match:
                avail_su = bin_group["Avail SU"].iloc[0]
                if avail_su >= 1:
                    new_avail_su = max(0, avail_su - 1)

                    for _, batch_row in su_group.iterrows():
                        assignments.append([
                            bin_group["Storage Type"].iloc[0], storage_bin, bin_moving_from,
                            batch_row["Storage Type"], batch_row["Material"], 
                            bin_group["Batch Number"].iloc[0], batch_row["Batch"],
                            bin_group["SU Capacity"].iloc[0], 1, new_avail_su, 
                            su_id, batch_row["Total Stock"]
                        ])

                    available_bins.loc[bin_group.index, "Avail SU"] = new_avail_su
                    break

    # Step 5: Return the results
    if assignments:  # Now this will work since assignments is always defined
        final_output = pd.DataFrame(assignments, columns=[
            "Open Space Storage Type", "Storage Bin", "Bin Moving From", "Endcap Storage Type",
            "Material", "Batch", "Original Batch", "SU Capacity", "SU Count", 
            "Avail SU", "Storage Unit", "Total Stock"
        ])
        return final_output
    return None
