import pandas as pd
import os
import re
from openpyxl import load_workbook
from datetime import datetime

def is_sheet_empty(df):
    """Check if sheet only has header row (first row) and rest is empty"""
    if len(df) <= 1:
        return True
    # Check if all rows after the first are empty
    if len(df) > 1:
        data_rows = df.iloc[1:]
        if data_rows.dropna(how='all').empty:
            return True
    return False

def extract_group_letter(sheet_name):
    """Extract the group letter from sheet name (e.g., '2a.WBPth' -> 'a')"""
    # Pattern: number + letter + "." + word
    match = re.match(r'\d+([a-zA-Z])\.', sheet_name)
    if match:
        return match.group(1).lower()
    return None

def add_minutes_from_time_zero(df, col_b):
    """
    Add a column showing minutes before/after time 0.
    Time 0 is defined as the "Deliver Compound" row.
    Rows before "Deliver Compound" have negative minutes.
    Rows after "Measuring Antidote" have positive minutes.
    The marker rows themselves are time 0.
    
    Uses row position as fallback when time-based calculation would result in all zeros
    (e.g., when pre-treatment period is less than 1 minute).
    """
    if col_b is None or col_b not in df.columns:
        return df
    
    # Find marker rows
    deliver_mask = df[col_b].astype(str).str.contains('Deliver Compound', case=False, na=False)
    antidote_mask = df[col_b].astype(str).str.contains('Measuring Antidote', case=False, na=False)
    
    deliver_indices = df.index[deliver_mask].tolist()
    antidote_indices = df.index[antidote_mask].tolist()
    
    if not deliver_indices:
        print(f"    Warning: No 'Deliver Compound' marker found, cannot calculate minutes from time 0")
        df['Minutes_from_Time0'] = None
        return df
    
    # Use the first "Deliver Compound" as the reference point
    deliver_idx = deliver_indices[0]
    antidote_idx = antidote_indices[0] if antidote_indices else deliver_idx
    
    # Get the time of the markers
    deliver_time = df.loc[deliver_idx, 'Time'] if 'Time' in df.columns else None
    antidote_time = df.loc[antidote_idx, 'Time'] if 'Time' in df.columns else None
    
    # Calculate minutes from time 0 for each row
    minutes_list = []
    
    for idx in df.index:
        row_time = df.loc[idx, 'Time'] if 'Time' in df.columns else None
        
        # Check if this is a marker row
        is_deliver = deliver_mask.loc[idx] if idx in deliver_mask.index else False
        is_antidote = antidote_mask.loc[idx] if idx in antidote_mask.index else False
        
        if is_deliver or is_antidote:
            minutes_list.append(0)
        elif row_time is not None and deliver_time is not None:
            try:
                # Parse times to calculate difference
                row_time_str = str(row_time)
                deliver_time_str = str(deliver_time)
                antidote_time_str = str(antidote_time) if antidote_time else deliver_time_str
                
                # Convert HH:MM to minutes
                def time_to_minutes(t):
                    parts = t.split(':')
                    return int(parts[0]) * 60 + int(parts[1])
                
                row_minutes = time_to_minutes(row_time_str)
                deliver_minutes = time_to_minutes(deliver_time_str)
                antidote_minutes = time_to_minutes(antidote_time_str)
                
                # Determine if this row is before "Deliver Compound" or after "Measuring Antidote"
                if idx < deliver_idx:
                    # Before Deliver Compound - negative minutes from Deliver
                    diff = row_minutes - deliver_minutes
                    # If diff is 0 but row is clearly before marker, use row-based offset
                    if diff == 0:
                        # Calculate relative position: negative for pre-treatment
                        diff = idx - deliver_idx  # Will be negative
                    minutes_list.append(diff)
                elif idx > antidote_idx:
                    # After Measuring Antidote - positive minutes from Antidote
                    diff = row_minutes - antidote_minutes
                    minutes_list.append(diff)
                else:
                    # Between markers (shouldn't happen as we deleted these)
                    minutes_list.append(0)
            except Exception as e:
                minutes_list.append(None)
        else:
            minutes_list.append(None)
    
    df['Minutes_from_Time0'] = minutes_list
    return df

def process_excel_file(input_path):
    """Process a single Excel file according to the requirements"""
    print(f"\n{'='*60}")
    print(f"Processing: {input_path}")
    print(f"{'='*60}")
    
    # Load the workbook to get sheet names
    xl = pd.ExcelFile(input_path)
    sheet_names = xl.sheet_names
    print(f"Found {len(sheet_names)} sheets: {sheet_names}")
    
    processed_sheets = {}
    
    for sheet_name in sheet_names:
        print(f"\n--- Processing sheet: {sheet_name} ---")
        
        # Read the sheet
        df = pd.read_excel(input_path, sheet_name=sheet_name)
        
        # Check if sheet is empty (only header row)
        if is_sheet_empty(df):
            print(f"  Sheet '{sheet_name}' is empty (only header row). Skipping...")
            continue
        
        print(f"  Original shape: {df.shape}")
        print(f"  Columns: {list(df.columns)}")
        
        # Step 1: Split date and time column (Column A - first column)
        first_col = df.columns[0]
        print(f"  First column name: '{first_col}'")
        
        # Check sample values
        sample_val = df[first_col].iloc[0] if len(df) > 0 else None
        print(f"  Sample value in first column: {sample_val} (type: {type(sample_val).__name__})")
        
        # Try to parse as datetime if it's not already
        datetime_col = None
        try:
            if pd.api.types.is_datetime64_any_dtype(df[first_col]):
                datetime_col = df[first_col]
            else:
                datetime_col = pd.to_datetime(df[first_col], errors='coerce')
            
            # Check if we have valid datetime data
            valid_count = datetime_col.notna().sum()
            print(f"  Valid datetime values: {valid_count}/{len(df)}")
            
            if valid_count > 0:
                # Create new columns for Date and Time
                new_cols = df.columns.tolist()
                
                # Insert Date and Time at the beginning
                df.insert(0, 'Date_New', datetime_col.dt.date)
                df.insert(1, 'Time_New', datetime_col.dt.strftime('%H:%M:%S'))
                
                # Drop the original first column (now at index 2)
                original_col_name = df.columns[2]
                df = df.drop(columns=[original_col_name])
                
                # Rename the new columns
                df = df.rename(columns={'Date_New': 'Date', 'Time_New': 'Time'})
                
                print(f"  Split datetime column into 'Date' and 'Time'")
                print(f"  New columns: {list(df.columns[:5])}...")
            else:
                print(f"  Warning: No valid datetime values found, keeping original column")
        except Exception as e:
            print(f"  Warning: Could not split datetime column: {e}")
        
        # Step 2: Delete rows between "Deliver Compound" and "Measuring Antidote" in column B
        # IMPORTANT: Keep the rows that contain "Deliver Compound" or "Measuring Antidote" - only delete rows in between
        # Column B is now the 3rd column (index 2) after splitting
        marker_rows_df = None  # Will store marker rows to preserve them through averaging
        col_b = None
        
        if len(df.columns) > 2:
            col_b = df.columns[2]
            print(f"  Looking for markers in column: '{col_b}'")
            
            # Find indices of marker rows
            deliver_mask = df[col_b].astype(str).str.contains('Deliver Compound', case=False, na=False)
            antidote_mask = df[col_b].astype(str).str.contains('Measuring Antidote', case=False, na=False)
            
            deliver_indices = df.index[deliver_mask].tolist()
            antidote_indices = df.index[antidote_mask].tolist()
            
            # These marker rows will be preserved
            marker_indices = deliver_indices + antidote_indices
            
            print(f"  Found 'Deliver Compound' at indices: {deliver_indices} (will be KEPT)")
            print(f"  Found 'Measuring Antidote' at indices: {antidote_indices} (will be KEPT)")
            
            # Save marker rows BEFORE any deletion (to preserve through averaging)
            if marker_indices:
                marker_rows_df = df.loc[marker_indices].copy()
                print(f"  Saved {len(marker_rows_df)} marker rows to preserve through averaging")
            
            # Delete rows between each pair (excluding the marker rows themselves)
            rows_to_delete = []
            marker_set = set(marker_indices)
            for deliver_idx in deliver_indices:
                # Find the next "Measuring Antidote" after this "Deliver Compound"
                next_antidote = [idx for idx in antidote_indices if idx > deliver_idx]
                if next_antidote:
                    antidote_idx = next_antidote[0]
                    # Mark rows strictly between the markers (exclusive - do NOT include marker rows)
                    for idx in df.index:
                        if idx > deliver_idx and idx < antidote_idx and idx not in marker_set:
                            rows_to_delete.append(idx)
                    print(f"  Marking rows between index {deliver_idx} and {antidote_idx} for deletion (markers preserved)")
            
            # Remove duplicates
            rows_to_delete = list(set(rows_to_delete))
            
            if rows_to_delete:
                df = df.drop(index=rows_to_delete)
                print(f"  Deleted {len(rows_to_delete)} rows between markers (marker rows preserved)")
            else:
                print(f"  No rows to delete between markers")
        
        # Reset index after deletion
        df = df.reset_index(drop=True)
        
        # Step 3: Calculate average values for each minute
        # Group by Date and Time (truncated to minute) and calculate mean
        # IMPORTANT: Exclude marker rows from averaging, then add them back
        if 'Time' in df.columns:
            # Identify marker rows in the current dataframe (after deletion)
            marker_mask = pd.Series([False] * len(df), index=df.index)
            if col_b is not None and col_b in df.columns:
                marker_mask = (
                    df[col_b].astype(str).str.contains('Deliver Compound', case=False, na=False) |
                    df[col_b].astype(str).str.contains('Measuring Antidote', case=False, na=False)
                )
            
            # Separate marker rows from data rows
            df_markers = df[marker_mask].copy()
            df_data = df[~marker_mask].copy()
            
            print(f"  Separating {len(df_markers)} marker rows from {len(df_data)} data rows for averaging")
            
            # Extract minute from time (HH:MM format) for data rows
            df_data['Minute'] = df_data['Time'].astype(str).str[:5]  # Get HH:MM
            
            # Get numeric columns for averaging
            numeric_cols = df_data.select_dtypes(include=['number']).columns.tolist()
            print(f"  Numeric columns for averaging: {numeric_cols}")
            
            if numeric_cols and len(df_data) > 0:
                # Group by Date and Minute
                group_cols = []
                if 'Date' in df_data.columns:
                    group_cols.append('Date')
                group_cols.append('Minute')
                
                # Create aggregation dictionary for numeric columns
                agg_dict = {col: 'mean' for col in numeric_cols}
                
                # For non-numeric columns (except grouping cols and Time), take first value
                non_numeric_cols = [col for col in df_data.columns 
                                   if col not in numeric_cols 
                                   and col not in group_cols 
                                   and col != 'Time']
                for col in non_numeric_cols:
                    agg_dict[col] = 'first'
                
                # Perform grouping and aggregation on DATA rows only
                df_averaged = df_data.groupby(group_cols, as_index=False).agg(agg_dict)
                
                # Rename Minute back to Time
                df_averaged = df_averaged.rename(columns={'Minute': 'Time'})
                
                print(f"  Averaged data: {len(df_data)} rows -> {len(df_averaged)} rows (by minute)")
                
                # Add Minute column to marker rows for proper time representation
                if len(df_markers) > 0:
                    df_markers = df_markers.copy()
                    df_markers['Time'] = df_markers['Time'].astype(str).str[:5]  # Truncate to HH:MM like averaged data
                
                # Combine averaged data with preserved marker rows
                df = pd.concat([df_averaged, df_markers], ignore_index=True)
                
                # Sort by Date and Time to maintain chronological order
                sort_cols = []
                if 'Date' in df.columns:
                    sort_cols.append('Date')
                if 'Time' in df.columns:
                    sort_cols.append('Time')
                if sort_cols:
                    df = df.sort_values(by=sort_cols).reset_index(drop=True)
                
                print(f"  Combined: {len(df_averaged)} averaged rows + {len(df_markers)} marker rows = {len(df)} total rows")
                
                # Reorder columns: Date, Time, then the rest
                cols = df.columns.tolist()
                if 'Date' in cols and 'Time' in cols:
                    new_order = ['Date', 'Time'] + [c for c in cols if c not in ['Date', 'Time']]
                    df = df[new_order]
            else:
                # No data to average, just keep markers
                if len(df_markers) > 0:
                    df = df_markers
                # Remove the Minute column if it exists
                if 'Minute' in df.columns:
                    df = df.drop(columns=['Minute'])
        
        # Step 4: Add "Minutes from Time 0" column
        print(f"  Adding 'Minutes_from_Time0' column...")
        df = add_minutes_from_time_zero(df, col_b)
        
        # Reorder columns to put Minutes_from_Time0 after Time
        if 'Minutes_from_Time0' in df.columns:
            cols = df.columns.tolist()
            cols.remove('Minutes_from_Time0')
            if 'Time' in cols:
                time_idx = cols.index('Time')
                cols.insert(time_idx + 1, 'Minutes_from_Time0')
            else:
                cols.insert(0, 'Minutes_from_Time0')
            df = df[cols]
        
        processed_sheets[sheet_name] = df
        print(f"  Final shape: {df.shape}")
    
    return processed_sheets

def group_sheets_by_letter(processed_sheets):
    """
    Group sheets by their letter (e.g., 2a.WBPth, 3a.WBPth -> group a)
    and average the numeric values across sheets in the same group.
    """
    print(f"\n{'='*60}")
    print("Grouping sheets by letter...")
    print(f"{'='*60}")
    
    # Organize sheets by group letter
    groups = {}
    for sheet_name, df in processed_sheets.items():
        letter = extract_group_letter(sheet_name)
        if letter:
            if letter not in groups:
                groups[letter] = []
            groups[letter].append((sheet_name, df))
            print(f"  Sheet '{sheet_name}' -> group '{letter}'")
        else:
            print(f"  Sheet '{sheet_name}' does not match pattern, skipping grouping")
    
    grouped_sheets = {}
    
    for letter, sheet_list in sorted(groups.items()):
        print(f"\n--- Creating 'group {letter}' from {len(sheet_list)} sheets ---")
        sheet_names = [s[0] for s in sheet_list]
        print(f"  Sheets: {sheet_names}")
        
        # Get all dataframes for this group
        dfs = [s[1].copy() for s in sheet_list]
        
        # We need to align sheets by "Minutes_from_Time0" for averaging
        # First, check if all sheets have this column
        if not all('Minutes_from_Time0' in df.columns for df in dfs):
            print(f"  Warning: Not all sheets have 'Minutes_from_Time0' column")
            continue
        
        # Get all unique Minutes_from_Time0 values across all sheets
        all_minutes = set()
        for df in dfs:
            all_minutes.update(df['Minutes_from_Time0'].dropna().unique())
        all_minutes = sorted(all_minutes)
        print(f"  Unique time points: {len(all_minutes)}")
        
        # Get numeric columns (excluding Minutes_from_Time0)
        numeric_cols = []
        for df in dfs:
            for col in df.select_dtypes(include=['number']).columns:
                if col not in numeric_cols and col != 'Minutes_from_Time0':
                    numeric_cols.append(col)
        print(f"  Numeric columns to average: {numeric_cols}")
        
        # Get non-numeric columns (for reference, take from first sheet)
        non_numeric_cols = [col for col in dfs[0].columns 
                          if col not in numeric_cols 
                          and col != 'Minutes_from_Time0']
        
        # Create the grouped dataframe
        result_data = []
        
        for minute in all_minutes:
            row_data = {'Minutes_from_Time0': minute}
            
            # Collect values from all sheets for this minute
            minute_rows = []
            for df in dfs:
                matching_rows = df[df['Minutes_from_Time0'] == minute]
                if len(matching_rows) > 0:
                    minute_rows.append(matching_rows.iloc[0])
            
            if minute_rows:
                # Average numeric columns
                for col in numeric_cols:
                    values = []
                    for row in minute_rows:
                        if col in row.index and pd.notna(row[col]):
                            values.append(row[col])
                    if values:
                        row_data[col] = sum(values) / len(values)
                    else:
                        row_data[col] = None
                
                # Take first non-null value for non-numeric columns
                for col in non_numeric_cols:
                    for row in minute_rows:
                        if col in row.index and pd.notna(row[col]):
                            row_data[col] = row[col]
                            break
                    else:
                        row_data[col] = None
            
            result_data.append(row_data)
        
        # Create the grouped dataframe
        grouped_df = pd.DataFrame(result_data)
        
        # Reorder columns: put Minutes_from_Time0 first, then non-numeric, then numeric
        cols_order = ['Minutes_from_Time0']
        for col in non_numeric_cols:
            if col in grouped_df.columns and col not in cols_order:
                cols_order.append(col)
        for col in numeric_cols:
            if col in grouped_df.columns and col not in cols_order:
                cols_order.append(col)
        
        # Add any remaining columns
        for col in grouped_df.columns:
            if col not in cols_order:
                cols_order.append(col)
        
        grouped_df = grouped_df[[c for c in cols_order if c in grouped_df.columns]]
        
        group_name = f"group {letter}"
        grouped_sheets[group_name] = grouped_df
        print(f"  Created '{group_name}' with shape: {grouped_df.shape}")
    
    return grouped_sheets

def save_processed_file(processed_sheets, original_path, suffix="_processed"):
    """Save processed sheets to a new Excel file"""
    # Create output filename
    base_name = os.path.splitext(os.path.basename(original_path))[0]
    output_path = os.path.join(os.path.dirname(original_path), f"{base_name}{suffix}.xlsx")
    
    if processed_sheets:
        with pd.ExcelWriter(output_path, engine='openpyxl') as writer:
            for sheet_name, df in processed_sheets.items():
                # Excel sheet names have a 31 character limit
                safe_name = sheet_name[:31]
                df.to_excel(writer, sheet_name=safe_name, index=False)
        print(f"\nSaved processed file to: {output_path}")
        return output_path
    else:
        print(f"\nNo sheets to save for {original_path}")
        return None

def main():
    # Directory containing the Excel files
    directory = r"C:\Users\user\ambs"
    
    # Find all xlsx files
    xlsx_files = [f for f in os.listdir(directory) if f.endswith('.xlsx') and not f.startswith('~$')]
    
    print(f"Found {len(xlsx_files)} Excel files: {xlsx_files}")
    
    for xlsx_file in xlsx_files:
        if '_processed' in xlsx_file or '_grouped' in xlsx_file:
            print(f"Skipping already processed file: {xlsx_file}")
            continue
            
        input_path = os.path.join(directory, xlsx_file)
        
        try:
            # Step 1: Process individual sheets
            processed_sheets = process_excel_file(input_path)
            
            # Save individual processed sheets
            save_processed_file(processed_sheets, input_path, suffix="_processed")
            
            # Step 2: Group sheets by letter and average
            grouped_sheets = group_sheets_by_letter(processed_sheets)
            
            # Save grouped sheets
            if grouped_sheets:
                save_processed_file(grouped_sheets, input_path, suffix="_grouped")
            
        except Exception as e:
            print(f"Error processing {xlsx_file}: {e}")
            import traceback
            traceback.print_exc()

if __name__ == "__main__":
    main()
