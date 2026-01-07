import pandas as pd
import os
import zipfile
import config

# -------------------
# Merge Function
# -------------------
def merge_files_flexible(filepaths, merge_options, col_values=None):
    """
    merge_options: per file: 'all', 'first_n', 'last_n', 'from_col'
    col_values: per file, number of columns or start column (1-based)
    """
    import pandas as pd
    import os

    dataframes = []

    for idx, fp in enumerate(filepaths):
        df = pd.read_excel(fp)
        option = merge_options[idx].lower()
        val = col_values[idx]

        if option == 'all':
            df_slice = df
        elif option == 'first_n' and val:
            df_slice = df.iloc[:, :val]
        elif option == 'last_n' and val:
            df_slice = df.iloc[:, -val:]
        elif option == 'from_col' and val:
            df_slice = df.iloc[:, val-1:]  # 1-based input
        else:
            df_slice = df  # fallback

        dataframes.append(df_slice.reset_index(drop=True))

    merged_df = pd.concat(dataframes, axis=1)
    merged_path = os.path.join(config.UPLOAD_FOLDER, 'merged_flexible.xlsx')
    merged_df.to_excel(merged_path, index=False)
    return merged_path




# -------------------
# Split Function
# -------------------
def split_file_custom(filepath, split_size=None, split_column=None):
    """
    Split an Excel file into multiple files.
    - split_size: number of rows per split
    - split_column: if provided, splits based on unique values in this column
    """
    df = pd.read_excel(filepath)

    split_files = []

    if split_column and split_column in df.columns:
        # Split by column values
        for val, group in df.groupby(split_column):
            split_filename = f"{os.path.splitext(os.path.basename(filepath))[0]}_{val}.xlsx"
            split_path = os.path.join(config.UPLOAD_FOLDER, split_filename)
            group.to_excel(split_path, index=False)
            split_files.append(split_path)
    elif split_size and split_size > 0:
        # Split by number of rows
        for i in range(0, len(df), split_size):
            split_df = df.iloc[i:i+split_size]
            split_filename = f"{os.path.splitext(os.path.basename(filepath))[0]}_part_{i//split_size + 1}.xlsx"
            split_path = os.path.join(config.UPLOAD_FOLDER, split_filename)
            split_df.to_excel(split_path, index=False)
            split_files.append(split_path)
    else:
        # No split, just return the same file
        return filepath

    # Zip all split files
    zip_path = os.path.join(config.UPLOAD_FOLDER, 'splits.zip')
    with zipfile.ZipFile(zip_path, 'w') as zipf:
        for f in split_files:
            zipf.write(f, os.path.basename(f))
    return zip_path


def cleanup_files(file_paths):
    """Safely delete files if they exist"""
    for path in file_paths:
        try:
            if os.path.exists(path):
                os.remove(path)
        except Exception as e:
            print(f"Cleanup failed for {path}: {e}")