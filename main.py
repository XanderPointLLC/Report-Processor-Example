
# Third Party Imports
import pandas as pd
import tkinter as tk
from tkinter import filedialog, messagebox
from pathlib import Path



# Define the Columns you want to keep at the end and output file, and initialize WARNING_COLLECTION
COLUMNS_TO_KEEP = ["Custom Column 1", "Custom Column 25", "Custom Column 136"]
WARNING_COLLECTION = []

# Access the Excel file
def choose_file():
    """Open a file dialog, read into a DataFrame, and return it."""
    # Create a hidden root window just for the dialog
    root = tk.Tk()
    root.withdraw()
    root.update()  # makes the dialog appear on top on some systems

    file_path = filedialog.askopenfilename(
        title="Select a data file",
        filetypes=(
            ("Data files", "*.csv *.xlsx *.xls *.json *.parquet"),
            ("CSV", "*.csv"),
            ("Excel", "*.xlsx *.xls")
        ),
    )
    root.destroy()

    if not file_path:
        # User cancelled
        return None

    ext = Path(file_path).suffix.lower()
    try:
        if ext == ".csv":
            df = pd.read_csv(file_path)
        elif ext in (".xlsx", ".xls"):
            df = pd.read_excel(file_path)
        else:
            messagebox.showerror("Unsupported file",
                                 f"Unsupported file type: {ext}")
            return None
        return df
    except Exception as e:
        # Show a friendly error popup and return None
        messagebox.showerror("Load error", f"Could not load file:\n{e}")
        return None

# Preprocess the file (Normalize, do calculations, pull only needed columns, etc.)
def preprocessing(df):
    """Preprocess data and return a dataframe."""
    df = df[COLUMNS_TO_KEEP]

    # Add Unique Row Identifier (start from 1 if you like Excel-style rows)
    df["Row Index"] = df.index + 1

    # Multiply Columns 1 and 25
    df.loc[:, "Custom Multiplied"] = df["Custom Column 1"] * df["Custom Column 25"]

    # Add errors where NA is found (boolean)
    df.loc[:, "Error Found"] = df.isna().any(axis=1)

    # Add a text list of which columns are missing per row
    df.loc[:, "Missing Columns"] = df.isna().apply(
        lambda r: ", ".join(df.columns[r.values]) if r.any() else "",
        axis=1
    )

    # count rows with any NA
    na_row_count = df["Error Found"].sum()  # since it's True/False, sum counts Trues

    if na_row_count > 0:
        WARNING_COLLECTION.append(
            f"Warning: {na_row_count} rows have missing or invalid values"
        )
    else:
        WARNING_COLLECTION.append("No missing or invalid values found.")

    return df

# Output the file
def output(df: pd.DataFrame):
    """Create the output report with warnings attached as a sheet."""

    # --- Step 1. Choose the output location ---
    root = tk.Tk()
    root.withdraw()
    file_path = filedialog.asksaveasfilename(
        title="Save report as",
        defaultextension=".xlsx",
        filetypes=(("Excel files", "*.xlsx"), ("All files", "*.*"))
    )
    root.destroy()
    if not file_path:
        print("No save location chosen.")
        return

    # --- Step 2. Add warnings ---
    # Find columns with NA
    cols_with_na = df.columns[df.isna().any()].tolist()

    warnings = WARNING_COLLECTION
    if cols_with_na:
        warnings.append(
            {"Warning": f"The following columns are missing values: {', '.join(cols_with_na)}"}
        )
    else:
        warnings.append({"Warning": "No missing values detected."})

    warnings_df = pd.DataFrame(warnings)

    # --- Step 3. Create the report with two sheets ---
    with pd.ExcelWriter(file_path, engine="xlsxwriter") as writer:
        df.to_excel(writer, sheet_name="Report", index=False)
        warnings_df.to_excel(writer, sheet_name="Warnings", index=False)

    print(f"Report saved to: {file_path}")

def main():
    df = choose_file()
    preprocessed = preprocessing(df)
    output(preprocessed)

if __name__ == "__main__":
    main()