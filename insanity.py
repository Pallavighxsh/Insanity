import sys
import pandas as pd
from openpyxl import Workbook

# ---------- Config ----------
print("\n--- Welcome to the Insanity CLI ---")
print("Please ensure your .py file and .xlsx files are on the Desktop and you are using a virtual environment.\n")

INPUT_FILE = input("Enter the path to your Excel file (e.g., ~/Desktop/dummy_categories.xlsx): ").strip()
OUTPUT_FILE_DEFAULT = "insane_workbook.xlsx"

# ---------- Load Data ----------
try:
    xls = pd.ExcelFile(INPUT_FILE)
except FileNotFoundError:
    print(f"File '{INPUT_FILE}' not found. Make sure the path is correct.")
    sys.exit(1)

all_data = []
for sheet_name in xls.sheet_names:
    df_sheet = pd.read_excel(xls, sheet_name=sheet_name)
    df_sheet["Category"] = sheet_name  # use sheet name as main category
    all_data.append(df_sheet)

df = pd.concat(all_data, ignore_index=True)

# Auto-detect columns
cols_lower = [c.lower() for c in df.columns]
title_col = next((c for c in df.columns if c.lower() in ["title", "item"]), None)
author_col = next((c for c in df.columns if c.lower() in ["authors", "client"]), None)
subcategory_col = next((c for c in df.columns if "sub" in c.lower() and "sub" not in c.lower().split("_")[0]), "Subcategory")
subsub_col = next((c for c in df.columns if "sub-sub" in c.lower()), "Sub-subcategory")

# Rename for consistency
df.rename(columns={title_col: "Item", author_col: "Client",
                   subcategory_col: "Subcategory", subsub_col: "Sub-subcategory"}, inplace=True)

# ---------- Helper Functions ----------
def get_main_categories():
    return sorted(df["Category"].dropna().unique())

def get_subcategories(category):
    return sorted(df[df["Category"] == category]["Subcategory"].dropna().unique())

def get_subsubcategories(category, subcategory):
    return sorted(df[(df["Category"] == category) & (df["Subcategory"] == subcategory)]["Sub-subcategory"].dropna().unique())

def get_items(category=None, subcategory=None, subsub=None):
    filtered = df
    if category:
        filtered = filtered[filtered["Category"] == category]
    if subcategory:
        filtered = filtered[filtered["Subcategory"] == subcategory]
    if subsub:
        filtered = filtered[filtered["Sub-subcategory"] == subsub]
    return filtered[["Category", "Subcategory", "Sub-subcategory", "Item", "Client"]]

def export_items_to_excel(items_df, output_file=OUTPUT_FILE_DEFAULT):
    wb = Workbook()
    wb.remove(wb.active)
    ws = wb.create_sheet(title="Items")
    ws.append(["Category", "Subcategory", "Sub-subcategory", "Item", "Client"])
    for _, row in items_df.iterrows():
        ws.append([row["Category"], row["Subcategory"], row["Sub-subcategory"], row["Item"], row["Client"]])
    wb.save(output_file)
    print(f"\nExported {len(items_df)} items to '{output_file}'\n")

# ---------- CLI ----------
def main():
    print("\Type:")
    print("  insanity → all main categories.")
    print("  back → go one category up.")
    print("  fix the insanity → select and export items to Excel.")
    print("  bye → see you later.\n")

    while True:
        try:
            cmd = input(">> ").strip()
        except KeyboardInterrupt:
            print("\nDetected Ctrl+C. Type 'bye' to exit gracefully.")
            continue

        if cmd.lower() == "bye":
            print("Goodbyes are never forever—but bye for now!")
            break

        elif cmd.lower() == "insanity":
            categories = get_main_categories()
            print("\nMain Categories:")
            for c in categories:
                print(f"- {c}")
            print("\nType a main category name to see its subcategories.")
            continue

        elif cmd.lower() == "fix the insanity":
            print("\nEnter a comma-separated list of categories, subcategories, or sub-subcategories to export items.")
            user_input = input("List: ").strip()
            selections = [x.strip() for x in user_input.split(",") if x.strip()]
            selected_items = pd.DataFrame()

            for sel in selections:
                # Try matching Category, Subcategory, or Sub-subcategory
                items = get_items(category=sel)
                if items.empty:
                    # maybe it's a subcategory
                    sub_items = df[df["Subcategory"] == sel]
                    if not sub_items.empty:
                        selected_items = pd.concat([selected_items, sub_items], ignore_index=True)
                        continue
                    subsub_items = df[df["Sub-subcategory"] == sel]
                    if not subsub_items.empty:
                        selected_items = pd.concat([selected_items, subsub_items], ignore_index=True)
                        continue
                else:
                    selected_items = pd.concat([selected_items, items], ignore_index=True)

            export_items_to_excel(selected_items)
            print("Type 'bye' to exit or 'insanity' to continue browsing.\n")
            continue

        # User typed a main category
        elif cmd in df["Category"].values:
            subcats = get_subcategories(cmd)
            if subcats:
                print(f"\nSubcategories of {cmd}:")
                for sc in subcats:
                    print(f"- {sc}")
                print(f"\nType a subcategory name to see its sub-subcategories or items.")
            else:
                items = get_items(category=cmd)
                for _, row in items.iterrows():
                    print(f"{row['Item']} | {row['Client']}")
            continue

        # User typed a subcategory
        elif cmd in df["Subcategory"].values:
            # Find its category
            cat = df[df["Subcategory"] == cmd]["Category"].iloc[0]
            subsubcats = get_subsubcategories(cat, cmd)
            if subsubcats:
                print(f"\nSub-subcategories of {cmd}:")
                for ssc in subsubcats:
                    print(f"- {ssc}")
                print(f"\nType a sub-subcategory name to see its items.")
            else:
                items = get_items(subcategory=cmd)
                for _, row in items.iterrows():
                    print(f"{row['Item']} | {row['Client']}")
            continue

        # User typed a sub-subcategory
        elif cmd in df["Sub-subcategory"].values:
            items = get_items(subsub=cmd)
            for _, row in items.iterrows():
                print(f"{row['Item']} | {row['Client']}")
            continue

        else:
            print("Unknown command. Type 'insanity' or 'bye'. Your favorite command is 'fix the insanity'.\n")

if __name__ == "__main__":
    main()
