import pandas as pd
import os

input_file = r'C:\Users\user139\Downloads\export-from-mizfa-tools (1).xlsx'  # The Excel file you want to clean
output_file = r'C:\Users\user139\Downloads\cleand.xlsx'  # The name of the new file

# Add any words here that you want to DELETE from your list
negative_words = ['بوتان']


def clean_only():
    if not os.path.exists(input_file):
        print(f"❌ Error: The file '{input_file}' was not found in this folder.")
        return

    df = pd.read_excel(input_file)

    # We assume the keyword is in the first column
    kw_col = df.columns[0]
    initial_count = len(df)

    df[kw_col] = df[kw_col].astype(str).str.strip()

    # the exclusion filter
    df_filtered = df.copy()
    for word in negative_words:
        df_filtered = df_filtered[~df_filtered[kw_col].str.contains(word, na=False)]

    # 5. REMOVE DUPLICATES
    df_filtered = df_filtered.drop_duplicates(subset=[kw_col])
    final_count = len(df_filtered)

    try:
        with pd.ExcelWriter(output_file, engine='xlsxwriter') as writer:
            df_filtered.to_excel(writer, index=False, sheet_name='Cleaned Keywords')

            worksheet = writer.sheets['Cleaned Keywords']
            worksheet.right_to_left()

        print("-" * 30)
        print(f"✅ Cleaning Complete!")
        print(f"Original keywords: {initial_count}")
        print(f"Removed: {initial_count - final_count}")
        print(f"Final unique keywords: {final_count}")
        print("-" * 30)

        os.startfile(output_file)

    except PermissionError:
        print(f"❌ Error: Please close '{output_file}' in Excel and run the script again.")


if __name__ == "main":
    clean_only()