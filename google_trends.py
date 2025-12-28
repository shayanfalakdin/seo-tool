from pytrends.request import TrendReq
import pandas as pd
import time
import os

# 1. Initialize TrendReq
# We remove 'retries' and 'backoff_factor' here to avoid the urllib3 TypeError
pytrends = TrendReq(
    hl='fa-IR',
    tz=270,
    timeout=(10, 25),
    requests_args={'verify': False}
)

# 2. Define Keywords
keyword = "پکیج بوتان"
modifiers = ["خرید", "قیمت", "نمایندگی", "بهترین", "نصب", "سرویس"]
all_keywords = [keyword] + [f"{m} {keyword}" for m in modifiers]

data_to_save = {}

# --- PART 1: Related Queries (Side-by-Side Separation) ---
print(f"Fetching related queries for: {keyword}...")
pytrends.build_payload(kw_list=[keyword], timeframe='today 12-m', geo='IR')
related_queries = pytrends.related_queries()

if related_queries.get(keyword):
    top = related_queries[keyword]['top']
    rising = related_queries[keyword]['rising']

    # Process Top Queries
    if top is not None:
        top = top.rename(columns={'query': 'جستجوهای برتر', 'value': 'امتیاز (0-100)'})
    else:
        top = pd.DataFrame(columns=['جستجوهای برتر', 'امتیاز (0-100)'])

    # Process Rising Queries
    if rising is not None:
        rising = rising.rename(columns={'query': 'جستجوهای نوظهور', 'value': 'رشد درصد (%)'})
        # Ensure numeric sort for rising values (e.g., 450, 130)
        rising['رشد درصد (%)'] = pd.to_numeric(rising['رشد درصد (%)'], errors='coerce')
        rising = rising.sort_values(by='رشد درصد (%)', ascending=False)
    else:
        rising = pd.DataFrame(columns=['جستجوهای نوظهور', 'رشد درصد (%)'])

    # Combine Top and Rising into different columns (axis=1)
    combined_related = pd.concat([top.reset_index(drop=True), rising.reset_index(drop=True)], axis=1)
    data_to_save['Related Queries'] = combined_related
else:
    print("No related query data found.")

# --- PART 2: Interest Over Time (Handling 5-Keyword Limit) ---
# Split list into chunks of 5 keywords each
chunks = [all_keywords[i:i + 5] for i in range(0, len(all_keywords), 5)]
interest_frames = []

for chunk in chunks:
    print(f"Fetching interest data for: {chunk}")
    try:
        pytrends.build_payload(kw_list=chunk, timeframe='today 12-m', geo='IR')
        df_chunk = pytrends.interest_over_time()
        if not df_chunk.empty:
            # Remove 'isPartial' column as it is not needed for the mean score
            interest_frames.append(df_chunk.drop(columns=['isPartial']))
    except Exception as e:
        print(f"Error fetching data for {chunk}: {e}")

    # Sleep to prevent being blocked by Google
    time.sleep(2)

# Calculate the mean score for each keyword
if interest_frames:
    final_interest = pd.concat(interest_frames, axis=1)
    mean_interest = final_interest.mean().reset_index()
    mean_interest.columns = ['کلمه کلیدی', 'میانگین امتیاز']
    # Sort by highest score
    data_to_save['Keyword Analysis'] = mean_interest.sort_values(by='میانگین امتیاز', ascending=False)

# --- PART 3: Save to Excel with RTL Support ---
if data_to_save:
    file_name = 'google_trends_report.xlsx'

    try:
        # Using XlsxWriter to support Right-to-Left (RTL) for Persian
        with pd.ExcelWriter(file_name, engine='xlsxwriter') as writer:
            for sheet_name, df in data_to_save.items():
                df.to_excel(writer, sheet_name=sheet_name, index=False)

                # Apply RTL formatting to the sheet
                worksheet = writer.sheets[sheet_name]
                worksheet.right_to_left()

        print("-" * 30)
        print(f"✅ Success! File saved at: {os.path.abspath(file_name)}")

        # Automatically try to open the file
        os.startfile(file_name)

    except PermissionError:
        print("-" * 30)
        print(f"❌ ERROR: Permission Denied.")
        print(f"Please CLOSE '{file_name}' in Excel and run the script again.")
else:
    print("No data was collected. Please check your internet connection or geo-limitations.")