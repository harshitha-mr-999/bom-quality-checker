import streamlit as st
import pandas as pd
import re
from collections import Counter

st.title("🔍 BOM Data Quality Checker")

# Upload Excel file
uploaded_file = st.file_uploader("Upload BOM Excel file", type=["xls","xlsx"])

if uploaded_file is not None:
    # Determine engine based on file extension
    if uploaded_file.name.endswith('.xls'):
        df = pd.read_excel(uploaded_file, skiprows=6, engine='xlrd')
    else:
        df = pd.read_excel(uploaded_file, skiprows=6, engine='openpyxl')

    st.success("✅ File uploaded successfully!")
    st.subheader("Preview of Data")
    st.dataframe(df.head())

    # -----------------------------
    # 1️⃣ Check unique Subject Numbers
    # -----------------------------
    subject_duplicates = df[df.duplicated('Subject number', keep=False)]
    if subject_duplicates.empty:
        subject_unique_result = "Yes"
        st.success("All Subject Numbers are unique ✅")
    else:
        subject_unique_result = "No"
        st.error("Duplicate Subject Numbers found ❌")
        st.dataframe(subject_duplicates)

    # -----------------------------
    # 2️⃣ Subject Numbers with missing Manufacturer number
    # -----------------------------
    missing_manufacturer = df[df['Manufacturer number'].isna() | (df['Manufacturer number'].astype(str).str.strip()=='')]
    if missing_manufacturer.empty:
        missing_manufacturer_result = "No"
        st.success("No missing Manufacturer numbers ✅")
    else:
        missing_manufacturer_result = "Yes"
        st.warning("Subject Numbers with missing Manufacturer number ❌")
        st.dataframe(missing_manufacturer[['Subject number','Manufacturer number']])

    # -----------------------------
    # 3️⃣ Number = 0
    # -----------------------------
    zero_number = df[df['Number'] == 0]
    if zero_number.empty:
        zero_number_result = "No"
        st.success("No 'Number=0' cases ✅")
    else:
        zero_number_result = "Yes"
        st.warning("Do not populate (Number=0) ❌")
        st.dataframe(zero_number[['Subject number','Number']])

    # -----------------------------
    # 4️⃣ Check if Number matches item count
    # -----------------------------
    def count_items(text):
        if pd.isna(text) or str(text).strip()=='':
            return None
        # Remove prefixes like TOP:, BOTTOM:, BS:, RS:, VS:
        s = re.sub(r'\b[A-Z]+\s*:\s*','',str(text))
        items = [i.strip() for i in re.split(r'[,\n;]+', s) if i.strip()]
        return len(items) if items else 1

    df['Item_Count'] = df['Item text'].apply(count_items)
    number_mismatches = df[df['Item_Count'].notna() & (df['Item_Count'] != df['Number'])]
    if number_mismatches.empty:
        number_match_result = "Yes"
        st.success("All Item text counts match Number column ✅")
    else:
        number_match_result = "No"
        st.error("Number column mismatches found ❌")
        st.dataframe(number_mismatches[['Subject number','Number','Item text','Item_Count']])

    # -----------------------------
    # 5️⃣ Item text uniqueness across dataset
    # -----------------------------
    all_items = []
    for text in df['Item text'].dropna():
        s = re.sub(r'\b[A-Z]+\s*:\s*', '', str(text))
        items = [i.strip().upper() for i in re.split(r'[,\n;]+', s) if i.strip()]
        all_items.extend(items)

    dup_items = {i:c for i,c in Counter(all_items).items() if c>1}
    if not dup_items:
        item_unique_result = "Yes"
        st.success("All items in Item text are unique ✅")
    else:
        item_unique_result = "No"
        st.error("Duplicate items found in Item text ❌")
        st.write(dup_items)

    # -----------------------------
    # Summary Table
    # -----------------------------
    summary = pd.DataFrame({
        "Check": [
            "All Subject Numbers unique",
            "Missing Manufacturer number",
            "Number = 0",
            "Number matches Item text count",
            "All items in Item text unique"
        ],
        "Result": [
            subject_unique_result,
            missing_manufacturer_result,
            zero_number_result,
            number_match_result,
            item_unique_result
        ]
    })

    st.subheader("✅ Summary")
    st.dataframe(summary)

    # -----------------------------
    # Download Excel Report
    # -----------------------------
    output_file = "bom_check_results.xlsx"
    with pd.ExcelWriter(output_file, engine="openpyxl") as writer:
        df.to_excel(writer, sheet_name="Original Data", index=False)
        if not subject_duplicates.empty:
            subject_duplicates.to_excel(writer, sheet_name="Duplicate Subjects", index=False)
        if not missing_manufacturer.empty:
            missing_manufacturer.to_excel(writer, sheet_name="Missing Manufacturer", index=False)
        if not zero_number.empty:
            zero_number.to_excel(writer, sheet_name="Number Zero Cases", index=False)
        if not number_mismatches.empty:
            number_mismatches.to_excel(writer, sheet_name="Number Mismatches", index=False)
        if dup_items:
            dup_df = pd.DataFrame(list(dup_items.items()), columns=["Item","Count"])
            dup_df.to_excel(writer, sheet_name="Duplicate Items", index=False)

    with open(output_file, "rb") as f:
        st.download_button("📥 Download Excel Report", f, file_name="bom_results.xlsx")
