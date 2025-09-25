import streamlit as st
import pandas as pd
import re
from collections import defaultdict

st.title("📊 Automated BOM Data Quality Checker")

# -----------------------------
# 1️⃣ Upload BOM Excel file
# -----------------------------
uploaded_file = st.file_uploader("Upload BOM Excel file", type=["xls", "xlsx"])

if uploaded_file is not None:
    # Read Excel file with appropriate engine
    if uploaded_file.name.endswith('.xls'):
        df_raw = pd.read_excel(uploaded_file, skiprows=6, engine='xlrd')
    else:
        df_raw = pd.read_excel(uploaded_file, skiprows=6, engine='openpyxl')

    st.success("✅ File uploaded successfully!")
    st.subheader("Preview of Raw Data")
    st.dataframe(df_raw.head())

    # -----------------------------
    # 2️⃣ Select Required Columns
    # -----------------------------
    required_columns = ["Level", "Subject number", "Manufacturer number", "Number", "Item text"]
    selected_columns = [col for col in required_columns if col in df_raw.columns]
    df = df_raw[selected_columns].copy()

    # Filter only alphanumeric Subject Numbers
    df_alphanumeric = df[df['Subject number'].str.contains(r'\d', na=False)].copy()

    # -----------------------------
    # 3️⃣ Check duplicate Subject Numbers
    # -----------------------------
    duplicates = df_alphanumeric[df_alphanumeric['Subject number'].duplicated(keep=False)]
    if duplicates.empty:
        subject_unique_result = "Yes"
    else:
        subject_unique_result = "No"

    # -----------------------------
    # 4️⃣ Subjects with Number = 0
    # -----------------------------
    zero_number = df_alphanumeric[df_alphanumeric['Number'] == 0]
    if zero_number.empty:
        zero_number_result = "No"
    else:
        zero_number_result = "Yes"

    # -----------------------------
    # 5️⃣ Check Number vs Item text count
    # -----------------------------
    def count_items(item_text):
        if pd.isna(item_text) or str(item_text).strip() == "":
            return "NA"
        text = str(item_text).strip()
        # Fully alphabetical -> 1
        if re.fullmatch(r'[A-Za-z\s]+', text):
            return 1
        # Remove prefixes like BS:, VS:, RS:, TOP:, BOTTOM:
        text_cleaned = re.sub(r'^[A-Z]+:\s*', '', text, flags=re.IGNORECASE)
        parts = [part.strip() for part in re.split(r',', text_cleaned) if part.strip()]
        return len(parts)

    df_alphanumeric['Item_Count'] = df_alphanumeric['Item text'].apply(count_items)
    number_mismatches = df_alphanumeric[
        (df_alphanumeric['Item_Count'] != df_alphanumeric['Number']) &
        (df_alphanumeric['Item_Count'] != 'NA')
    ]
    number_match_result = "Yes" if number_mismatches.empty else "No"

    # -----------------------------
    # 6️⃣ Check Item text uniqueness across dataset
    # -----------------------------
    def get_duplicate_items_with_subjects(df, item_col='Item text', subject_col='Subject number'):
        item_to_subjects = defaultdict(list)
        for idx, row in df.iterrows():
            text = row[item_col]
            subject = row[subject_col]
            if pd.isna(text) or str(text).strip() == '':
                continue
            s = re.sub(r'\b[A-Z]+\s*:\s*', '', str(text), flags=re.IGNORECASE)
            items = [i.strip().upper() for i in re.split(r'[,\n;]+', s) if i.strip()]
            for item in items:
                item_to_subjects[item].append(subject)
        duplicates = {item: subjects for item, subjects in item_to_subjects.items() if len(subjects) > 1}
        if not duplicates:
            return pd.DataFrame(), "Yes"
        else:
            duplicate_rows = []
            for item, subjects in duplicates.items():
                for sub in subjects:
                    duplicate_rows.append({'Item': item, 'Subject number': sub})
            return pd.DataFrame(duplicate_rows), "No"

    duplicate_items_df, item_unique_result = get_duplicate_items_with_subjects(df_alphanumeric)

    # -----------------------------
    # 7️⃣ Check missing Manufacturer numbers
    # -----------------------------
    missing_manufacturer_df = df_alphanumeric[
        df_alphanumeric['Manufacturer number'].isna() |
        (df_alphanumeric['Manufacturer number'].astype(str).str.strip() == "")
    ]
    missing_manufacturer_result = "No" if missing_manufacturer_df.empty else "Yes"

    # -----------------------------
    # 8️⃣ Summary Table
    # -----------------------------
    summary = pd.DataFrame({
        "Check": [
            "All Subject Numbers unique",
            "Subjects with Number=0",
            "Number matches Item text count",
            "All items in Item text unique",
            "Missing Manufacturer numbers"
        ],
        "Result": [
            subject_unique_result,
            zero_number_result,
            number_match_result,
            item_unique_result,
            missing_manufacturer_result
        ]
    })

    st.subheader("📄 Summary")
    st.dataframe(summary)

    # -----------------------------
    # 9️⃣ Download Clean Excel Report
    # -----------------------------
    output_file = "bom_quality_report.xlsx"
    with pd.ExcelWriter(output_file, engine="openpyxl") as writer:
        summary.to_excel(writer, sheet_name="Summary", index=False)
        if not duplicates.empty:
            duplicates.to_excel(writer, sheet_name="Duplicate Subjects", index=False)
        if not zero_number.empty:
            zero_number.to_excel(writer, sheet_name="Number Zero Cases", index=False)
        if not number_mismatches.empty:
            number_mismatches.to_excel(writer, sheet_name="Number Mismatches", index=False)
        if not duplicate_items_df.empty:
            duplicate_items_df.to_excel(writer, sheet_name="Duplicate Items", index=False)
        if not missing_manufacturer_df.empty:
            missing_manufacturer_df.to_excel(writer, sheet_name="Missing Manufacturer", index=False)

    with open(output_file, "rb") as f:
        st.download_button("📥 Download Clean Excel Report", f, file_name="bom_quality_report.xlsx")
