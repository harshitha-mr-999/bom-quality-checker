import streamlit as st
import pandas as pd
import re
from collections import Counter

st.title("🔍 BOM Data Quality Checker")

# Upload Excel file
# Check file extension
if uploaded_file.name.endswith('.xls'):
    df = pd.read_excel(uploaded_file, skiprows=6, engine='xlrd')
else:  # .xlsx
    df = pd.read_excel(uploaded_file, skiprows=6, engine='openpyxl')

if uploaded_file is not None:
    df = pd.read_excel(uploaded_file, skiprows=6)
    st.write("✅ File uploaded successfully!")
    
    st.subheader("Preview of Data")
    st.dataframe(df.head())

    # 1️⃣ Check unique Subject Numbers
    if df['Subject number'].is_unique:
        st.success("All Subject Numbers are unique ✅")
    else:
        st.error("Duplicate Subject Numbers found ❌")
        st.dataframe(df[df.duplicated('Subject number', keep=False)])

    # 2️⃣ Check Number = 0
    zero_num = df[df['Number'] == 0]
    if zero_num.empty:
        st.success("No 'Number=0' cases ✅")
    else:
        st.warning("Do not populate (Number=0) ❌")
        st.dataframe(zero_num[['Subject number','Number']])

    # 3️⃣ Check Item text uniqueness
    all_items = []
    for text in df['Item text'].dropna():
        # Remove prefixes like TOP:, BOTTOM:, BS:, RS:, VS:
        s = re.sub(r'\b[A-Z]+\s*:\s*', '', str(text))
        items = [i.strip().upper() for i in re.split(r'[,\n;]+', s) if i.strip()]
        all_items.extend(items)

    dup_items = {i:c for i,c in Counter(all_items).items() if c>1}
    if not dup_items:
        st.success("All items in Item text are unique ✅")
    else:
        st.error("Duplicate items found in Item text ❌")
        st.write(dup_items)

    # 4️⃣ Download Excel Report
    output_file = "bom_check_results.xlsx"
    with pd.ExcelWriter(output_file, engine="openpyxl") as writer:
        df.to_excel(writer, sheet_name="Original Data", index=False)
        if not zero_num.empty:
            zero_num.to_excel(writer, sheet_name="Number Zero Cases", index=False)
        if not dup_items == {}:
            dup_df = pd.DataFrame(list(dup_items.items()), columns=["Item","Count"])
            dup_df.to_excel(writer, sheet_name="Duplicate Items", index=False)

    with open(output_file, "rb") as f:
        st.download_button("📥 Download Excel Report", f, file_name="bom_results.xlsx")

