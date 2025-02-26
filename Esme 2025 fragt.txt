import streamlit as st
import pandas as pd

# Funktion til at behandle Excel-data
def process_excel(file):
    xls = pd.ExcelFile(file)
    sheet_names = xls.sheet_names
    column_targets = {sheet: 15 for sheet in sheet_names}  # Alle faner bruger kolonne P (index 15)
    
    final_subtotals = {}
    
    for sheet, price_col in column_targets.items():
        try:
            df = pd.read_excel(xls, sheet_name=sheet, skiprows=5, dtype=str)
            df["Price"] = pd.to_numeric(df.iloc[:, price_col].str.replace(",", ".", regex=False), errors='coerce').fillna(0)
            subtotal = df.groupby(df.iloc[:, 8])["Price"].sum().reset_index()
            subtotal.columns = ["Receiver Country", "Subtotal Price"]
            final_subtotals[sheet] = subtotal
        except Exception as e:
            final_subtotals[sheet] = pd.DataFrame(columns=["Receiver Country", "Subtotal Price"])
    
    subtotal_combined = pd.concat(final_subtotals.values(), keys=final_subtotals.keys())
    total_per_country = subtotal_combined.groupby("Receiver Country")["Subtotal Price"].sum().reset_index()
    total_per_country["Subtotal Price"] = total_per_country["Subtotal Price"].round(2)
    grand_total = total_per_country["Subtotal Price"].sum().round(2)
    grand_total_row = pd.DataFrame([["GRAND TOTAL", grand_total]], columns=["Receiver Country", "Subtotal Price"])
    final_totals = pd.concat([total_per_country, grand_total_row], ignore_index=True)
    
    # Fjern tomme rækker
    final_totals = final_totals[final_totals["Receiver Country"].str.strip() != ""]
    final_totals = final_totals[final_totals["Receiver Country"] != "Receiver Country"]
    
    return final_totals

# Streamlit UI
st.title("Fragt Data Analyse")

uploaded_file = st.file_uploader("Upload en Excel-fil", type=["xlsx"])

if uploaded_file:
    result_df = process_excel(uploaded_file)
    st.write("### Resultat: Subtotal per Receiver Country")
    st.dataframe(result_df)
    
    # Download-knap
    csv = result_df.to_csv(index=False).encode('utf-8')
    st.download_button("Download resultater som CSV", csv, "fragt_data.csv", "text/csv")
