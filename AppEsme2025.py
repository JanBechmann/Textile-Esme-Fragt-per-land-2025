import streamlit as st
import pandas as pd

# Liste over EU-lande
EU_COUNTRIES = {
    "AT", "BE", "BG", "HR", "CY", "CZ", "DK", "EE", "FI", "FR", "DE", "GR", "HU", "IE", "IT",
    "LV", "LT", "LU", "MT", "NL", "PL", "PT", "RO", "SK", "SI", "ES", "SE"
}

# Landekoder til landenavne (alle ISO-lande)
COUNTRY_NAMES = {
    "AT": "Austria", "BE": "Belgium", "BG": "Bulgaria", "HR": "Croatia", "CY": "Cyprus", "CZ": "Czech Republic",
    "DK": "Denmark", "EE": "Estonia", "FI": "Finland", "FR": "France", "DE": "Germany", "GR": "Greece", "HU": "Hungary",
    "IE": "Ireland", "IT": "Italy", "LV": "Latvia", "LT": "Lithuania", "LU": "Luxembourg", "MT": "Malta",
    "NL": "Netherlands", "PL": "Poland", "PT": "Portugal", "RO": "Romania", "SK": "Slovakia", "SI": "Slovenia",
    "ES": "Spain", "SE": "Sweden", "NO": "Norway", "GB": "United Kingdom", "CH": "Switzerland", "US": "United States"
}

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
            
            # Hvis fanen er "UPS DE", så sum også kolonne 17 (index 16)
            if sheet == "UPS DE":
                df["Price"] += pd.to_numeric(df.iloc[:, 16].str.replace(",", ".", regex=False), errors='coerce').fillna(0)
            
            subtotal = df.groupby(df.iloc[:, 8])["Price"].sum().reset_index()
            subtotal.columns = ["Receiver Country", "Subtotal Price"]
            final_subtotals[sheet] = subtotal
        except Exception:
            final_subtotals[sheet] = pd.DataFrame(columns=["Receiver Country", "Subtotal Price"])
    
    subtotal_combined = pd.concat(final_subtotals.values(), keys=final_subtotals.keys())
    total_per_country = subtotal_combined.groupby("Receiver Country")["Subtotal Price"].sum().reset_index()
    
    # Tilføj landenavn
    total_per_country["Country Name"] = total_per_country["Receiver Country"].map(COUNTRY_NAMES)

    # Opdel i "Med Moms" og "Uden Moms"
    total_per_country["Med Moms"] = total_per_country.apply(
        lambda row: row["Subtotal Price"] if row["Receiver Country"] in EU_COUNTRIES else 0, axis=1)
    total_per_country["Uden Moms"] = total_per_country.apply(
        lambda row: row["Subtotal Price"] if row["Receiver Country"] not in EU_COUNTRIES and row["Country Name"] else 0, axis=1)

    # Beregn total og formater tal
    total_per_country["Subtotal Price"] = total_per_country["Subtotal Price"].round(2)
    total_per_country["Med Moms"] = total_per_country["Med Moms"].round(2)
    total_per_country["Uden Moms"] = total_per_country["Uden Moms"].round(2)
    total_per_country["Total"] = total_per_country["Subtotal Price"].round(2)

    # Fjern tomme rækker
    total_per_country = total_per_country[total_per_country["Receiver Country"].str.strip() != ""]
    total_per_country = total_per_country[total_per_country["Receiver Country"] != "Receiver Country"]

    # Beregn Grand Total
    grand_total = total_per_country[["Subtotal Price", "Med Moms", "Uden Moms", "Total"]].sum().round(2)
    grand_total_row = pd.DataFrame([["GRAND TOTAL", "", grand_total["Subtotal Price"], grand_total["Med Moms"], grand_total["Uden Moms"], grand_total["Total"]]],
                                   columns=["Receiver Country", "Country Name", "Subtotal Price", "Med Moms", "Uden Moms", "Total"])
    
    final_totals = pd.concat([total_per_country, grand_total_row], ignore_index=True)

    return final_totals

# Streamlit UI
st.title("Fragt Data Analyse")

uploaded_file = st.file_uploader("Upload en Excel-fil", type=["xlsx"])

if uploaded_file:
    result_df = process_excel(uploaded_file)
    st.write("### Resultat: Subtotal per Receiver Country")
    st.dataframe(result_df)
    
    # Download-knap
    csv = result_df.to_csv(index=False, decimal=',').encode('utf-8')
    st.download_button("Download resultater som CSV", csv, "fragt_data.csv", "text/csv")
