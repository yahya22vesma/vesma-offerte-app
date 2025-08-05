import streamlit as st
import pandas as pd

# Load product data
@st.cache_data
def load_product_data():
    df = pd.read_excel("VesmaVloer.Collection.Website.2025 (3).xlsx", sheet_name="Sheet1")
    df = df[["Model", "Merk", "COMISSIE prijs /m2/m (ex. Btw)"]].dropna()
    df.columns = ["Model", "Merk", "Prijs_per_m2"]
    return df

df_products = load_product_data()

st.title("🧾 VesmaVloer Offerte Generator")

# Customer details
client_name = st.text_input("Naam van klant")
project_name = st.text_input("Project naam / Referentie")
date = st.date_input("Datum")

# Product selection
selected_model = st.selectbox("Kies een productmodel", df_products["Model"].unique())
product_info = df_products[df_products["Model"] == selected_model].iloc[0]
aantal_m2 = st.number_input("Aantal m²", min_value=1.0, step=0.5)

# Price calculation
prijs_per_m2 = product_info["Prijs_per_m2"]
totaal_prijs = round(prijs_per_m2 * aantal_m2, 2)

# Display offerte summary
st.subheader("📄 Offerte")
st.write(f"**Klantnaam:** {client_name}")
st.write(f"**Project:** {project_name}")
st.write(f"**Datum:** {date}")
st.write(f"**Product:** {selected_model} ({product_info['Merk']})")
st.write(f"**Prijs per m² (excl. BTW):** € {prijs_per_m2}")
st.write(f"**Aantal m²:** {aantal_m2}")
st.write(f"**Totaalprijs (excl. BTW):** € {totaal_prijs}")

# Save offerte to Excel
if st.button("💾 Genereer offerte (Excel)"):
    offerte_df = pd.DataFrame({
        "Klant": [client_name],
        "Project": [project_name],
        "Datum": [date],
        "Product": [selected_model],
        "Merk": [product_info["Merk"]],
        "Prijs per m²": [prijs_per_m2],
        "Aantal m²": [aantal_m2],
        "Totaalprijs": [totaal_prijs]
    })
    offerte_file = "offerte_output.xlsx"
    offerte_df.to_excel(offerte_file, index=False)
    with open(offerte_file, "rb") as f:
        st.download_button("📥 Download offerte", f, file_name=offerte_file)