import streamlit as st
import pandas as pd

# Load product data from Excel
@st.cache_data
def load_product_data():
    df = pd.read_excel("VesmaVloer.Collection.Website.2025 (3).xlsx")
    df = df[["model", "merk", "COMISSIE prijs /m2/m (ex. btw)"]].dropna()
    df.columns = ["Model", "Merk", "Prijs_per_m2"]  # Rename for consistency
    return df

df_products = load_product_data()

st.title("🧾 Vesma Vloer Offerte Generator")

st.header("1. Klantinformatie")
klant_naam = st.text_input("Naam van de klant")

st.header("2. Productkeuze")
product_model = st.selectbox("Kies een laminaat product", df_products["Model"].unique())

# Fetch product details
product_info = df_products[df_products["Model"] == product_model].iloc[0]
prijs_per_m2 = product_info["Prijs_per_m2"]
merk = product_info["Merk"]

st.markdown(f"**Merk:** {merk}")
st.markdown(f"**Prijs per m²:** €{prijs_per_m2:.2f} (excl. BTW)")

st.header("3. Oppervlakte")
aantal_m2 = st.number_input("Hoeveel m² heeft de klant nodig?", min_value=1)

totaal_prijs = prijs_per_m2 * aantal_m2
btw = totaal_prijs * 0.21
totaal_incl_btw = totaal_prijs + btw

st.header("4. Offerte Overzicht")
st.markdown(f"**Klant:** {klant_naam}")
st.markdown(f"**Product:** {product_model}")
st.markdown(f"**Merk:** {merk}")
st.markdown(f"**Aantal m²:** {aantal_m2}")
st.markdown(f"**Totaal (excl. BTW):** €{totaal_prijs:.2f}")
st.markdown(f"**BTW (21%):** €{btw:.2f}")
st.markdown(f"**Totaal (incl. BTW):** €{totaal_incl_btw:.2f}")

# Optionally download as XLS
if st.button("💾 Genereer XLS offerte"):
    offerte_df = pd.DataFrame({
        "Klant": [klant_naam],
        "Product": [product_model],
        "Merk": [merk],
        "Aantal m²": [aantal_m2],
        "Prijs/m²": [prijs_per_m2],
        "Totaal (excl. BTW)": [totaal_prijs],
        "BTW (21%)": [btw],
        "Totaal (incl. BTW)": [totaal_incl_btw]
    })

    offerte_df.to_excel("offerte_output.xlsx", index=False)
    with open("offerte_output.xlsx", "rb") as f:
        st.download_button("📥 Download offerte als XLSX", f, file_name="Offerte-VesmaVloer.xlsx")

