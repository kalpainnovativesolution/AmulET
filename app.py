# app.py
import streamlit as st
import pandas as pd
import matplotlib.pyplot as plt
from io import BytesIO
import os
from PIL import Image


def process_excel_file(filepath: str) -> dict:
    """
    Processes an Excel file to analyze embryo transfer data.
    Returns a dictionary containing dataframes and matplotlib figures.
    """
    # Load the Excel file and read the specific sheet
    df = pd.read_excel(filepath, sheet_name='Sheet2', skiprows=1)

    # Extract necessary columns
    pregnancy_data = df[['Date of Embryo Transfer', 'Pregnacy report']].copy()

    # Convert date column
    pregnancy_data['Date of Embryo Transfer'] = pd.to_datetime(
        pregnancy_data['Date of Embryo Transfer'], 
        errors='coerce'
    )

    # Drop rows with invalid dates
    pregnancy_data.dropna(subset=['Date of Embryo Transfer'], inplace=True)

    # Create Year and Quarter columns
    pregnancy_data['Year'] = pregnancy_data['Date of Embryo Transfer'].dt.year
    pregnancy_data['Quarter'] = pregnancy_data['Date of Embryo Transfer'].dt.to_period('Q').dt.quarter

    # Filter positive pregnancies
    positive_df = pregnancy_data[
        pregnancy_data['Pregnacy report'].str.strip().str.lower() == 'positive'
    ].copy()

    # Group by Year and Quarter
    total_et_counts = pregnancy_data.groupby(['Year', 'Quarter']).size().unstack(fill_value=0)
    positive_et_counts = positive_df.groupby(['Year', 'Quarter']).size().unstack(fill_value=0)

    # Combine into a dict of summary DataFrames year-wise
    summary_dfs = {}
    figs = {}

    for year in sorted(total_et_counts.index):
        summary_df = pd.DataFrame({
            'Total ETs': total_et_counts.loc[year],
            'Positive Pregnancies': positive_et_counts.loc[year] if year in positive_et_counts.index else 0
        }).fillna(0)

        summary_dfs[year] = summary_df

        # Plot for this year
        fig, ax = plt.subplots(figsize=(8, 5))
        summary_df.plot(kind='bar', ax=ax)
        ax.set_title(f'Total vs Positive Pregnancies - {year}')
        ax.set_ylabel('Count')
        ax.set_xlabel('Quarter')
        plt.xticks(rotation=0)
        plt.tight_layout()

        figs[year] = fig

    return {
        "tables": {
            "Pregnancy Data": pregnancy_data
        } | {f"Summary {year}": df for year, df in summary_dfs.items()},
        "graphs": figs
    }

# Make sure uploads folder exists
os.makedirs("uploads", exist_ok=True)

st.set_page_config(page_title="Excel ML Analyzer", layout="wide")

# Display logo on the top-right
logo_path = "uploads/Prompt_Logo.png"  # Adjust path if logo is in your GitHub repo or elsewhere

# Save the uploaded logo file (if not already in the repo)
with open(logo_path, "wb") as f:
    f.write(open("Prompt_Logo.png", "rb").read())  # You can copy or move logo to uploads folder manually

# Render with column layout
col1, col2 = st.columns([6, 1])
with col1:
    st.title("📊 AMUL ET Pregnancy: Data-Driven Insights")
with col2:
    logo = Image.open(logo_path)
    st.image(logo, use_container_width=True)

uploaded_file = st.file_uploader("Upload your Excel file", type=["xlsx"])

if uploaded_file is not None:
    with st.spinner("Processing..."):
        # Save the file
        filepath = os.path.join("uploads", uploaded_file.name)
        with open(filepath, "wb") as f:
            f.write(uploaded_file.getbuffer())
        
        # Call your ML processing function
        result = process_excel_file(filepath)

        # Display tables
        for name, df in result.get("tables", {}).items():
            st.subheader(f"📄 Table: {name}")
            st.dataframe(df)

        # Display graphs
        for year, fig in result.get("graphs", {}).items():
            st.subheader(f"📈 Graph: {year}")
            st.pyplot(fig)
