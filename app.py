# app.py
import streamlit as st
import pandas as pd
import matplotlib.pyplot as plt
from io import BytesIO

import os

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

    # Create Quarter column
    pregnancy_data['Quarter'] = pregnancy_data['Date of Embryo Transfer'].dt.to_period('Q')

    # Filter positive pregnancies
    positive_df = pregnancy_data[
        pregnancy_data['Pregnacy report'].str.strip().str.lower() == 'positive'
    ].copy()
    positive_df['Quarter'] = positive_df['Date of Embryo Transfer'].dt.to_period('Q')

    # Count total ETs and positive pregnancies per quarter
    total_et_counts = pregnancy_data['Quarter'].value_counts().sort_index()
    positive_et_counts = positive_df['Quarter'].value_counts().sort_index()

    # Combine into a summary DataFrame
    summary_df = pd.DataFrame({
        'Total ETs': total_et_counts,
        'Positive Pregnancies': positive_et_counts
    }).fillna(0)

    # Create bar chart
    fig, ax = plt.subplots(figsize=(10, 6))
    summary_df.plot(kind='bar', ax=ax)
    ax.set_title('Total vs Positive Pregnancies per Quarter')
    ax.set_ylabel('Count')
    ax.set_xlabel('Quarter')
    plt.xticks(rotation=45)
    plt.tight_layout()

    return {
        "tables": {
            "Pregnancy Data": pregnancy_data,
            "Summary": summary_df
        },
        "graphs": {
            "Pregnancy Chart": fig
        }
    }

# Make sure uploads folder exists
os.makedirs("uploads", exist_ok=True)

st.set_page_config(page_title="Excel ML Analyzer", layout="wide")

st.title("📊 Excel File Analyzer with ML")

uploaded_file = st.file_uploader("Upload your Excel file", type=["xlsx"])

if uploaded_file is not None:
    with st.spinner("Processing..."):
        # Save the file
        filepath = os.path.join("uploads", uploaded_file.name)
        with open(filepath, "wb") as f:
            f.write(uploaded_file.getbuffer())
        
        # Call your ML processing function
        result = process_excel_file(filepath)  # assuming this returns dict of {tables, graphs}

        # Display tables
        for name, df in result.get("tables", {}).items():
            st.subheader(f"📄 Table: {name}")
            st.dataframe(df)

        # Display graphs
        for name, fig in result.get("graphs", {}).items():
            st.subheader(f"📈 Graph: {name}")
            st.plotly_chart(fig)  # Or st.pyplot(fig) based on your graph library
