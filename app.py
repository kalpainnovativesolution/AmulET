# app.py
import streamlit as st
import pandas as pd
import matplotlib.pyplot as plt
import seaborn as sns
import plotly.express as px
import os

def process_excel_file(filepath: str) -> dict:
    df = pd.read_excel(filepath, sheet_name='Sheet2', skiprows=1)

    ###################### Feature 1 to 5 (same as previous, omitted here for brevity) ######################
    # Keep your previous 5 feature code here as we did before (Year-Quarter, Site-CL, Embryo Stage-Age, Semen Type, Organization)

    ###################### Feature 6: Dam and Sire ID Combination ######################
    df['Dam'] = df['Dam'].astype(str).str.strip().str.title()
    df['Sire id '] = df['Sire id '].astype(str).str.strip().str.title()
    df['Pregnacy report'] = df['Pregnacy report'].astype(str).str.strip().str.lower()

    summary = (
        df.groupby(['Dam', 'Sire id ', 'Pregnacy report'])
        .size()
        .unstack(fill_value=0)
        .reset_index()
    )

    summary['Total ETs'] = summary.sum(axis=1, numeric_only=True)
    summary['% Positive'] = round(summary['positive'] / summary['Total ETs'] * 100, 2) if 'positive' in summary.columns else 0.0
    summary = summary.sort_values(by='positive', ascending=False, na_position='last')

    # Plotly interactive bar chart
    fig_plotly = px.bar(
        summary,
        x='Dam',
        y='positive',
        color='Sire id ',
        title="Positive Pregnancy Report by Dam and Sire ID",
        labels={'positive': 'Number of Positive Reports'},
        text='% Positive'
    )
    fig_plotly.update_traces(texttemplate='%{text:.2f}%', textposition='outside')
    fig_plotly.update_layout(
        width=1500,
        height=900,
        margin=dict(l=50, r=50, t=80, b=100),
        font=dict(size=12),
        xaxis_tickangle=45
    )

    # Matplotlib bar chart for top 15 combinations
    top_combinations = summary[summary['positive'] > 0].copy()
    top_combinations['Dam × Sire'] = top_combinations['Dam'] + ' × ' + top_combinations['Sire id ']
    top_combinations = top_combinations.sort_values(by='positive', ascending=False).head(15)

    fig_top15, ax = plt.subplots(figsize=(12, 6))
    bars = ax.bar(top_combinations['Dam × Sire'], top_combinations['positive'], color='mediumseagreen')

    for bar in bars:
        height = bar.get_height()
        ax.text(bar.get_x() + bar.get_width()/2, height + 0.5, f'{int(height)}', ha='center', va='bottom', fontsize=9)

    ax.set_xticklabels(top_combinations['Dam × Sire'], rotation=45, ha='right')
    ax.set_xlabel('Dam × Sire Combination')
    ax.set_ylabel('Positive Pregnancies')
    ax.set_title('Top 15 Dam–Sire Combinations by Positive Pregnancy Outcome')
    ax.grid(axis='y', linestyle='--', alpha=0.7')
    plt.tight_layout()

    return {
        # Include previous features in return as before...
        "dam_sire_summary": summary,
        "dam_sire_plotly": fig_plotly,
        "dam_sire_top15": fig_top15
    }

###################### Streamlit UI ######################
os.makedirs("uploads", exist_ok=True)
st.set_page_config(page_title="Excel ML Analyzer", layout="wide")
st.title("📊 Excel File Analyzer with ML")

uploaded_file = st.file_uploader("Upload your Excel file", type=["xlsx"])
if uploaded_file is not None:
    with st.spinner("Processing..."):
        filepath = os.path.join("uploads", uploaded_file.name)
        with open(filepath, "wb") as f:
            f.write(uploaded_file.getbuffer())

        result = process_excel_file(filepath)

        # Existing 5 features display code stays here...

        # Feature 6: Dam and Sire ID
        st.subheader("📄 Pregnancy Report by Dam and Sire ID")
        st.dataframe(result["dam_sire_summary"])

        st.subheader("📈 Interactive Chart: Positive Pregnancies by Dam and Sire ID")
        st.plotly_chart(result["dam_sire_plotly"], use_container_width=True)

        st.subheader("📈 Top 15 Dam–Sire Combinations by Positive Pregnancies")
        st.pyplot(result["dam_sire_top15"])
