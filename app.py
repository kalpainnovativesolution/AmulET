# app.py
import streamlit as st
import pandas as pd
import matplotlib.pyplot as plt
import os

def process_excel_file(filepath: str) -> dict:
    """
    Processes an Excel file to analyze embryo transfer data.
    Returns a dictionary containing dataframes and matplotlib figures.
    """
    df = pd.read_excel(filepath, sheet_name='Sheet2', skiprows=1)

    ###################### Feature 1: Pregnancy Report by Year & Quarter ######################

    pregnancy_data = df[['Date of Embryo Transfer', 'Pregnacy report']].copy()

    pregnancy_data['Date of Embryo Transfer'] = pd.to_datetime(
        pregnancy_data['Date of Embryo Transfer'], errors='coerce'
    )
    pregnancy_data.dropna(subset=['Date of Embryo Transfer'], inplace=True)

    pregnancy_data['Year'] = pregnancy_data['Date of Embryo Transfer'].dt.year
    pregnancy_data['Quarter'] = pregnancy_data['Date of Embryo Transfer'].dt.to_period('Q').dt.quarter

    positive_df = pregnancy_data[
        pregnancy_data['Pregnacy report'].str.strip().str.lower() == 'positive'
    ].copy()

    total_et_counts = pregnancy_data.groupby(['Year', 'Quarter']).size().unstack(fill_value=0)
    positive_et_counts = positive_df.groupby(['Year', 'Quarter']).size().unstack(fill_value=0)

    summary_dfs = {}
    figs = {}

    for year in sorted(total_et_counts.index):
        summary_df = pd.DataFrame({
            'Total ETs': total_et_counts.loc[year],
            'Positive Pregnancies': positive_et_counts.loc[year] if year in positive_et_counts.index else 0
        }).fillna(0)

        summary_dfs[year] = summary_df

        fig, ax = plt.subplots(figsize=(8, 5))
        bar_container = summary_df.plot(kind='bar', ax=ax)

        for container in bar_container.containers:
            ax.bar_label(container, fmt='%d', padding=3)

        ax.set_title(f'Total vs Positive Pregnancies - {year}')
        ax.set_ylabel('Count')
        ax.set_xlabel('Quarter')
        plt.xticks(rotation=0)
        plt.tight_layout()

        figs[year] = fig

    ###################### Feature 2: Pregnancy Report by Site & Grade of CL ######################

    pd_data_CL = df[['Site of CL & grade', 'Pregnacy report']].copy()
    pd_data_CL['Site of CL & grade'] = pd_data_CL['Site of CL & grade'].str.strip()
    pd_data_CL['Pregnacy report'] = pd_data_CL['Pregnacy report'].str.strip().str.lower()

    site_summary = pd_data_CL.pivot_table(index='Site of CL & grade',
                                          columns='Pregnacy report',
                                          aggfunc='size',
                                          fill_value=0)

    site_summary['Total ETs'] = site_summary.sum(axis=1)
    if 'positive' in site_summary.columns:
        site_summary['Positive %'] = (site_summary['positive'] / site_summary['Total ETs'] * 100).round(2)
    else:
        site_summary['Positive %'] = 0

    ordered_cols = ['Total ETs'] + sorted([col for col in site_summary.columns if col not in ['Total ETs', 'Positive %']]) + ['Positive %']
    site_summary = site_summary[ordered_cols]

    fig2, ax2 = plt.subplots(figsize=(12, 6))
    bar_width = 0.4
    x = range(len(site_summary))

    bars_total = ax2.bar(x, site_summary['Total ETs'], width=bar_width, label='Total ETs', color='cornflowerblue')

    if 'positive' in site_summary.columns:
        bars_positive = ax2.bar([i + bar_width for i in x], site_summary['positive'], width=bar_width, label='Positive Pregnancies', color='mediumseagreen')

        for bar in bars_positive:
            height = bar.get_height()
            ax2.text(bar.get_x() + bar.get_width()/2, height + 1, f'{int(height)}', ha='center', va='bottom', fontsize=9)

    for bar in bars_total:
        height = bar.get_height()
        ax2.text(bar.get_x() + bar.get_width()/2, height + 1, f'{int(height)}', ha='center', va='bottom', fontsize=9)

    ax2.set_xticks([i + bar_width/2 for i in x])
    ax2.set_xticklabels(site_summary.index, rotation=45, ha='right')
    ax2.set_xlabel('Site of CL & Grade')
    ax2.set_ylabel('Count')
    ax2.set_title('Pregnancy Report by Site of CL & Grade')
    ax2.legend()
    ax2.grid(axis='y', linestyle='--', alpha=0.6)
    plt.tight_layout()

    return {
        "tables": {
            "Pregnancy Data": pregnancy_data,
            "Pregnancy Report by Site & Grade of CL": site_summary
        } | {f"Summary {year}": df for year, df in summary_dfs.items()},
        "graphs": figs,
        "site_cl_graph": fig2
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

        # Display Feature 1: Tables & Graphs Year-wise
        for name, df in result.get("tables", {}).items():
            st.subheader(f"📄 Table: {name}")
            st.dataframe(df)

        for year, fig in result.get("graphs", {}).items():
            st.subheader(f"📈 Graph: {year}")
            st.pyplot(fig)

        # Display Feature 2: Site & Grade of CL Table & Graph
        st.subheader("📄 Pregnancy Report by Site & Grade of CL")
        st.dataframe(result["tables"]["Pregnancy Report by Site & Grade of CL"])

        st.subheader("📈 Pregnancy Report by Site & Grade of CL")
        st.pyplot(result["site_cl_graph"])
