# app.py
import streamlit as st
import pandas as pd
import matplotlib.pyplot as plt
import seaborn as sns
import os

def process_excel_file(filepath: str) -> dict:
    df = pd.read_excel(filepath, sheet_name='Sheet2', skiprows=1)

    ###################### Feature 1: Year & Quarter ######################
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

    # Year-Quarter counts
    total_et_counts = pregnancy_data.groupby(['Year', 'Quarter']).size().unstack(fill_value=0)
    positive_et_counts = positive_df.groupby(['Year', 'Quarter']).size().unstack(fill_value=0)

    summary_dfs = {}
    figs = {}

    # ---------- Year-wise ----------
    for year in sorted(total_et_counts.index):
        summary_df = pd.DataFrame({
            'Total ETs': total_et_counts.loc[year],
            'Positive Pregnancies': (
                positive_et_counts.loc[year] if year in positive_et_counts.index else 0
            )
        }).fillna(0)

        summary_df['Positive %'] = (
            (summary_df['Positive Pregnancies'] / summary_df['Total ETs'] * 100)
            .replace([float("inf"), -float("inf")], 0)
            .round(2)
        )

        summary_dfs[year] = summary_df

        # Plot
        fig, ax = plt.subplots(figsize=(8, 5))
        bar_container = summary_df[['Total ETs', 'Positive Pregnancies']].plot(
            kind='bar', ax=ax
        )
        for container in bar_container.containers:
            ax.bar_label(container, fmt='%d', padding=3)

        # Add percentage labels above Positive Pregnancies bars
        for i, val in enumerate(summary_df['Positive %']):
            ax.text(
                i + 0.25,
                summary_df['Positive Pregnancies'].iloc[i] + 1,
                f"{val}%",
                ha='center', va='bottom', fontsize=9, color='darkgreen'
            )

        ax.set_title(f'Total vs Positive Pregnancies - {year}')
        ax.set_ylabel('Count')
        ax.set_xlabel('Quarter')
        plt.xticks(rotation=0)
        plt.tight_layout()
        figs[year] = fig

    # ---------- Combined (All Years) ----------
    combined_total = pregnancy_data.groupby('Quarter').size()
    combined_positive = positive_df.groupby('Quarter').size()

    combined_df = pd.DataFrame({
        'Total ETs': combined_total,
        'Positive Pregnancies': combined_positive
    }).fillna(0)

    combined_df['Positive %'] = (
        (combined_df['Positive Pregnancies'] / combined_df['Total ETs'] * 100)
        .replace([float("inf"), -float("inf")], 0)
        .round(2)
    )

    summary_dfs['All Years'] = combined_df

    fig_combined, ax_combined = plt.subplots(figsize=(8, 5))
    bar_container = combined_df[['Total ETs', 'Positive Pregnancies']].plot(kind='bar', ax=ax_combined)
    for container in bar_container.containers:
        ax_combined.bar_label(container, fmt='%d', padding=3)

    for i, val in enumerate(combined_df['Positive %']):
        ax_combined.text(
            i + 0.25,
            combined_df['Positive Pregnancies'].iloc[i] + 1,
            f"{val}%",
            ha='center', va='bottom', fontsize=9, color='darkgreen'
        )

    ax_combined.set_title('Total vs Positive Pregnancies - All Years')
    ax_combined.set_ylabel('Count')
    ax_combined.set_xlabel('Quarter')
    plt.xticks(rotation=0)
    plt.tight_layout()

    figs['All Years'] = fig_combined

    ###################### Feature 2: Site & Grade of CL ######################
    pd_data_CL = df[['Site of CL & grade', 'Pregnacy report']].copy()
    pd_data_CL['Site of CL & grade'] = pd_data_CL['Site of CL & grade'].str.strip()
    pd_data_CL['Pregnacy report'] = pd_data_CL['Pregnacy report'].str.strip().str.lower()
    site_summary = pd_data_CL.pivot_table(index='Site of CL & grade', columns='Pregnacy report', aggfunc='size', fill_value=0)
    site_summary['Total ETs'] = site_summary.sum(axis=1)
    site_summary['Positive %'] = (site_summary['positive'] / site_summary['Total ETs'] * 100).round(2) if 'positive' in site_summary.columns else 0
    ordered_cols = ['Total ETs'] + sorted([col for col in site_summary.columns if col not in ['Total ETs', 'Positive %']]) + ['Positive %']
    site_summary = site_summary[ordered_cols]
    fig2, ax2 = plt.subplots(figsize=(12, 6))
    bar_width = 0.4
    x = range(len(site_summary))
    bars_total = ax2.bar(x, site_summary['Total ETs'], width=bar_width, label='Total ETs', color='cornflowerblue')
    if 'positive' in site_summary.columns:
        bars_positive = ax2.bar([i + bar_width for i in x], site_summary['positive'], width=bar_width, label='Positive Pregnancies', color='mediumseagreen')
        for bar in bars_positive:
            ax2.text(bar.get_x() + bar.get_width()/2, bar.get_height() + 1, f'{int(bar.get_height())}', ha='center', va='bottom', fontsize=9)
    for bar in bars_total:
        ax2.text(bar.get_x() + bar.get_width()/2, bar.get_height() + 1, f'{int(bar.get_height())}', ha='center', va='bottom', fontsize=9)
    ax2.set_xticks([i + bar_width/2 for i in x])
    ax2.set_xticklabels(site_summary.index, rotation=45, ha='right')
    ax2.set_xlabel('Site of CL & Grade')
    ax2.set_ylabel('Count')
    ax2.set_title('Pregnancy Report by Site & Grade of CL')
    ax2.legend()
    ax2.grid(axis='y', linestyle='--', alpha=0.6)
    plt.tight_layout()

    ###################### Feature 3: Stage & Age of Embryo ######################
    embryo_data = df[['Stage of embryo', 'Age of embryo', 'Pregnacy report']].copy()
    embryo_data['Stage of embryo'] = embryo_data['Stage of embryo'].str.strip()
    embryo_data['Age of embryo'] = embryo_data['Age of embryo'].astype(str).str.strip()
    embryo_data['Pregnacy report'] = embryo_data['Pregnacy report'].str.strip().str.lower()
    positive_embryo_data = embryo_data[embryo_data['Pregnacy report'] == 'positive']
    summary = positive_embryo_data.groupby(['Stage of embryo', 'Age of embryo']).size().reset_index(name='Positive Count')
    pivot_table = summary.pivot(index='Stage of embryo', columns='Age of embryo', values='Positive Count').fillna(0)
    fig3, ax3 = plt.subplots(figsize=(10, 6))
    sns.heatmap(pivot_table, annot=True, fmt=".0f", cmap="cividis", ax=ax3)
    ax3.set_title("Positive Pregnancies by Embryo Stage and Age")
    ax3.set_xlabel("Age of Embryo")
    ax3.set_ylabel("Stage of Embryo")
    plt.tight_layout()

    ###################### Feature 4: Type of Semen Used ######################
    semen_data = df[['Conventional/ Sexed semen', 'Pregnacy report']].copy()
    semen_data['Conventional/ Sexed semen'] = semen_data['Conventional/ Sexed semen'].str.strip().str.title()
    semen_data['Pregnacy report'] = semen_data['Pregnacy report'].str.strip().str.lower()
    semen_summary = semen_data.pivot_table(index='Conventional/ Sexed semen', columns='Pregnacy report', aggfunc='size', fill_value=0)
    semen_summary['Total ETs'] = semen_summary.sum(axis=1)
    semen_summary['Positive %'] = (semen_summary['positive'] / semen_summary['Total ETs'] * 100).round(2) if 'positive' in semen_summary.columns else 0
    cols = ['Total ETs'] + sorted([col for col in semen_summary.columns if col not in ['Total ETs', 'Positive %']]) + ['Positive %']
    semen_summary = semen_summary[cols]
    fig4, ax4 = plt.subplots(figsize=(8, 5))
    bar_width = 0.35
    x = range(len(semen_summary))
    bars_total = ax4.bar(x, semen_summary['Total ETs'], width=bar_width, label='Total ETs', color='skyblue')
    if 'positive' in semen_summary.columns:
        bars_positive = ax4.bar([i + bar_width for i in x], semen_summary['positive'], width=bar_width, label='Positive Pregnancies', color='mediumseagreen')
        for bar in bars_positive:
            ax4.text(bar.get_x() + bar.get_width()/2, bar.get_height() + 1, f'{int(bar.get_height())}', ha='center', va='bottom', fontsize=9)
    for bar in bars_total:
        ax4.text(bar.get_x() + bar.get_width()/2, bar.get_height() + 1, f'{int(bar.get_height())}', ha='center', va='bottom', fontsize=9)
    ax4.set_xticks([i + bar_width/2 for i in x])
    ax4.set_xticklabels(semen_summary.index, rotation=0)
    ax4.set_xlabel('Semen Type')
    ax4.set_ylabel('Count')
    ax4.set_title('Pregnancy Report by Semen Type')
    ax4.legend()
    ax4.grid(axis='y', linestyle='--', alpha=0.6)
    plt.tight_layout()

    ###################### Feature 5: Agency/Organization Wise ######################
    org_data = df[['Organization', 'Pregnacy report']].copy()
    org_data['Organization'] = org_data['Organization'].str.strip().str.title()
    org_data['Pregnacy report'] = org_data['Pregnacy report'].str.strip().str.lower()
    org_summary = org_data.pivot_table(index='Organization', columns='Pregnacy report', aggfunc='size', fill_value=0)
    org_summary['Total ETs'] = org_summary.sum(axis=1)
    org_summary['Positive %'] = (org_summary['positive'] / org_summary['Total ETs'] * 100).round(2) if 'positive' in org_summary.columns else 0
    cols = ['Total ETs'] + sorted([col for col in org_summary.columns if col not in ['Total ETs', 'Positive %']]) + ['Positive %']
    org_summary = org_summary[cols]
    fig5, ax5 = plt.subplots(figsize=(12, 6))
    bar_width = 0.35
    x = range(len(org_summary))
    bars_total = ax5.bar(x, org_summary['Total ETs'], width=bar_width, label='Total ETs', color='cornflowerblue')
    if 'positive' in org_summary.columns:
        bars_positive = ax5.bar([i + bar_width for i in x], org_summary['positive'], width=bar_width, label='Positive Pregnancies', color='mediumseagreen')
        for bar in bars_positive:
            ax5.text(bar.get_x() + bar.get_width()/2, bar.get_height() + 1, f'{int(bar.get_height())}', ha='center', va='bottom', fontsize=9)
    for bar in bars_total:
        ax5.text(bar.get_x() + bar.get_width()/2, bar.get_height() + 1, f'{int(bar.get_height())}', ha='center', va='bottom', fontsize=9)
    ax5.set_xticks([i + bar_width/2 for i in x])
    ax5.set_xticklabels(org_summary.index, rotation=45, ha='right')
    ax5.set_xlabel('Organization')
    ax5.set_ylabel('Count')
    ax5.set_title('Pregnancy Report by Organization')
    ax5.legend()
    ax5.grid(axis='y', linestyle='--', alpha=0.6)
    plt.tight_layout()

    return {
        "tables": {
            "Pregnancy Data": pregnancy_data,
            "Pregnancy Report by Site & Grade of CL": site_summary,
            "Pregnancy Report by Stage & Age of Embryo": pivot_table,
            "Pregnancy Report by Semen Type": semen_summary,
            "Pregnancy Report by Organization": org_summary
        } | {f"Summary {year}": df for year, df in summary_dfs.items()},
        "graphs": figs,
        "site_cl_graph": fig2,
        "embryo_stage_graph": fig3,
        "semen_graph": fig4,
        "org_graph": fig5
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
        for name, df in result.get("tables", {}).items():
            if name.startswith("Summary"):
                st.subheader(f"📄 Table: {name}")
                st.dataframe(df)
        for year, fig in result.get("graphs", {}).items():
            st.subheader(f"📈 Graph: {year}")
            st.pyplot(fig)
        st.subheader("📄 Pregnancy Report by Site & Grade of CL")
        st.dataframe(result["tables"]["Pregnancy Report by Site & Grade of CL"])
        st.pyplot(result["site_cl_graph"])
        st.subheader("📄 Positive Pregnancies by Stage & Age of Embryo")
        st.dataframe(result["tables"]["Pregnancy Report by Stage & Age of Embryo"])
        st.pyplot(result["embryo_stage_graph"])
        st.subheader("📄 Pregnancy Report by Semen Type")
        st.dataframe(result["tables"]["Pregnancy Report by Semen Type"])
        st.pyplot(result["semen_graph"])
        st.subheader("📄 Pregnancy Report by Organization")
        st.dataframe(result["tables"]["Pregnancy Report by Organization"])
        st.pyplot(result["org_graph"])
