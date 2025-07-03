# app.py
import streamlit as st
import pandas as pd
import matplotlib.pyplot as plt
import seaborn as sns
import plotly.express as px
import os

def process_excel_file(filepath: str) -> dict:
    df = pd.read_excel(filepath, sheet_name='Sheet2', skiprows=1)

    ###################### Feature 1: Year & Quarter ######################
    pregnancy_data = df[['Date of Embryo Transfer', 'Pregnacy report']].copy()
    pregnancy_data['Date of Embryo Transfer'] = pd.to_datetime(pregnancy_data['Date of Embryo Transfer'], errors='coerce')
    pregnancy_data.dropna(subset=['Date of Embryo Transfer'], inplace=True)
    pregnancy_data['Year'] = pregnancy_data['Date of Embryo Transfer'].dt.year
    pregnancy_data['Quarter'] = pregnancy_data['Date of Embryo Transfer'].dt.to_period('Q').dt.quarter
    positive_df = pregnancy_data[pregnancy_data['Pregnacy report'].str.strip().str.lower() == 'positive'].copy()
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

    ###################### Feature 4: Semen Type ######################
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

    ###################### Feature 5: Organization ######################
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

    # ###################### Feature 6: Dam & Sire Combination ######################
    # df['Dam'] = df['Dam'].astype(str).str.strip().str.title()
    # df['Sire id '] = df['Sire id '].astype(str).str.strip().str.title()
    # df['Pregnacy report'] = df['Pregnacy report'].astype(str).str.strip().str.lower()
    # summary = df.groupby(['Dam', 'Sire id ', 'Pregnacy report']).size().unstack(fill_value=0).reset_index()
    # summary['Total ETs'] = summary.sum(axis=1, numeric_only=True)
    # summary['% Positive'] = round(summary['positive'] / summary['Total ETs'] * 100, 2) if 'positive' in summary.columns else 0.0
    # summary = summary.sort_values(by='positive', ascending=False, na_position='last')
    # fig_plotly = px.bar(summary, x='Dam', y='positive', color='Sire id ', title="Positive Pregnancy Report by Dam and Sire ID", labels={'positive': 'Number of Positive Reports'}, text='% Positive')
    # fig_plotly.update_traces(texttemplate='%{text:.2f}%', textposition='outside')
    # fig_plotly.update_layout(width=1500, height=900, margin=dict(l=50, r=50, t=80, b=100), font=dict(size=12), xaxis_tickangle=45)
    # top_combinations = summary[summary['positive'] > 0].copy()
    # top_combinations['Dam × Sire'] = top_combinations['Dam'] + ' × ' + top_combinations['Sire id ']
    # top_combinations = top_combinations.sort_values(by='positive', ascending=False).head(15)
    # fig_top15, ax = plt.subplots(figsize=(12, 6))
    # bars = ax.bar(top_combinations['Dam × Sire'], top_combinations['positive'], color='mediumseagreen')
    # for bar in bars:
    #     height = bar.get_height()
    #     ax.text(bar.get_x() + bar.get_width()/2, height + 0.5, f'{int(height)}', ha='center', va='bottom', fontsize=9)
    # ax.set_xticklabels(top_combinations['Dam × Sire'], rotation=45, ha='right')
    # ax.set_xlabel('Dam × Sire Combination')
    # ax.set_ylabel('Positive Pregnancies')
    # ax.set_title('Top 15 Dam–Sire Combinations by Positive Pregnancy Outcome')
    # ax.grid(axis='y', linestyle='--', alpha=0.7)
    # plt.tight_layout()

    # ###################### Feature 7: Dam Breed & Sire Breed Combination ######################
    # df['Dam Breed'] = df['Dam Breed'].astype(str).str.strip().str.title()
    # df['Sire Breed'] = df['Sire Breed'].astype(str).str.strip().str.title()
    # positive_data = df[df['Pregnacy report'] == 'positive']
    # grouped_positive = positive_data.groupby(['Dam Breed', 'Sire Breed']).size().reset_index(name='Positive Count')
    # total_reports = df.groupby(['Dam Breed', 'Sire Breed']).size().reset_index(name='Total Reports')
    # merged_data = pd.merge(grouped_positive, total_reports, on=['Dam Breed', 'Sire Breed'])
    # merged_data['% Positive'] = (merged_data['Positive Count'] / merged_data['Total Reports']) * 100
    # fig_breed_bar = px.bar(merged_data, x='Dam Breed', y='% Positive', color='Sire Breed', title="Positive Pregnancy Report % by Dam Breed and Sire Breed", labels={'% Positive': 'Percentage of Positive'}, text='% Positive')
    # fig_breed_bar.update_traces(texttemplate='%{text:.2f}%', textposition='outside')
    # fig_breed_bar.update_layout(width=800, height=600, margin=dict(l=50, r=50, t=80, b=50), font=dict(size=12))
    # fig_breed_heatmap = px.density_heatmap(merged_data, x='Sire Breed', y='Dam Breed', z='% Positive', text_auto='.2f', color_continuous_scale='Cividis', title='% Positive Pregnancy Report Heatmap by Dam Breed and Sire Breed', labels={'% Positive': '% Positive'})
    # fig_breed_heatmap.update_layout(width=800, height=800, margin=dict(l=60, r=60, t=80, b=60), font=dict(size=12))

    # return {
    #     "tables": summary_dfs,
    #     "graphs": figs,
    #     "site_summary": site_summary,
    #     "site_graph": fig2,
    #     "embryo_table": pivot_table,
    #     "embryo_graph": fig3,
    #     "semen_summary": semen_summary,
    #     "semen_graph": fig4,
    #     "org_summary": org_summary,
    #     "org_graph": fig5,
    #     "dam_sire_summary": summary,
    #     "dam_sire_plotly": fig_plotly,
    #     "dam_sire_top15": fig_top15,
    #     "breed_summary": merged_data,
    #     "breed_bar": fig_breed_bar,
    #     "breed_heatmap": fig_breed_heatmap
    # }
