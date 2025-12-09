
import pandas as pd
import numpy as np
import streamlit as st
from io import BytesIO
import math
import altair as alt
import re
from tqdm import tqdm
import plotly.express as px
from datetime import datetime, timedelta

# Set the page title
st.set_page_config(page_title="Audit Applications")

# Inject custom CSS (simplified to avoid HTML escaping issues on Cloud)
st.markdown(
    """
    <style>
    div[data-baseweb="select"] > div {font-size:16px;background-color:Tomato;color:black;}
    input[type="text"] {border:3px solid black;border-radius:4px;padding:8px;font-size:16px;}
    .stDownloadButton {padding:8px 16px;font-size:16px;font-weight:bold;background-color:white;color:black;}
    .stTextInput input, div[data-baseweb="select"] > div {font-size:16px;}
    </style>
    """,
    unsafe_allow_html=True,
)

def load_data(uploaded_file):
    """
    Load data from an uploaded Excel file.
    """
    try:
        # Explicit engine avoids Cloud defaults issues & xlrd/xlsx mismatch
        df = pd.read_excel(uploaded_file, header=4, engine='openpyxl')

        # Drop only if present; schema can vary
        if 'Unnamed: 0' in df.columns:
            df = df.drop(['Unnamed: 0'], axis=1)

        # Validate required columns to avoid downstream KeyErrors
        required_cols = [
            'Journal ID', 'Debit', 'Credit', 'Transaction Date',
            'Posted Date', 'User', 'Account Name'
        ]
        missing = [c for c in required_cols if c not in df.columns]
        if missing:
            st.error(f"Missing required columns in the uploaded file: {missing}")
            return None

        # Coerce dates; warn if any invalid
        df['Transaction Date'] = pd.to_datetime(df['Transaction Date'], errors='coerce')
        df['Posted Date'] = pd.to_datetime(df['Posted Date'], errors='coerce')
        if df['Transaction Date'].isna().any() or df['Posted Date'].isna().any():
            st.warning("Some rows have invalid dates and were coerced to NaT. "
                       "These rows may be excluded from date-based calculations.")
        return df
    except Exception as e:
        st.error(f"Error loading file: {e}")
        return None

def clean_user_column(df):
    """Clean the 'User' column."""
    if 'User' not in df.columns:
        st.error("Column 'User' not found in the uploaded file.")
        return df
    df['User'] = (
        df['User'].astype(str)
        .str.lower()
        .str.strip()
        .str.replace(r'\s+', ' ', regex=True)
        .apply(lambda x: re.sub(r'\W+', ' ', x))
        .str.title()
    )
    return df

def process_manual_journals(df):
    """Process the manual journal data."""

    # Keep only valid rows for grouping
    needed = ['Journal ID', 'Debit', 'Transaction Date', 'Posted Date', 'User']
    valid_df = df.dropna(subset=needed)

    # Group by Journal ID
    grouped_journal_id = valid_df.groupby('Journal ID', as_index=False).agg({
        'Debit': 'sum',
        'Transaction Date': 'first',
        'User': 'first',
        'Posted Date': 'first',
    }).rename(columns={'Debit': 'Journal Value'})

    # Date parts
    grouped_journal_id['day'] = grouped_journal_id['Posted Date'].dt.day
    grouped_journal_id['day_name'] = grouped_journal_id['Posted Date'].dt.day_name()
    grouped_journal_id['week_number'] = grouped_journal_id['Posted Date'].dt.isocalendar().week
    grouped_journal_id['month'] = grouped_journal_id['Posted Date'].dt.month
    grouped_journal_id['month_name'] = grouped_journal_id['Posted Date'].dt.month_name()

    # Differences
    grouped_journal_id['days_between_post_transaction'] = (
        grouped_journal_id['Posted Date'] - grouped_journal_id['Transaction Date']
    ).dt.days

    # For display
    grouped_journal_id['Posted Date_formatted'] = grouped_journal_id['Posted Date'].dt.strftime('%d/%m/%Y')
    grouped_journal_id['Transaction Date_formatted'] = grouped_journal_id['Transaction Date'].dt.strftime('%d/%m/%Y')

    # By month
    grouped_month = grouped_journal_id.groupby('month_name', as_index=False).agg({
        'Journal ID': 'count',
        'Journal Value': 'sum'
    }).rename(columns={
        'month_name': 'Manual journals by month',
        'Journal ID': '01 July 2022 - 30 June 2023',
        'Journal Value': 'Sum of Journal Value'
    })

    # By user
    grouped_user = grouped_journal_id.groupby('User', as_index=False).agg({
        'Journal ID': 'count',
        'Journal Value': 'sum'
    }).rename(columns={
        'User': 'Manual journals by user',
        'Journal ID': '01 July 2022 - 30 June 2023',
        'Journal Value': 'Sum of Journal Value'
    }).sort_values(by='Manual journals by user').reset_index(drop=True)

    # Difference bins
    bins = [0, 7, 14, 21, 28, 35, float('inf')]
    labels = ['0 - 7 days', '7 - 14 days', '14 - 21 days', '21 - 28 days', '28 - 35 days', '> 35 days']
    grouped_journal_id['days_difference_group'] = pd.cut(
        grouped_journal_id['days_between_post_transaction'],
        bins=bins, labels=labels, right=False
    )
    grouped_difference_days = grouped_journal_id.groupby('days_difference_group', observed=True).agg(
        Count_of_trx_01_July_2022__30_June_2023=('days_between_post_transaction', 'size'),
        value_of_trx=('Journal Value', 'sum')
    ).reset_index()

    # Filters
    negative_days_grouped_journal_id = grouped_journal_id[grouped_journal_id['days_between_post_transaction'] < 0]
    sunday_days_df = grouped_journal_id[grouped_journal_id['day_name'] == 'Sunday']
    saturday_days_df = grouped_journal_id[grouped_journal_id['day_name'] == 'Saturday']

    return (
        grouped_journal_id, grouped_month, grouped_user, grouped_difference_days,
        negative_days_grouped_journal_id, sunday_days_df, saturday_days_df
    )

def main():
    # Sidebar
    st.sidebar.title("Audit App List")
    pages = ["Manual Journals-Xero", "Manual Journals-MYOB"]
    selection = st.sidebar.radio("Select a Module", pages)

    st.title("Audit App")

    # Page 1
    if selection == "Manual Journals-Xero":
        st.header("Manual Journals-Xero")
        st.subheader(" Upload the Manual Journals/ Downloaded from the Xero")

        # Session state
        if "data" not in st.session_state:
            st.session_state.data = None

        uploaded_file1 = st.file_uploader("Choose an Excel file", type=["xlsx"], key="file_uploader_1")
        if uploaded_file1 is not None:
            st.session_state.data = load_data(uploaded_file1)
            df = st.session_state.data
            if df is not None:
                # Clean + sample
                df = clean_user_column(df)
                st.write("Number of rows present in the file: " + str(len(df)))
                st.write("Below is the sample data of the uploaded file")
                st.write(df.head(2))

                # Process
                (
                    grouped_journal_id, grouped_month, sorted_grouped_user, grouped_difference_days,
                    negative_days_grouped_journal_id, sunday_days_df, saturday_days_df
                ) = process_manual_journals(df)

                # Display summary
                st.write("# Manual Journal Data")
                st.write("Number of Journals: " + str(len(grouped_journal_id)))
                st.write("Debit Value: " + str(round(df['Debit'].sum(), 2)))
                st.write("Credit Value: " + str(round(df['Credit'].sum(), 2)))
                st.write("Earliest_date: " + str(grouped_journal_id['Transaction Date'].min()))
                st.write("Latest_date: " + str(grouped_journal_id['Transaction Date'].max()))

                # By month
                total_journal_id = grouped_month['01 July 2022 - 30 June 2023'].sum()
                total_journal_value_sum = round(grouped_month['Sum of Journal Value'].sum(), 2)

                st.write("## Displaying Manual Journals posted by month")
                st.write(grouped_month)
                st.write("Total Journal IDs: " + str(total_journal_id))
                st.write("Total Sum of Journal Values: " + str(total_journal_value_sum))

                # Plotly line chart
                fig = px.line(grouped_month, x='Manual journals by month', y='Sum of Journal Value', title='Line Graph')
                fig.update_traces(text=grouped_month['Sum of Journal Value'], textposition='top center')
                fig.update_layout(xaxis_tickangle=-45)
                st.plotly_chart(fig)

                st.write("## Assessment of Journals performed by the Users")
                st.write(sorted_grouped_user)

                st.write("## Assigning Journals into date difference bins")
                st.write(grouped_difference_days)

                st.write("### Below are the rows, where transaction date is higher than posted date")
                st.write(negative_days_grouped_journal_id)

                st.write("## Below are the journals posted on Sunday:")
                st.write(sunday_days_df)
                st.write("Number of journals posted on Sunday: " + str(len(sunday_days_df)))
                st.write("Sum of journal Values posted on Sunday: " + str(sunday_days_df['Journal Value'].sum()))

                st.write("## Below are the journals posted on Saturday:")
                st.write(saturday_days_df)
                st.write("Number of journals posted on Saturday: " + str(len(saturday_days_df)))
                st.write("Sum of journal Values posted on Saturday: " + str(saturday_days_df['Journal Value'].sum()))

                # Download workbook
                flnme = st.text_input('Enter your file name to download the report (e.g. report1.xlsx)')
                if flnme and not flnme.endswith(".xlsx"):
                    flnme += ".xlsx"

                buffer = BytesIO()
                li = [
                    grouped_journal_id, grouped_month, sorted_grouped_user,
                    grouped_difference_days, sunday_days_df, saturday_days_df
                ]
                with pd.ExcelWriter(buffer, engine='xlsxwriter') as writer:
                    for k, df1 in enumerate(li, start=1):
                        df1.to_excel(writer, sheet_name=f'Report{k}', index=False)

                # IMPORTANT: rewind + correct MIME
                buffer.seek(0)
                if flnme:
                    st.download_button(
                        label="Download Excel workbook",
                        data=buffer,
                        file_name=flnme,
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    )
                else:
                    st.warning("Please provide a file name.")

                # Per-account view
                st.write('##### Select the Account name from the below dropdown, to view/download the data related to it')
                st.write("##### Account Names")
                unique_account_names = df['Account Name'].dropna().unique()
                selected_account = st.selectbox('Select Account Name', sorted(unique_account_names))
                selected_df = df[df['Account Name'] == selected_account]
                st.write("Number of Journals with the account name :: " + f' {selected_account} ' + "are : " + str(len(selected_df)))
                st.write("Displaying 5 sample rows for the selected Account Name")
                st.write(selected_df.head(5))

    # Page 2
    elif selection == "Manual Journals-MYOB":
        st.header("Manual Journals-MYOB")
        st.subheader(" Upload the Manual Journals/ Downloaded from the Xero")
        uploaded_file2 = st.file_uploader("Choose an Excel file", type=["xlsx"], key="file_uploader_2")
        if uploaded_file2 is not None:
            df = load_data(uploaded_file2)
            if df is not None:
                df = clean_user_column(df)
                st.write("Number of rows present in the file: " + str(len(df)))
                st.write("Below is the sample data of the uploaded file")
                st.write(df.head(2))

if __name__ == "__main__":
    main()
