import streamlit as st
import pandas as pd
from datetime import datetime

st.set_page_config(page_title="Email Outreach Report", layout="wide")
st.title("📧 Email Outreach Report Tool")

uploaded_file = st.file_uploader("Upload Outreach Excel File", type=["xlsx"])


# Normalize function
def normalize(col):
    return col.strip().lower().replace(" ", "").replace("_", "")


# Expected normalized headers
expected_normalized = {
    'website': 'Website',
    'email': 'Email',
    'firstemaildate': 'First Email Date',
    'reminder1': 'Reminder1',
    'reminder2': 'Reminder2',
    'reminder3': 'Reminder 3'
}


# Parse date columns
def parse_dates(df, date_cols):
    for col in date_cols:
        df[col] = pd.to_datetime(df[col], errors='coerce')
    return df


if uploaded_file:
    # Get sheet names
    all_sheets = pd.ExcelFile(uploaded_file).sheet_names
    selected_sheet = st.selectbox("📄 Select Sheet to Analyze", all_sheets)

    # Load selected sheet
    df_raw = pd.read_excel(uploaded_file, sheet_name=selected_sheet)

    # Normalize and map headers
    normalized_map = {normalize(col): col for col in df_raw.columns}
    missing = [v for k, v in expected_normalized.items() if k not in normalized_map]

    if missing:
        st.error(f"Missing required columns in sheet '{selected_sheet}': {', '.join(missing)}")
    else:
        rename_dict = {normalized_map[k]: v for k, v in expected_normalized.items()}
        df = df_raw.rename(columns=rename_dict)

        # Parse dates
        df = parse_dates(df, ['First Email Date', 'Reminder1', 'Reminder2', 'Reminder 3'])


        # Last interaction logic
        def get_last_reminder(row):
            for col in ['Reminder 3', 'Reminder2', 'Reminder1']:
                if pd.notnull(row[col]):
                    return col, row[col]
            return ("First Email Date", row['First Email Date'])


        df[['Last Interaction Type', 'Last Interaction Date']] = df.apply(
            lambda row: pd.Series(get_last_reminder(row)), axis=1
        )

        df['Total Reminders Sent'] = df[['Reminder1', 'Reminder2', 'Reminder 3']].notnull().sum(axis=1)
        no_reminders_df = df[df['Total Reminders Sent'] == 0]
        latest_activity = df['Last Interaction Date'].max()

        # 📅 Date filter
        st.subheader("📅 Filter by Specific Date")
        selected_date = st.date_input("Select a date to filter outreach activity")

        if selected_date:
            filtered_df = df[
                (df['First Email Date'].dt.date == selected_date) |
                (df['Reminder1'].dt.date == selected_date) |
                (df['Reminder2'].dt.date == selected_date) |
                (df['Reminder 3'].dt.date == selected_date) |
                (df['Last Interaction Date'].dt.date == selected_date)
                ]

            st.markdown(f"### 🔍 Contacts Interacted on {selected_date}")
            if filtered_df.empty:
                st.info("No contacts found with activity on this date.")
            else:
                st.dataframe(
                    filtered_df[['Website', 'Email', 'First Email Date', 'Reminder1', 'Reminder2', 'Reminder 3',
                                 'Last Interaction Type', 'Last Interaction Date', 'Total Reminders Sent']])

        # 📌 Filter by Reminder
        st.subheader("📌 Filter by Specific Reminder Sent")
        selected_reminder = st.selectbox("Select Reminder Type",
                                         ['First Email Date', 'Reminder1', 'Reminder2', 'Reminder 3'])
        reminder_df = df[pd.notnull(df[selected_reminder])]

        st.markdown(f"### 📤 Contacts with {selected_reminder} Sent")
        if reminder_df.empty:
            st.info(f"No records found with {selected_reminder} sent.")
        else:
            st.dataframe(reminder_df[['Website', 'Email', 'First Email Date', 'Reminder1', 'Reminder2', 'Reminder 3',
                                      'Last Interaction Type', 'Last Interaction Date', 'Total Reminders Sent']])

        # 📊 Summary
        st.subheader("📊 Outreach Summary")
        st.markdown(f"**Sheet:** `{selected_sheet}`")
        st.markdown(f"**Total Contacts:** {len(df)}")
        st.markdown(f"**Latest Outreach Date:** {latest_activity.date() if pd.notnull(latest_activity) else 'N/A'}")
        st.markdown(f"**Contacts With No Reminders:** {len(no_reminders_df)}")

        # 📋 Full Report
        st.subheader("📋 Full Outreach Report")
        st.dataframe(df[['Website', 'Email', 'First Email Date', 'Reminder1', 'Reminder2', 'Reminder 3',
                         'Last Interaction Type', 'Last Interaction Date', 'Total Reminders Sent']])

        # ⚠️ Contacts without reminders
        st.subheader("⚠️ Contacts Without Reminders")
        st.dataframe(no_reminders_df[['Website', 'Email', 'First Email Date']])
else:
    st.info("Please upload your Excel file to get started.")
