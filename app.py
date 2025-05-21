import streamlit as st
import pandas as pd
from datetime import datetime

st.set_page_config(page_title="KPI Processor", layout="wide")

@st.cache_data
def load_excel(file):
    return pd.read_excel(file, engine='openpyxl')

@st.cache_data
def convert_df_to_csv(df):
    return df.to_csv(index=False).encode('utf-8')

st.title("📊 KPI Data Processor")

file = st.file_uploader("Upload your Excel file (max 80 MB)", type=["xlsx"])

if file:
    with st.spinner("Reading and processing file..."):
        df_raw = load_excel(file)

        # Keep only required columns
        needed_columns = [
            'Ticket Number Inap', 'Ticket Number SWFM', 'Severity', 'Type Ticket', 'Site Id',
            'Site Name', 'Site Class', 'Cluster TO', 'Occured Time', 'Created At',
            'Ticket Inap Status', 'Ticket SWFM Status', 'PIC Take Over Ticket', 'NOP',
            'Regional', 'Area', 'Cleared Time', 'Take Over Date', 'Check In At',
            'SLA Status', 'Fault Level', 'Incident Priority', 'Hub', 'Is Excluded In KPI',
            'Site Cleared On', 'Rank', 'RCA Validated'
        ]
        df = df_raw[needed_columns].copy()

        # Remove excluded rows
        df = df[df['Is Excluded In KPI'].str.upper() != 'YES']

        # Map OldKPI
        fault_mapping = {
            'Controller P2': 'B22', 'Controller P1': 'B22',
            'Enva Controller': 'B21', 'Enva Site': 'B21', 'Enva Site GSB': 'B21',
            'Enva Site Simpul': 'B21', 'Enva Site VIP': 'B21',
            'L2 Configuration': 'B3', 'P1': 'B23', 'P1 VIP': 'B23',
            'P2': 'B3', 'P2 VIP': 'B3', 'Vandalism': 'B3',
            'L2 License': 'B3', 'P3': 'B3'
        }
        df['OldKPI'] = df['Fault Level'].map(fault_mapping)
        df.loc[df['Type Ticket'] == 'Incident', 'OldKPI'] = 'B1'

        # Datetime parsing
        df['Occured Time'] = pd.to_datetime(df['Occured Time'], errors='coerce')
        df['Cleared Time'] = pd.to_datetime(df['Cleared Time'], errors='coerce')
        df['Site Cleared On'] = pd.to_datetime(df['Site Cleared On'], errors='coerce')

        # MTTR calculation
        df['MTTR'] = (
            df['Site Cleared On'].fillna(df['Cleared Time']) - df['Occured Time']
        ).dt.total_seconds() / 3600
        df['MTTR'] = df['MTTR'].fillna(0)

        # SLAMTTR
        slamap = {
            ('Incident', 'Critical'): 4, ('Incident', 'Major'): 8,
            ('Incident', 'Minor'): 10, ('Incident', 'Low'): 13,
            ('Event', 'Critical'): 2, ('Event', 'Major'): 4,
            ('Event', 'Minor'): 15, ('Event', 'Low'): 48
        }
        df['SLAMTTR'] = df.apply(lambda row: slamap.get((row['Type Ticket'], row['Severity']), 0), axis=1)

        # ScoreMTTR
        df['ScoreMTTR'] = df['MTTR'] / df['SLAMTTR']
        df['ScoreMTTR'] = df['ScoreMTTR'].apply(lambda x: min(x, 1.5) if x > 0 else 0)

        # Handling and ScoreTO
        df['Handling'] = df['PIC Take Over Ticket'].apply(lambda x: 'Autoclear' if pd.isna(x) else 'Takeover')
        df['ScoreTO'] = df['Handling'].map({'Takeover': 0.3, 'Autoclear': 0})

        # Remove rows with Handling == 'Autoclear' and MTTR < 1
        df = df[~((df['Handling'] == 'Autoclear') & (df['MTTR'] < 1))]

        # Visitation and ScoreVisit
        df['Visitation'] = df['Check In At'].apply(lambda x: 'Visit' if pd.notna(x) else 'NoVisit')
        df['ScoreVisit'] = df['Visitation'].map({'Visit': 0.5, 'NoVisit': 0})

        # ScoreRCA
        df['ScoreRCA'] = df['RCA Validated'].map({'Yes': 0.1, 'No': 0}).fillna(0)

        # ScoreClosed
        df['ScoreClosed'] = df['Ticket SWFM Status'].apply(lambda x: 0.1 if x == 'Closed' else 0)

        # Final SCORE
        df['SCORE'] = (df['ScoreTO'] + df['ScoreVisit'] + df['ScoreRCA'] + df['ScoreClosed']) * df['ScoreMTTR']

        # Summary
        summary = df.groupby(
            ['Area', 'Regional', 'NOP', 'OldKPI', 'Severity']
        ).agg(
            Ticket_Count=('Ticket Number Inap', 'count'),
            Avg_SCORE=('SCORE', 'mean'),
            Sum_SCORE=('SCORE', 'sum'),
        ).reset_index()

    # Display results
    st.subheader("📊 Summary Table")
    st.dataframe(summary)

    # Download summary
    st.download_button(
        label="📥 Download Summary CSV",
        data=convert_df_to_csv(summary),
        file_name="summary.csv",
        mime="text/csv"
    )

    # Download processed data
    csv = df.to_csv(index=False).encode('utf-8')
    st.download_button(
        label="📥 Download Processed Data as CSV",
        data=csv,
        file_name='processed_kpi_data.csv',
        mime='text/csv'
    )

    # Show full processed data
    with st.expander("🔍 View Full Processed Data"):
        st.dataframe(df)

    
