import streamlit as st
import pandas as pd
from datetime import datetime

st.set_page_config(page_title="KPI Processor", layout="wide")

@st.cache_data
def load_excel(files):
    dfs = []
    for file in files:
        df = pd.read_excel(file, engine='openpyxl')
        dfs.append(df)
    return pd.concat(dfs, ignore_index=True)

@st.cache_data
def convert_df_to_csv(df):
    return df.to_csv(index=False).encode('utf-8')

st.title("📊 KPI Data Processor")

uploaded_files = st.file_uploader(
    "Upload one or more SWFM data files (max total 200 MB)", 
    type=["xlsx"], 
    accept_multiple_files=True
)

if uploaded_files:
    with st.spinner("Reading and processing files..."):
        df_raw = load_excel(uploaded_files)

        # [Same processing logic as before...]
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

        # Extract Month and Year from "Occured Time"
        df['Occured Time'] = pd.to_datetime(df['Occured Time'], errors='coerce')
        df['Month'] = df['Occured Time'].dt.strftime('%B')  # Full month name (e.g., March)
        df['Year'] = df['Occured Time'].dt.year

        # Mapping OldKPI
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

        # Date parsing
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
            ['Month', 'Year', 'Area', 'Regional', 'NOP', 'OldKPI', 'Severity']
        ).agg(
            Ticket_Count=('Ticket Number SWFM', 'count'),
            Avg_SCORE=('SCORE', 'mean'),
            Sum_SCORE=('SCORE', 'sum'),
        ).reset_index()

    # Get unique filter options
    month_options = sorted(df['Month'].dropna().unique())
    year_options = sorted(df['Year'].dropna().unique())
    area_options = sorted(df['Area'].dropna().unique())
    
    # Set up 5 columns for filters
    col1, col2, col3, col4, col5 = st.columns(5)
    
    # Month filter
    with col1:
        selected_month = st.selectbox("📅 Month", ['All'] + month_options)
    
    # Year filter
    with col2:
        selected_year = st.selectbox("📆 Year", ['All'] + list(map(str, year_options)))
    
    # Area filter
    with col3:
        selected_area = st.selectbox("🌍 Area", ['All'] + area_options)
    
    # Regional filter (cascading from Area)
    if selected_area != 'All':
        regional_options = sorted(df[df['Area'] == selected_area]['Regional'].dropna().unique())
    else:
        regional_options = sorted(df['Regional'].dropna().unique())
    
    with col4:
        selected_regional = st.selectbox("📍 Regional", ['All'] + regional_options)
    
    # NOP filter (cascading from Regional)
    if selected_regional != 'All':
        nop_options = sorted(df[df['Regional'] == selected_regional]['NOP'].dropna().unique())
    else:
        nop_options = sorted(df['NOP'].dropna().unique())
    
    with col5:
        selected_nop = st.selectbox("🏢 NOP", ['All'] + nop_options)
    
    # Apply filters to summary
    filtered_summary = summary.copy()
    if selected_month != 'All':
        filtered_summary = filtered_summary[filtered_summary['Month'] == selected_month]
    if selected_year != 'All':
        filtered_summary = filtered_summary[filtered_summary['Year'] == int(selected_year)]
    if selected_area != 'All':
        filtered_summary = filtered_summary[filtered_summary['Area'] == selected_area]
    if selected_regional != 'All':
        filtered_summary = filtered_summary[filtered_summary['Regional'] == selected_regional]
    if selected_nop != 'All':
        filtered_summary = filtered_summary[filtered_summary['NOP'] == selected_nop]
        
    # Display results
    st.subheader("📊 Summary Table")
    st.dataframe(filtered_summary)
    
    # Download button
    st.download_button("📥 Download Summary CSV", convert_df_to_csv(filtered_summary), "summary.csv", "text/csv")

    # Show full processed data
    with st.expander("🔍 View Full Processed Data"):
        st.dataframe(df)
        
    # Download processed data
    csv = df.to_csv(index=False).encode('utf-8')
    st.download_button(
        label="📥 Download Processed Data as CSV",
        data=csv,
        file_name='processed_kpi_data.csv',
        mime='text/csv'
    )

