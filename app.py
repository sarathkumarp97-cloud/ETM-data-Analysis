import streamlit as st
import pandas as pd
import os
from datetime import datetime, date
import calendar
import numpy as np
import time
import sys
import subprocess
import traceback
import base64

# Custom CSS for styling
st.markdown("""
    <style>
    .stApp {
        background: linear-gradient(135deg, #f5f7fa 0%, #e4edf5 100%);
    }
    .header {
        color: #2c3e50;
        text-align: center;
        padding: 1rem;
        font-size: 2.5rem;
        font-weight: 700;
        margin-bottom: 1.5rem;
        background: white;
        border-radius: 10px;
        box-shadow: 0 4px 6px rgba(0,0,0,0.1);
    }
    .section {
        background: white;
        border-radius: 10px;
        padding: 1.5rem;
        margin-bottom: 1.5rem;
        box-shadow: 0 4px 6px rgba(0,0,0,0.1);
    }
    .section-title {
        color: #3498db;
        border-bottom: 2px solid #3498db;
        padding-bottom: 0.5rem;
        margin-bottom: 1.5rem;
    }
    .config-box {
        background: #f8f9fa;
        border-left: 4px solid #3498db;
        padding: 1rem;
        margin-bottom: 1rem;
        border-radius: 4px;
    }
    .btn {
        background: #3498db !important;
        color: white !important;
        border: none !important;
        padding: 0.5rem 1rem !important;
        border-radius: 4px !important;
        font-weight: 600 !important;
    }
    .btn:hover {
        background: #2980b9 !important;
    }
    .file-count {
        font-size: 1.2rem;
        font-weight: 600;
        color: #2c3e50;
    }
    .day-type-btn {
        padding: 0.75rem 1.5rem !important;
        font-size: 1.1rem !important;
        margin: 0.5rem;
    }
    .selected {
        background: #2c3e50 !important;
        color: white !important;
    }
    </style>
""", unsafe_allow_html=True)

# Initialize session state
if 'analysis_complete' not in st.session_state:
    st.session_state.analysis_complete = False
if 'output_path' not in st.session_state:
    st.session_state.output_path = ""
if 'all_service' not in st.session_state:
    st.session_state.all_service = pd.DataFrame()
if 'all_route' not in st.session_state:
    st.session_state.all_route = pd.DataFrame()
if 'all_sample' not in st.session_state:
    st.session_state.all_sample = pd.DataFrame()
if 'selected_day_type' not in st.session_state:
    st.session_state.selected_day_type = ""

# Header
st.markdown('<div class="header">ETM Data Analysis Platform PS</div>', unsafe_allow_html=True)
st.markdown('<div class="subscript">By Sarath Kumar Ponugumati</div>', unsafe_allow_html=True)
# Sidebar Configuration
with st.sidebar:
    st.header("⚙️ Configuration")
    
    # Path Configuration
    input_folder = st.text_input("Input Folder", r"please paste the path that contains Etm data in Csv or bz2 format")
    output_folder = st.text_input("Output Folder", r"please paste path of output folder")
    lookup_file_path = st.text_input("Lookup File Path", r"please paste path of lookup file")
    lookup_sheet_name = st.text_input("Lookup Sheet Name", "Sheet3")
    
    # Verify paths button
    if st.button("Verify Paths", key="verify_paths"):
        errors = []
        valid = True
        
        if not os.path.exists(input_folder):
            errors.append(f"Input folder not found: {input_folder}")
            valid = False
        else:
            csv_files = [f for f in os.listdir(input_folder) if f.lower().endswith(('.csv', '.csv.bz2'))]
            if not csv_files:
                errors.append(f"No CSV files found in input folder: {input_folder}")
                valid = False
        
        if not os.path.exists(lookup_file_path):
            errors.append(f"Lookup file not found: {lookup_file_path}")
            valid = False
        elif not lookup_file_path.lower().endswith(('.xlsx', '.xls')):
            errors.append("Lookup file must be an Excel file (.xlsx or .xls)")
            valid = False
        
        os.makedirs(output_folder, exist_ok=True)
        
        if not valid:
            st.error("Errors found in configuration:")
            for error in errors:
                st.error(error)
        else:
            st.success("All paths are valid!")

# Main Content
st.header("🔍 Select Analysis Method")
analysis_method = st.radio("", [
    "Process all files in input folder",
    "Process files by date range",
    "Process files by month",
    "Process files by weekday/weekend"
], index=0)

# Get list of files
all_files = []
if os.path.exists(input_folder):
    all_files = [f for f in os.listdir(input_folder) if f.lower().endswith(('.csv', '.csv.bz2'))]
    st.info(f"Found {len(all_files)} CSV files in input folder")

# Date range selection
date_files = []
if analysis_method == "Process files by date range" and all_files:
    st.subheader("📅 Select Date Range")
    
    # Extract dates from filenames
    file_dates = []
    for file in all_files:
        try:
            if "Ticket-Data" in file:
                date_part = file.split('-')[2:5]
                date_str = '-'.join(date_part).split('.')[0]
                dt = datetime.strptime(date_str, '%Y-%m-%d')
            else:
                date_str = file.split('_')[1].split('.')[0]
                if len(date_str) == 8 and date_str.isdigit():
                    dt = datetime.strptime(date_str, '%d%m%Y')
                else:
                    continue
            file_dates.append(dt.date())
        except:
            continue
    
    if file_dates:
        min_date = min(file_dates)
        max_date = max(file_dates)
        
        col1, col2 = st.columns(2)
        with col1:
            start_date = st.date_input("Start Date", min_date)
        with col2:
            end_date = st.date_input("End Date", max_date)
        
        # Filter files by date range
        date_files = []
        for file in all_files:
            try:
                if "Ticket-Data" in file:
                    date_part = file.split('-')[2:5]
                    date_str = '-'.join(date_part).split('.')[0]
                    dt = datetime.strptime(date_str, '%Y-%m-%d')
                else:
                    date_str = file.split('_')[1].split('.')[0]
                    if len(date_str) == 8 and date_str.isdigit():
                        dt = datetime.strptime(date_str, '%d%m%Y')
                    else:
                        continue
                if start_date <= dt.date() <= end_date:
                    date_files.append(file)
            except:
                continue
        
        st.info(f"Found {len(date_files)} files in selected date range")

# Monthly analysis selection
month_files = []
if analysis_method == "Process files by month" and all_files:
    st.subheader("🗓️ Select Month")
    
    # Get monthly options
    month_year_options = {}
    for file in all_files:
        try:
            if "Ticket-Data" in file:
                date_part = file.split('-')[2:5]
                date_str = '-'.join(date_part).split('.')[0]
                dt = datetime.strptime(date_str, '%Y-%m-%d')
            else:
                date_str = file.split('_')[1].split('.')[0]
                if len(date_str) == 8 and date_str.isdigit():
                    dt = datetime.strptime(date_str, '%d%m%Y')
                else:
                    continue

            month_year = f"{calendar.month_name[dt.month]} {dt.year}"

            if month_year not in month_year_options:
                month_year_options[month_year] = []
            month_year_options[month_year].append(file)
        except:
            continue
    
    if month_year_options:
        sorted_months = sorted(month_year_options.keys(), key=lambda x: datetime.strptime(x, '%B %Y'))
        selected_month = st.selectbox("Available Months", sorted_months)
        month_files = month_year_options.get(selected_month, [])
        st.info(f"Found {len(month_files)} files for {selected_month}")
    else:
        st.warning("No valid monthly data found in filenames")

# Weekday/Weekend analysis selection
day_type_files = []
if analysis_method == "Process files by weekday/weekend" and all_files:
    st.subheader("📅 Select Day Type")
    
    # Extract day of week from filenames
    weekday_files = []
    weekend_files = []
    
    for file in all_files:
        try:
            if "Ticket-Data" in file:
                date_part = file.split('-')[2:5]
                date_str = '-'.join(date_part).split('.')[0]
                dt = datetime.strptime(date_str, '%Y-%m-%d')
            else:
                date_str = file.split('_')[1].split('.')[0]
                if len(date_str) == 8 and date_str.isdigit():
                    dt = datetime.strptime(date_str, '%d%m%Y')
                else:
                    continue
            
            # Monday=0, Sunday=6
            if dt.weekday() < 5:  # Monday to Friday
                weekday_files.append(file)
            else:  # Saturday and Sunday
                weekend_files.append(file)
        except:
            continue
    
    st.info(f"Found {len(weekday_files)} weekday files and {len(weekend_files)} weekend files")
    
    col1, col2 = st.columns(2)
    with col1:
        if st.button("Weekdays (Mon-Fri)", key="weekdays", type="primary" if st.session_state.selected_day_type == "weekdays" else "secondary"):
            st.session_state.selected_day_type = "weekdays"
    with col2:
        if st.button("Weekends (Sat-Sun)", key="weekends", type="primary" if st.session_state.selected_day_type == "weekends" else "secondary"):
            st.session_state.selected_day_type = "weekends"
    
    if st.session_state.selected_day_type == "weekdays":
        day_type_files = weekday_files
        st.success(f"Selected Weekdays: {len(day_type_files)} files")
    elif st.session_state.selected_day_type == "weekends":
        day_type_files = weekend_files
        st.success(f"Selected Weekends: {len(day_type_files)} files")

# Function to parse dates
def parse_date(date_str):
    try:
        for fmt in ('%Y-%m-%d', '%d-%m-%Y', '%d/%m/%Y', '%m/%d/%Y'):
            try:
                return datetime.strptime(date_str, fmt)
            except ValueError:
                continue
        return pd.to_datetime(date_str, dayfirst=True)
    except Exception:
        return None

# Main processing function
def process_file(file_path, lookup_df):
    try:
        # Read the CSV file
        try:
            df = pd.read_csv(
                file_path,
                compression='bz2' if file_path.endswith('.bz2') else None
            )
        except Exception as e:
            st.error(f"Error reading {os.path.basename(file_path)}: {str(e)}")
            return pd.DataFrame(), pd.DataFrame(), pd.DataFrame()

        # Check required columns
        required_cols = {'schedule_no', 'fleet_no', 'from_stage', 'to_stage', 'ticket_type'}
        if not required_cols.issubset(df.columns):
            st.error(f"Missing required columns in {os.path.basename(file_path)}")
            return pd.DataFrame(), pd.DataFrame(), pd.DataFrame()

        # Process schedule numbers
        try:
            split_result = df['schedule_no'].str.split('-', expand=True)
            if split_result.shape[1] >= 4:
                df[['service_type', 'route_no', 'service_id', 'shift']] = split_result.iloc[:, :4]
                df['service_no'] = df['route_no'].astype(str) + "-" + df['service_id'].astype(str)
                df['trip_shift'] = df['trip_no'].astype(str) + "-" + df['shift'].astype(str)
                df['service_trip_shift'] = df['service_no'].astype(str) + "_" + df['trip_shift'].astype(str) + "_" + df['service_type'].astype(str)
            else:
                st.error(f"Invalid schedule_no format in {os.path.basename(file_path)}")
                return pd.DataFrame(), pd.DataFrame(), pd.DataFrame()
        except Exception as e:
            st.error(f"Error processing schedule numbers in {os.path.basename(file_path)}: {str(e)}")
            return pd.DataFrame(), pd.DataFrame(), pd.DataFrame()

        # Convert stage columns to numeric
        df['to_stage'] = pd.to_numeric(df['to_stage'], errors='coerce')
        df['from_stage'] = pd.to_numeric(df['from_stage'], errors='coerce')

        # Removing dummy trips with KM zero
        df['KM'] = pd.to_numeric(df['KM'], errors='coerce')
        df = df[df['KM'].ne(0) | df['KM'].isna()]

        # Total Stages per each route
        df['route_description'] = df['Trip_Origin'].astype(str) + " To " + df['Trip_Destination'].astype(str)
        df['route_desc'] = df['route_no'].astype(str) + "-" + df['route_description'].astype(str)
        
        # Calculate total stages correctly
        df['max_stage'] = df[['to_stage', 'from_stage']].max(axis=1)
        df['total_stages'] = df.groupby('route_desc')['max_stage'].transform('max')
        df.drop(columns=['max_stage'], inplace=True)

        # Parse dates
        base_date = parse_date(df["date"].iloc[0])
        if base_date:
            date_str = base_date.strftime('%d-%m-%y')
            day_str = base_date.strftime('%A')
        else:
            date_str = df["date"].iloc[0]
            day_str = "Unknown"

        # Sample Dataframe creation with Ticket Counts
        ticket_counts = {
            'date': date_str,
            'day': day_str,
            'TC01_Adult': len(df[df['ticket_type'] == 'Adult Ticket']),
            'TC02_Difference': len(df[df['ticket_type'] == 'Difference Fare Ticket']),
            'TC03_Dis_esc': len(df[df["ticket_type"] == "Disable Person Escort Ticket"]),
            'TC04_Dis': len(df[df["ticket_type"] == "Disable Person Ticket"]),
            'TC05_Luggage': len(df[df["ticket_type"] == "Luggage Ticket"]),
            'TC06_Senior': len(df[df["ticket_type"] == "Senior Citizen Ticket"]),
            'TC07_Transgender': len(df[df['ticket_type'] == 'Transgender Ticket']),
            'TC08_Women': len(df[df["ticket_type"] == "Women Ticket"]),
            'all_tickets': len(df)}
        
        sample_size = pd.DataFrame([ticket_counts])

        # Passengers Counts for Each type of Ticket
        passenger_counts = df.groupby('ticket_type')['passenger_count'].sum()
        sample_size['PC01'] = passenger_counts.get('Adult Ticket', 0)
        sample_size['PC02'] = passenger_counts.get('Difference Fare Ticket', 0)
        sample_size['PC03'] = passenger_counts.get('Disable Person Escort Ticket', 0)
        sample_size['PC04'] = passenger_counts.get('Disable Person Ticket', 0)
        sample_size['PC06'] = passenger_counts.get('Senior Citizen Ticket', 0)
        sample_size['PC07'] = passenger_counts.get('Transgender Ticket', 0)
        sample_size['PC08'] = passenger_counts.get('Women Ticket', 0)
        sample_size['total_passengers'] = df['passenger_count'].sum()

        # Revenue Calculations
        revenue_counts = df.groupby('ticket_type')['total_ticket_amount'].sum()
        luggage_revenue = df[df['ticket_type'] == 'Luggage Ticket']['luggage_amount'].sum()
        
        sample_size['R01'] = revenue_counts.get('Adult Ticket', 0)
        sample_size['R02'] = revenue_counts.get('Difference Fare Ticket', 0)
        sample_size['R03'] = sample_size['PC03'] * 16
        sample_size['R04'] = sample_size['PC04'] * 16
        sample_size['R05'] = luggage_revenue
        sample_size['R06'] = sample_size['PC06'] * 21
        sample_size['R07'] = sample_size['PC07'] * 16
        sample_size['R08'] = sample_size['PC08'] * 16
        sample_size['Earnings_per_day'] = sample_size[
            ['R01', 'R02', 'R03', 'R04', 'R05', 'R06', 'R07', 'R08']
        ].sum(axis=1)

        # Merge capacity data
        df = pd.merge(df, lookup_df[['fleet_no', 'total_capacity']], on='fleet_no', how='left')
        df["total_capacity"] = df["total_capacity"].fillna(73)

        # Demand calculations
        df['stages_travelled'] = np.where(df['to_stage'] == 0,
            0, (df['to_stage'] - df['from_stage']).abs()
        )
        
        df['demand'] = df['passenger_count'] * df['stages_travelled']
        df['supply'] = df['total_capacity'] * (df['total_stages'] - 1)

        # Create helping DataFrame
        helping_df = df[['service_trip_shift']].drop_duplicates()
        helping_df[['service_no', 'trip_shift', 'service_type']] = helping_df['service_trip_shift'].str.split("_", expand=True)
        helping_df['route_no'] = helping_df['service_no'].str.split("-").str[0]
        
        # Aggregate metrics
        ticket_types_free = [
            ('TCH03', 'Disable Person Escort Ticket'),
            ('TCH04', 'Disable Person Ticket'),
            ('TCH06', 'Senior Citizen Ticket'),
            ('TCH07', 'Transgender Ticket'),
            ('TCH08', 'Women Ticket')]
        
        for col, ticket_type in ticket_types_free:
            helping_df[col] = helping_df['service_trip_shift'].map(df[df['ticket_type'] == ticket_type]
                    .groupby('service_trip_shift')['passenger_count'].sum()).fillna(0)
        
        helping_df['zero_value_ticket_pass'] = helping_df[['TCH03','TCH04','TCH06','TCH07','TCH08']].sum(axis=1)
            
        helping_df = helping_df.merge(df.groupby('service_trip_shift')['stages_travelled'].mean().rename('average_stages'), 
            on='service_trip_shift').round(2)
        helping_df['zero_value_ticket_demand'] = (helping_df['zero_value_ticket_pass'] * helping_df['average_stages']).round(2)
        helping_df = helping_df.merge(df.groupby('service_trip_shift')['KM'].mean().rename('kms'), on='service_trip_shift').round(2)
        helping_df = helping_df.merge(df.groupby('service_trip_shift')['total_stages'].mean().rename('stages'), on='service_trip_shift')
        helping_df = helping_df.merge(df.groupby('service_trip_shift')['total_capacity'].mean().rename('capacity'),on='service_trip_shift').round(2)
        helping_df['supply'] = (helping_df['capacity'] * helping_df['stages']).round(2)

        # Service-level aggregates
        service_help = helping_df.groupby('service_no').agg(service_kms=('kms', 'sum'),service_supply=('supply', 'sum'), 
                zero_value_ticket_demand_service=('zero_value_ticket_demand', 'sum')).reset_index().round(2)
        
        # Route-level aggregates
        route_help = helping_df.groupby('route_no').agg(route_kms=('kms', 'sum'),route_supply=('supply', 'sum'),
                zero_value_ticket_demand_route=('zero_value_ticket_demand', 'sum')).reset_index().round(2)

        # Service DataFrame
        service_df = df[["service_no", "service_type"]].drop_duplicates()
        service_df['route_no'] = service_df['service_no'].str.split("-").str[0]
        service_df = service_df[['service_no', 'service_type', 'route_no']]
        
        # Parse date for service_df
        duty_date = parse_date(df["duty_date"].iloc[0])
        if duty_date:
            service_df["date"] = duty_date.strftime('%d-%m-%y')
            service_df["day"] = duty_date.strftime('%A')
        else:
            service_df["date"] = df["duty_date"].iloc[0]
            service_df["day"] = "Unknown"

        # Ticket Count for service
        ticket_types = [
            ('TCS01', 'Adult Ticket'),
            ('TCS02', 'Difference Fare Ticket'),
            ('TCS03', 'Disable Person Escort Ticket'),
            ('TCS04', 'Disable Person Ticket'),
            ('TCS05', 'Luggage Ticket'),
            ('TCS06', 'Senior Citizen Ticket'),
            ('TCS07', 'Transgender Ticket'),
            ('TCS08', 'Women Ticket')]
        
        for col, ticket_type in ticket_types:
            service_df[col] = service_df['service_no'].map(df[df['ticket_type'] == ticket_type].groupby('service_no').size()).fillna(0)
        
        service_df['TTPS'] = service_df[[f'TCS{i:02d}' for i in range(1, 9)]].sum(axis=1)
        service_df = service_df.merge(df.groupby('service_no').size().rename('TTPS_Check'), on='service_no',how='left').fillna(0)

        # Passenger count per service
        passenger_cols = [
            ('PCS01', 'Adult Ticket'),
            ('PCS02', 'Difference Fare Ticket'),
            ('PCS03', 'Disable Person Escort Ticket'),
            ('PCS04', 'Disable Person Ticket'),
            ('PCS06', 'Senior Citizen Ticket'),
            ('PCS07', 'Transgender Ticket'),
            ('PCS08', 'Women Ticket')
        ]
        
        for col, ticket_type in passenger_cols:
            service_df[col] = service_df['service_no'].map(
                df[df['ticket_type'] == ticket_type]
                .groupby('service_no')['passenger_count'].sum()
            ).fillna(0)
            
        service_df['PPSPD'] = service_df[[
            'PCS01', 'PCS02', 'PCS03', 'PCS04', 'PCS06', 'PCS07', 'PCS08'
        ]].sum(axis=1)
        
        service_df = service_df.merge(
            df.groupby('service_no')['passenger_count'].sum().rename('PPSPD_Check'),
            on='service_no',
            how='left'
        ).fillna(0)

        # Revenue per service
        service_df['RS01'] = service_df['service_no'].map(
    df[df['ticket_type'] == 'Adult Ticket']
    .groupby('service_no')['passenger_amount']
    .sum()
).fillna(0)
        service_df['RS02'] = service_df['service_no'].map(
    df[df['ticket_type'] == 'Difference Fare Ticket']
    .groupby('service_no')['passenger_amount']
    .sum()
).fillna(0)
        service_df['RS03'] = (service_df['PCS03'] * 16).astype(int)
        service_df['RS04'] = (service_df['PCS04'] * 16).astype(int)
        service_df['RS05'] = service_df['service_no'].map(
    df[df['ticket_type'] == 'Luggage Ticket']
    .groupby('service_no')['luggage_amount']
    .sum()
).fillna(0)
        service_df['RS06'] = (service_df['PCS06'] * 21).astype(int)
        service_df['RS07'] = (service_df['PCS07'] * 16).astype(int)
        service_df['RS08'] = (service_df['PCS08'] * 16).astype(int)

        service_df['EPSPD'] = service_df[[f'RS{i:02d}' for i in range(1, 9)]].sum(axis=1)
        
        # Service metrics
        service_metrics = df.groupby("service_no").agg(trips_per_service=('trip_no', 'nunique')).reset_index()
        service_df = service_df.merge(service_metrics, on="service_no", how='left')

        # Service Demand and supply
        service_demand = df.groupby('service_no')['demand'].sum().reset_index()
        service_df = service_df.merge(service_demand, on='service_no', how='left').fillna(0)
        service_df = pd.merge(service_df, service_help, on='service_no', how='left')
        service_df["service_DSRWOW"] = (service_df["demand"] / service_df["service_supply"]).round(2).fillna(0)
        service_df['total_service_demand'] = (service_df["demand"] + service_df["zero_value_ticket_demand_service"]).round(2)
        service_df['service_DSRWW'] = (service_df['total_service_demand']/service_df["service_supply"]).round(2)
        service_df.sort_values('service_DSRWW', ascending=False)
        sample_size['service_count'] = len(service_df)

        # Route DataFrame
        route_df = df[['route_no']].drop_duplicates()
        
        # Parse date for route_df
        duty_date = parse_date(df["duty_date"].iloc[0])
        if duty_date:
            route_df["date"] = duty_date.strftime('%d-%m-%y')
            route_df["day"] = duty_date.strftime('%A')
        else:
            route_df["date"] = df["duty_date"].iloc[0]
            route_df["day"] = "Unknown"

        # Ticket Count for route
        for i, (col, ticket_type) in enumerate(ticket_types, 1):
            route_df[f'TCR{i:02d}'] = route_df['route_no'].map(
                df[df['ticket_type'] == ticket_type]
                .groupby('route_no').size()
            ).fillna(0)
            
        route_df['TTPR'] = route_df[[f'TCR{i:02d}' for i in range(1, 9)]].sum(axis=1)
        route_df = route_df.merge(df.groupby('route_no').size().rename('TTPR_Check'),on='route_no',how='left').fillna(0)\
        #passenger count per route
        route_df['PCR01'] = route_df['route_no'].map(
    df[df['ticket_type'] == 'Adult Ticket']
    .groupby('route_no')['passenger_count']
    .sum()
).fillna(0)
        route_df['PCR02'] = route_df['route_no'].map(
    df[df['ticket_type'] == 'Difference Fare Ticket']
    .groupby('route_no')['passenger_count']
    .sum()
).fillna(0)
        route_df['PCR03'] = route_df['route_no'].map(
    df[df['ticket_type'] == 'Disable Person Escort Ticket']
    .groupby('route_no')['passenger_count']
    .sum()
).fillna(0)
        route_df['PCR04'] = route_df['route_no'].map(
    df[df['ticket_type'] == 'Disable Person Ticket']
    .groupby('route_no')['passenger_count']
    .sum()
).fillna(0)
        route_df['PCR06'] = route_df['route_no'].map(
    df[df['ticket_type'] == 'Senior Citizen Ticket']
    .groupby('route_no')['passenger_count']
    .sum()
).fillna(0)
        route_df['PCR07'] = route_df['route_no'].map(
    df[df['ticket_type'] == 'Transgender Ticket']
    .groupby('route_no')['passenger_count']
    .sum()
).fillna(0)
        route_df['PCR08'] = route_df['route_no'].map(
    df[df['ticket_type'] == 'Women Ticket']
    .groupby('route_no')['passenger_count']
    .sum()
).fillna(0)
        # Total no of passengers per Service
        route_df['PPRPD'] = route_df[['PCR01', "PCR02","PCR03","PCR04","PCR06","PCR07","PCR08"]].sum(axis=1)
        route_df = route_df.merge(df.groupby('route_no')['passenger_count'].sum().rename('PPRPD_Check'), on = 'route_no')
    #REVENUE FOR ROUTE
        route_df['RR01'] = route_df['route_no'].map(
    df[df['ticket_type'] == 'Adult Ticket']
    .groupby('route_no')['passenger_amount']
    .sum()
).fillna(0)
        route_df['RR02'] = route_df['route_no'].map(
    df[df['ticket_type'] == 'Difference Fare Ticket']
    .groupby('route_no')['passenger_amount']
    .sum()
).fillna(0)
        route_df['RR03'] = (route_df['PCR03'] * 16).astype(int)
        route_df['RR04'] = (route_df['PCR04'] * 16).astype(int)
        route_df['RR05'] = route_df['route_no'].map(
    df[df['ticket_type'] == 'Luggage Ticket']
    .groupby('route_no')['luggage_amount']
    .sum()
).fillna(0)
        route_df['RR06'] = (route_df['PCR06'] * 21).astype(int)
        route_df['RR07'] = (route_df['PCR07'] * 16).astype(int)
        route_df['RR08'] = (route_df['PCR08'] * 16).astype(int)
        
        route_df['EPRPD'] = route_df[['RR01', 'RR02', 'RR03', 'RR04', 'RR05', 'RR06', 'RR07', 'RR08']].sum(axis=1)

        # Demand and Supply Ratio Calculations
        route_demand = df.groupby('route_no')['demand'].sum().reset_index()
        route_df = route_df.merge(route_demand, on='route_no').fillna(0)
        route_df = pd.merge(route_df, route_help, on='route_no', how='left')
        route_df['route_DSRWOW'] = (route_df['demand'] / route_df['route_supply']).round(2).fillna(0)
        route_df['total_route_demand'] = (route_df["demand"] + route_df["zero_value_ticket_demand_route"]).round(2)
        route_df['route_DSRWW'] = (route_df['total_route_demand'] / route_df['route_supply']).round(2)
        
        route_df.sort_values('route_DSRWW', ascending=False)
        
        sample_size['route_count'] = len(route_df)

        return service_df, route_df, sample_size

    except Exception as e:
        st.error(f"Error processing {os.path.basename(file_path)}: {str(e)}")
        st.text(traceback.format_exc())
        return pd.DataFrame(), pd.DataFrame(), pd.DataFrame()

# Run Analysis Button
if st.button("🚀 Run Analysis", key="run_analysis"):
    # Determine which files to process
    if analysis_method == "Process all files in input folder":
        files_to_process = all_files
    elif analysis_method == "Process files by date range":
        files_to_process = date_files
    elif analysis_method == "Process files by month":
        files_to_process = month_files
    elif analysis_method == "Process files by weekday/weekend":
        files_to_process = day_type_files
    
    if not files_to_process:
        st.error("No files selected for processing!")
        st.stop()
    
    # Verify paths
    errors = []
    valid = True
    
    if not os.path.exists(input_folder):
        errors.append(f"Input folder not found: {input_folder}")
        valid = False
    else:
        csv_files = [f for f in os.listdir(input_folder) if f.lower().endswith(('.csv', '.csv.bz2'))]
        if not csv_files:
            errors.append(f"No CSV files found in input folder: {input_folder}")
            valid = False
    
    if not os.path.exists(lookup_file_path):
        errors.append(f"Lookup file not found: {lookup_file_path}")
        valid = False
    elif not lookup_file_path.lower().endswith(('.xlsx', '.xls')):
        errors.append("Lookup file must be an Excel file (.xlsx or .xls)")
        valid = False
    
    if not valid:
        st.error("Errors found in configuration:")
        for error in errors:
            st.error(error)
        st.stop()
    
    # Load lookup data
    try:
        lookup_df = pd.read_excel(lookup_file_path, sheet_name=lookup_sheet_name)
        if 'fleet_no' not in lookup_df.columns or 'total_capacity' not in lookup_df.columns:
            st.error("Lookup file missing required columns ('fleet_no' and/or 'total_capacity')")
            st.stop()
    except Exception as e:
        st.error(f"Error loading lookup file: {e}")
        st.stop()
    
    # Initialize progress
    progress_bar = st.progress(0)
    status_text = st.empty()
    
    # Initialize results
    all_service = pd.DataFrame()
    all_route = pd.DataFrame()
    all_sample = pd.DataFrame()
    
    # Process files
    for i, file in enumerate(files_to_process):
        status_text.text(f"Processing {i+1}/{len(files_to_process)}: {file}")
        progress_bar.progress(int((i + 1) / len(files_to_process) * 100))
        
        try:
            service, route, sample = process_file(os.path.join(input_folder, file), lookup_df)
            
            if not service.empty:
                all_service = pd.concat([all_service, service], ignore_index=True)
            if not route.empty:
                all_route = pd.concat([all_route, route], ignore_index=True)
            if not sample.empty:
                all_sample = pd.concat([all_sample, sample], ignore_index=True)
            
        except Exception as e:
            st.error(f"Error processing {file}: {str(e)}")
            st.text(traceback.format_exc())
    
    # Store results in session state
    st.session_state.all_service = all_service
    st.session_state.all_route = all_route
    st.session_state.all_sample = all_sample
    st.session_state.analysis_complete = True
    
    # Clear progress bar and status
    progress_bar.empty()
    status_text.empty()
    
    st.success(f"✅ Analysis completed successfully! Processed {len(files_to_process)} files.")

# Display results if analysis is complete
if st.session_state.analysis_complete:
    st.header("📊 Analysis Results")
    
    if not st.session_state.all_service.empty:
        st.subheader("Service Data")
        st.dataframe(st.session_state.all_service)
    
    if not st.session_state.all_route.empty:
        st.subheader("Route Data")
        st.dataframe(st.session_state.all_route)
    
    if not st.session_state.all_sample.empty:
        st.subheader("Summary Data")
        st.dataframe(st.session_state.all_sample)
    
    # Download button for results
    timestamp = datetime.now().strftime("%d-%m-%Y_%H%M%S")
    output_excel_name = f"ETM_Analysis_PS_{timestamp}.xlsx"
    output_path = os.path.join(output_folder, output_excel_name)
    
    # Save results to Excel
    try:
        with pd.ExcelWriter(output_path, engine='openpyxl') as writer:
            if not st.session_state.all_service.empty:
                st.session_state.all_service.to_excel(writer, sheet_name='Service_Data', index=False)
            if not st.session_state.all_route.empty:
                st.session_state.all_route.to_excel(writer, sheet_name='Route_Data', index=False)
            if not st.session_state.all_sample.empty:
                st.session_state.all_sample.to_excel(writer, sheet_name='Summary_Data', index=False)
        
        st.session_state.output_path = output_path
        st.info(f"Results saved to: {output_path}")
        
        # Download button
        with open(output_path, "rb") as f:
            bytes_data = f.read()
            b64 = base64.b64encode(bytes_data).decode()
            href = f'<a href="data:application/vnd.openxmlformats-officedocument.spreadsheetml.sheet;base64,{b64}" download="{output_excel_name}">📥 Download Full Analysis Results</a>'
            st.markdown(href, unsafe_allow_html=True)
            
    except Exception as e:
        st.error(f"Error saving Excel file: {e}")

# Instructions Section
st.sidebar.header("ℹ️ Instructions")
st.sidebar.info("""
1. **Configure Paths**: Set input/output folders and lookup file
2. **Select Analysis Method**:
   - All Files: Process every CSV in input folder
   - Date Range: Select start/end dates
   - Monthly: Choose month from dropdown
   - Weekday/Weekend: Analyze Mon-Fri or Sat-Sun
3. **Verify Paths**: Check configuration is correct
4. **Run Analysis**: Process files and generate results
5. **Download**: Get Excel report when complete
""")

st.sidebar.header("📝 Notes")
st.sidebar.warning("""
- Ensure lookup file has 'fleet_no' and 'total_capacity' columns
- Files should be named with dates for date/month analysis
- Output folder will be created if it doesn't exist
- Weekday/Weekend analysis requires date in filename
""")
