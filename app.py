import streamlit as st
import pandas as pd
import numpy as np
import os
import io
import gc
import calendar
import traceback
from datetime import datetime

# Only the columns the analysis actually uses (keeps memory low on Streamlit Cloud)
USECOLS = {
    'schedule_no', 'fleet_no', 'from_stage', 'to_stage', 'ticket_type',
    'trip_no', 'KM', 'Trip_Origin', 'Trip_Destination', 'date', 'duty_date',
    'passenger_count', 'total_ticket_amount', 'luggage_amount', 'passenger_amount'
}

MAX_ROWS_SHOWN = 1000


@st.cache_data(show_spinner=False)
def get_embedded_fleet_lookup():
    """Builds the embedded MTC fleet capacity lookup DataFrame (3,390 vehicles)."""
    rows = []
    for capacity, fleet_str in FLEET_CAPACITY_GROUPS.items():
        for fleet_no in fleet_str.split():
            rows.append((fleet_no, capacity))
    return pd.DataFrame(rows, columns=["fleet_no", "total_capacity"])


def find_sample_file():
    """Finds the sample ETM file in the repo root regardless of .csv or .csv.bz2 extension."""
    for candidate in ("sample_etm_data.csv.bz2", "sample_etm_data.csv", "sample_etm_data.bz2"):
        if os.path.exists(candidate):
            return candidate
    for f in os.listdir("."):
        if f.lower().startswith("sample_etm_data") and f.lower().endswith((".csv", ".bz2")):
            return f
    return None


def detect_compression(source):
    """Detects bzip2 compression from actual file bytes (magic header 'BZh'), regardless of filename."""
    if isinstance(source, str):
        with open(source, "rb") as f:
            header = f.read(3)
    else:
        pos = source.tell()
        source.seek(0)
        header = source.read(3)
        source.seek(pos)
    return "bz2" if header == b"BZh" else None


def read_etm_source(source):
    """Reads only the needed columns of an ETM file (auto-detects .bz2 content)."""
    comp = detect_compression(source)
    if not isinstance(source, str):
        source.seek(0)
    return pd.read_csv(
        source,
        compression=comp,
        usecols=lambda c: c in USECOLS,
    )


def parse_date(date_str):
    try:
        date_str = str(date_str).strip()
        for fmt in ('%Y-%m-%d', '%d-%m-%Y', '%d/%m/%Y', '%m/%d/%Y', '%d-%m-%y'):
            try:
                return datetime.strptime(date_str, fmt)
            except ValueError:
                continue
        return pd.to_datetime(date_str, dayfirst=True)
    except Exception:
        return None


def extract_file_date(file_name, source):
    """Extracts the date from the filename first, falling back to reading the first row inside the file."""
    try:
        if "Ticket-Data" in file_name:
            date_part = file_name.split('-')[2:5]
            date_str = '-'.join(date_part).split('.')[0]
            return datetime.strptime(date_str, '%Y-%m-%d').date()
        elif '_' in file_name:
            date_str = file_name.split('_')[1].split('.')[0]
            if len(date_str) == 8 and date_str.isdigit():
                return datetime.strptime(date_str, '%d%m%Y').date()
    except Exception:
        pass

    try:
        comp = detect_compression(source)
        if not isinstance(source, str):
            source.seek(0)
        sample_df = pd.read_csv(source, compression=comp, nrows=1)
        if not isinstance(source, str):
            source.seek(0)

        date_col = 'date' if 'date' in sample_df.columns else ('duty_date' if 'duty_date' in sample_df.columns else None)
        if date_col and not sample_df.empty:
            parsed = parse_date(sample_df[date_col].iloc[0])
            if parsed is not None and not pd.isna(parsed):
                return parsed.date()
    except Exception:
        pass
    return None


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
        margin-bottom: 0.5rem;
        background: white;
        border-radius: 10px;
        box-shadow: 0 4px 6px rgba(0,0,0,0.1);
    }
    .subscript {
        text-align: center;
        color: #5d6d7e;
        font-weight: 600;
        margin-bottom: 1.5rem;
    }
    </style>
""", unsafe_allow_html=True)

# Initialize session state
if 'analysis_complete' not in st.session_state:
    st.session_state.analysis_complete = False
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

# Load embedded fleet lookup
lookup_df = get_embedded_fleet_lookup()

# Sidebar Configuration
with st.sidebar:
    st.header("⚙️ Data Source & Info")
    data_source = st.radio(
        "Select Data Source",
        ["Use Sample ETM Data", "Upload ETM File(s)"],
        index=0
    )
    st.caption(f"✅ Fleet lookup built-in ({len(lookup_df):,} vehicles; missing fleet capacity = 65)")

# Collect available files as list of (display_name, file_source)
available_items = []
if data_source == "Use Sample ETM Data":
    sample_path = find_sample_file()
    if sample_path:
        available_items = [(sample_path, sample_path)]
        st.info(f"Loaded built-in sample ETM dataset: `{sample_path}`")
    else:
        st.error("Sample ETM file (`sample_etm_data.csv.bz2` or `sample_etm_data.csv`) was not found in the repository root.")
else:
    uploaded_files = st.file_uploader(
        "Upload one or more ETM files (.csv or .bz2)",
        type=["csv", "bz2"],
        accept_multiple_files=True
    )
    if uploaded_files:
        available_items = [(uf.name, uf) for uf in uploaded_files]
        st.info(f"Uploaded {len(available_items)} ETM file(s)")

# Main Content: Analysis Method
st.header("🔍 Select Analysis Method")
analysis_method = st.radio(
    "Choose how to filter files:",
    [
        "Process all selected files",
        "Process files by date range",
        "Process files by month",
        "Process files by weekday/weekend"
    ],
    index=0
)

# Map files to dates when date-based filtering is selected
item_date_map = {}
if available_items and analysis_method != "Process all selected files":
    for name, src in available_items:
        dt = extract_file_date(name, src)
        if dt is not None:
            item_date_map[name] = (dt, src)

# Date range selection
filtered_items = []
if analysis_method == "Process files by date range" and available_items:
    st.subheader("📅 Select Date Range")
    if item_date_map:
        dates = [v[0] for v in item_date_map.values()]
        col1, col2 = st.columns(2)
        with col1:
            start_date = st.date_input("Start Date", min(dates))
        with col2:
            end_date = st.date_input("End Date", max(dates))
        filtered_items = [(name, src) for name, (dt, src) in item_date_map.items() if start_date <= dt <= end_date]
        st.info(f"Found {len(filtered_items)} file(s) in selected date range")
    else:
        st.warning("No valid dates found inside the files.")

# Monthly analysis selection
elif analysis_method == "Process files by month" and available_items:
    st.subheader("🗓️ Select Month")
    if item_date_map:
        month_year_options = {}
        for name, (dt, src) in item_date_map.items():
            month_year = f"{calendar.month_name[dt.month]} {dt.year}"
            month_year_options.setdefault(month_year, []).append((name, src))
        sorted_months = sorted(month_year_options.keys(), key=lambda x: datetime.strptime(x, '%B %Y'))
        selected_month = st.selectbox("Available Months", sorted_months)
        filtered_items = month_year_options.get(selected_month, [])
        st.info(f"Found {len(filtered_items)} file(s) for {selected_month}")
    else:
        st.warning("No valid monthly data found inside the files.")

# Weekday/Weekend analysis selection
elif analysis_method == "Process files by weekday/weekend" and available_items:
    st.subheader("📅 Select Day Type")
    weekday_items = [(name, src) for name, (dt, src) in item_date_map.items() if dt.weekday() < 5]
    weekend_items = [(name, src) for name, (dt, src) in item_date_map.items() if dt.weekday() >= 5]

    st.info(f"Found {len(weekday_items)} weekday file(s) and {len(weekend_items)} weekend file(s)")

    col1, col2 = st.columns(2)
    with col1:
        if st.button("Weekdays (Mon-Fri)", key="weekdays", type="primary" if st.session_state.selected_day_type == "weekdays" else "secondary"):
            st.session_state.selected_day_type = "weekdays"
    with col2:
        if st.button("Weekends (Sat-Sun)", key="weekends", type="primary" if st.session_state.selected_day_type == "weekends" else "secondary"):
            st.session_state.selected_day_type = "weekends"

    if st.session_state.selected_day_type == "weekdays":
        filtered_items = weekday_items
        st.success(f"Selected Weekdays: {len(filtered_items)} file(s)")
    elif st.session_state.selected_day_type == "weekends":
        filtered_items = weekend_items
        st.success(f"Selected Weekends: {len(filtered_items)} file(s)")


# Main processing function
def process_file(file_name, file_source, lookup_df):
    try:
        try:
            df = read_etm_source(file_source)
        except Exception as e:
            st.error(f"Error reading {file_name}: {str(e)}")
            return pd.DataFrame(), pd.DataFrame(), pd.DataFrame()

        required_cols = {'schedule_no', 'fleet_no', 'from_stage', 'to_stage', 'ticket_type', 'trip_no'}
        if not required_cols.issubset(df.columns):
            missing = required_cols - set(df.columns)
            st.error(f"Missing required columns in {file_name}: {', '.join(sorted(missing))}")
            return pd.DataFrame(), pd.DataFrame(), pd.DataFrame()

        # Process schedule numbers
        try:
            split_result = df['schedule_no'].astype(str).str.split('-', expand=True)
            if split_result.shape[1] >= 4:
                df[['service_type', 'route_no', 'service_id', 'shift']] = split_result.iloc[:, :4]
                del split_result
                df['service_no'] = df['route_no'].astype(str) + "-" + df['service_id'].astype(str)
                df['trip_shift'] = df['trip_no'].astype(str) + "-" + df['shift'].astype(str)
                df['service_trip_shift'] = df['service_no'].astype(str) + "_" + df['trip_shift'].astype(str) + "_" + df['service_type'].astype(str)
                df.drop(columns=['service_id', 'shift', 'trip_shift', 'schedule_no'], inplace=True, errors='ignore')
            else:
                st.error(f"Invalid schedule_no format in {file_name}")
                return pd.DataFrame(), pd.DataFrame(), pd.DataFrame()
        except Exception as e:
            st.error(f"Error processing schedule numbers in {file_name}: {str(e)}")
            return pd.DataFrame(), pd.DataFrame(), pd.DataFrame()

        # Convert stage columns to numeric
        df['to_stage'] = pd.to_numeric(df['to_stage'], errors='coerce')
        df['from_stage'] = pd.to_numeric(df['from_stage'], errors='coerce')

        # Removing dummy trips with KM zero
        df['KM'] = pd.to_numeric(df['KM'], errors='coerce')
        df = df[df['KM'].ne(0) | df['KM'].isna()].copy()

        # Total Stages per each route
        df['route_description'] = df['Trip_Origin'].astype(str) + " To " + df['Trip_Destination'].astype(str)
        df['route_desc'] = df['route_no'].astype(str) + "-" + df['route_description'].astype(str)

        df['max_stage'] = df[['to_stage', 'from_stage']].max(axis=1)
        df['total_stages'] = df.groupby('route_desc')['max_stage'].transform('max')
        df.drop(columns=['max_stage', 'route_description', 'route_desc'], inplace=True)

        # Parse dates
        base_date = parse_date(df["date"].iloc[0])
        if base_date is not None and not pd.isna(base_date):
            date_str = base_date.strftime('%d-%m-%y')
            day_str = base_date.strftime('%A')
        else:
            date_str = str(df["date"].iloc[0])
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
            'all_tickets': len(df)
        }
        sample_size = pd.DataFrame([ticket_counts])

        # Passenger Counts for Each type of Ticket
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

        # Merge embedded capacity data; default to 65 if fleet_no is missing or unmatched
        df['fleet_no_clean'] = df['fleet_no'].astype(str).str.strip().str.upper()
        lookup_clean = lookup_df.copy()
        lookup_clean['fleet_no_clean'] = lookup_clean['fleet_no'].astype(str).str.strip().str.upper()
        df = pd.merge(df, lookup_clean[['fleet_no_clean', 'total_capacity']], on='fleet_no_clean', how='left')
        df["total_capacity"] = df["total_capacity"].fillna(65)
        df.drop(columns=['fleet_no_clean'], inplace=True)

        # Demand calculations
        df['stages_travelled'] = np.where(
            df['to_stage'] == 0,
            0,
            (df['to_stage'] - df['from_stage']).abs()
        )
        df['demand'] = df['passenger_count'] * df['stages_travelled']
        df['supply'] = df['total_capacity'] * (df['total_stages'] - 1)

        # Create helping DataFrame
        helping_df = df[['service_trip_shift']].drop_duplicates()
        helping_df[['service_no', 'trip_shift', 'service_type']] = helping_df['service_trip_shift'].str.split("_", expand=True)
        helping_df['route_no'] = helping_df['service_no'].str.split("-").str[0]

        ticket_types_free = [
            ('TCH03', 'Disable Person Escort Ticket'),
            ('TCH04', 'Disable Person Ticket'),
            ('TCH06', 'Senior Citizen Ticket'),
            ('TCH07', 'Transgender Ticket'),
            ('TCH08', 'Women Ticket')
        ]

        for col, ticket_type in ticket_types_free:
            helping_df[col] = helping_df['service_trip_shift'].map(
                df[df['ticket_type'] == ticket_type].groupby('service_trip_shift')['passenger_count'].sum()
            ).fillna(0)

        helping_df['zero_value_ticket_pass'] = helping_df[['TCH03', 'TCH04', 'TCH06', 'TCH07', 'TCH08']].sum(axis=1)

        helping_df = helping_df.merge(
            df.groupby('service_trip_shift')['stages_travelled'].mean().rename('average_stages'),
            on='service_trip_shift'
        ).round(2)
        helping_df['zero_value_ticket_demand'] = (helping_df['zero_value_ticket_pass'] * helping_df['average_stages']).round(2)
        helping_df = helping_df.merge(df.groupby('service_trip_shift')['KM'].mean().rename('kms'), on='service_trip_shift').round(2)
        helping_df = helping_df.merge(df.groupby('service_trip_shift')['total_stages'].mean().rename('stages'), on='service_trip_shift')
        helping_df = helping_df.merge(df.groupby('service_trip_shift')['total_capacity'].mean().rename('capacity'), on='service_trip_shift').round(2)
        helping_df['supply'] = (helping_df['capacity'] * helping_df['stages']).round(2)

        # Service-level aggregates
        service_help = helping_df.groupby('service_no').agg(
            service_kms=('kms', 'sum'),
            service_supply=('supply', 'sum'),
            zero_value_ticket_demand_service=('zero_value_ticket_demand', 'sum')
        ).reset_index().round(2)

        # Route-level aggregates
        route_help = helping_df.groupby('route_no').agg(
            route_kms=('kms', 'sum'),
            route_supply=('supply', 'sum'),
            zero_value_ticket_demand_route=('zero_value_ticket_demand', 'sum')
        ).reset_index().round(2)

        del helping_df
        gc.collect()

        # Service DataFrame
        service_df = df[["service_no", "service_type"]].drop_duplicates()
        service_df['route_no'] = service_df['service_no'].str.split("-").str[0]
        service_df = service_df[['service_no', 'service_type', 'route_no']]

        duty_date = parse_date(df["duty_date"].iloc[0])
        if duty_date is not None and not pd.isna(duty_date):
            service_df["date"] = duty_date.strftime('%d-%m-%y')
            service_df["day"] = duty_date.strftime('%A')
        else:
            service_df["date"] = str(df["duty_date"].iloc[0])
            service_df["day"] = "Unknown"

        ticket_types = [
            ('TCS01', 'Adult Ticket'),
            ('TCS02', 'Difference Fare Ticket'),
            ('TCS03', 'Disable Person Escort Ticket'),
            ('TCS04', 'Disable Person Ticket'),
            ('TCS05', 'Luggage Ticket'),
            ('TCS06', 'Senior Citizen Ticket'),
            ('TCS07', 'Transgender Ticket'),
            ('TCS08', 'Women Ticket')
        ]

        for col, ticket_type in ticket_types:
            service_df[col] = service_df['service_no'].map(
                df[df['ticket_type'] == ticket_type].groupby('service_no').size()
            ).fillna(0)

        service_df['TTPS'] = service_df[[f'TCS{i:02d}' for i in range(1, 9)]].sum(axis=1)
        service_df = service_df.merge(df.groupby('service_no').size().rename('TTPS_Check'), on='service_no', how='left').fillna(0)

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
                df[df['ticket_type'] == ticket_type].groupby('service_no')['passenger_count'].sum()
            ).fillna(0)

        service_df['PPSPD'] = service_df[
            ['PCS01', 'PCS02', 'PCS03', 'PCS04', 'PCS06', 'PCS07', 'PCS08']
        ].sum(axis=1)

        service_df = service_df.merge(
            df.groupby('service_no')['passenger_count'].sum().rename('PPSPD_Check'),
            on='service_no',
            how='left'
        ).fillna(0)

        service_df['RS01'] = service_df['service_no'].map(
            df[df['ticket_type'] == 'Adult Ticket'].groupby('service_no')['passenger_amount'].sum()
        ).fillna(0)
        service_df['RS02'] = service_df['service_no'].map(
            df[df['ticket_type'] == 'Difference Fare Ticket'].groupby('service_no')['passenger_amount'].sum()
        ).fillna(0)
        service_df['RS03'] = (service_df['PCS03'] * 16).astype(int)
        service_df['RS04'] = (service_df['PCS04'] * 16).astype(int)
        service_df['RS05'] = service_df['service_no'].map(
            df[df['ticket_type'] == 'Luggage Ticket'].groupby('service_no')['luggage_amount'].sum()
        ).fillna(0)
        service_df['RS06'] = (service_df['PCS06'] * 21).astype(int)
        service_df['RS07'] = (service_df['PCS07'] * 16).astype(int)
        service_df['RS08'] = (service_df['PCS08'] * 16).astype(int)

        service_df['EPSPD'] = service_df[[f'RS{i:02d}' for i in range(1, 9)]].sum(axis=1)

        service_metrics = df.groupby("service_no").agg(trips_per_service=('trip_no', 'nunique')).reset_index()
        service_df = service_df.merge(service_metrics, on="service_no", how='left')

        service_demand = df.groupby('service_no')['demand'].sum().reset_index()
        service_df = service_df.merge(service_demand, on='service_no', how='left').fillna(0)
        service_df = pd.merge(service_df, service_help, on='service_no', how='left')
        service_df["service_DSRWOW"] = (service_df["demand"] / service_df["service_supply"]).round(2).fillna(0)
        service_df['total_service_demand'] = (service_df["demand"] + service_df["zero_value_ticket_demand_service"]).round(2)
        service_df['service_DSRWW'] = (service_df['total_service_demand'] / service_df["service_supply"]).round(2)
        service_df = service_df.sort_values('service_DSRWW', ascending=False).reset_index(drop=True)
        sample_size['service_count'] = len(service_df)

        # Route DataFrame
        route_df = df[['route_no']].drop_duplicates()

        duty_date = parse_date(df["duty_date"].iloc[0])
        if duty_date is not None and not pd.isna(duty_date):
            route_df["date"] = duty_date.strftime('%d-%m-%y')
            route_df["day"] = duty_date.strftime('%A')
        else:
            route_df["date"] = str(df["duty_date"].iloc[0])
            route_df["day"] = "Unknown"

        for i, (col, ticket_type) in enumerate(ticket_types, 1):
            route_df[f'TCR{i:02d}'] = route_df['route_no'].map(
                df[df['ticket_type'] == ticket_type].groupby('route_no').size()
            ).fillna(0)

        route_df['TTPR'] = route_df[[f'TCR{i:02d}' for i in range(1, 9)]].sum(axis=1)
        route_df = route_df.merge(df.groupby('route_no').size().rename('TTPR_Check'), on='route_no', how='left').fillna(0)

        route_df['PCR01'] = route_df['route_no'].map(
            df[df['ticket_type'] == 'Adult Ticket'].groupby('route_no')['passenger_count'].sum()
        ).fillna(0)
        route_df['PCR02'] = route_df['route_no'].map(
            df[df['ticket_type'] == 'Difference Fare Ticket'].groupby('route_no')['passenger_count'].sum()
        ).fillna(0)
        route_df['PCR03'] = route_df['route_no'].map(
            df[df['ticket_type'] == 'Disable Person Escort Ticket'].groupby('route_no')['passenger_count'].sum()
        ).fillna(0)
        route_df['PCR04'] = route_df['route_no'].map(
            df[df['ticket_type'] == 'Disable Person Ticket'].groupby('route_no')['passenger_count'].sum()
        ).fillna(0)
        route_df['PCR06'] = route_df['route_no'].map(
            df[df['ticket_type'] == 'Senior Citizen Ticket'].groupby('route_no')['passenger_count'].sum()
        ).fillna(0)
        route_df['PCR07'] = route_df['route_no'].map(
            df[df['ticket_type'] == 'Transgender Ticket'].groupby('route_no')['passenger_count'].sum()
        ).fillna(0)
        route_df['PCR08'] = route_df['route_no'].map(
            df[df['ticket_type'] == 'Women Ticket'].groupby('route_no')['passenger_count'].sum()
        ).fillna(0)

        route_df['PPRPD'] = route_df[['PCR01', 'PCR02', 'PCR03', 'PCR04', 'PCR06', 'PCR07', 'PCR08']].sum(axis=1)
        route_df = route_df.merge(df.groupby('route_no')['passenger_count'].sum().rename('PPRPD_Check'), on='route_no')

        route_df['RR01'] = route_df['route_no'].map(
            df[df['ticket_type'] == 'Adult Ticket'].groupby('route_no')['passenger_amount'].sum()
        ).fillna(0)
        route_df['RR02'] = route_df['route_no'].map(
            df[df['ticket_type'] == 'Difference Fare Ticket'].groupby('route_no')['passenger_amount'].sum()
        ).fillna(0)
        route_df['RR03'] = (route_df['PCR03'] * 16).astype(int)
        route_df['RR04'] = (route_df['PCR04'] * 16).astype(int)
        route_df['RR05'] = route_df['route_no'].map(
            df[df['ticket_type'] == 'Luggage Ticket'].groupby('route_no')['luggage_amount'].sum()
        ).fillna(0)
        route_df['RR06'] = (route_df['PCR06'] * 21).astype(int)
        route_df['RR07'] = (route_df['PCR07'] * 16).astype(int)
        route_df['RR08'] = (route_df['PCR08'] * 16).astype(int)

        route_df['EPRPD'] = route_df[['RR01', 'RR02', 'RR03', 'RR04', 'RR05', 'RR06', 'RR07', 'RR08']].sum(axis=1)

        route_demand = df.groupby('route_no')['demand'].sum().reset_index()
        route_df = route_df.merge(route_demand, on='route_no').fillna(0)
        route_df = pd.merge(route_df, route_help, on='route_no', how='left')
        route_df['route_DSRWOW'] = (route_df['demand'] / route_df['route_supply']).round(2).fillna(0)
        route_df['total_route_demand'] = (route_df["demand"] + route_df["zero_value_ticket_demand_route"]).round(2)
        route_df['route_DSRWW'] = (route_df['total_route_demand'] / route_df['route_supply']).round(2)

        route_df = route_df.sort_values('route_DSRWW', ascending=False).reset_index(drop=True)
        sample_size['route_count'] = len(route_df)

        # Free the big raw table before returning
        del df
        gc.collect()

        return service_df, route_df, sample_size

    except Exception as e:
        st.error(f"Error processing {file_name}: {str(e)}")
        st.text(traceback.format_exc())
        return pd.DataFrame(), pd.DataFrame(), pd.DataFrame()


# Run Analysis Button
if st.button("🚀 Run Analysis", key="run_analysis"):
    items_to_process = available_items if analysis_method == "Process all selected files" else filtered_items

    if not items_to_process:
        st.error("No files available or selected for processing!")
        st.stop()

    progress_bar = st.progress(0)
    status_text = st.empty()

    service_parts, route_parts, sample_parts = [], [], []
    processed_count = 0

    for i, (file_name, file_source) in enumerate(items_to_process):
        status_text.text(f"Processing {i+1}/{len(items_to_process)}: {file_name}")
        progress_bar.progress(int((i + 1) / len(items_to_process) * 100))

        if isinstance(file_source, str) and os.path.exists(file_source) and os.path.getsize(file_source) == 0:
            st.error(f"File '{file_name}' is empty (0 bytes). Please upload a valid non-empty ETM file.")
            continue

        service, route, sample = process_file(file_name, file_source, lookup_df)

        if not service.empty or not route.empty or not sample.empty:
            processed_count += 1
        if not service.empty:
            service_parts.append(service)
        if not route.empty:
            route_parts.append(route)
        if not sample.empty:
            sample_parts.append(sample)

        gc.collect()

    progress_bar.empty()
    status_text.empty()

    all_service = pd.concat(service_parts, ignore_index=True) if service_parts else pd.DataFrame()
    all_route = pd.concat(route_parts, ignore_index=True) if route_parts else pd.DataFrame()
    all_sample = pd.concat(sample_parts, ignore_index=True) if sample_parts else pd.DataFrame()

    if processed_count == 0:
        st.session_state.analysis_complete = False
        st.warning("No valid data was processed. Please check the file error message above.")
    else:
        st.session_state.all_service = all_service
        st.session_state.all_route = all_route
        st.session_state.all_sample = all_sample
        st.session_state.analysis_complete = True
        st.success(f"✅ Analysis completed successfully! Processed {processed_count} of {len(items_to_process)} file(s).")

# Display results if analysis is complete
if st.session_state.analysis_complete:
    has_service = not st.session_state.all_service.empty
    has_route = not st.session_state.all_route.empty
    has_sample = not st.session_state.all_sample.empty

    if has_service or has_route or has_sample:
        st.header("📊 Analysis Results")

        if has_service:
            st.subheader("Service Data")
            if len(st.session_state.all_service) > MAX_ROWS_SHOWN:
                st.caption(f"Showing first {MAX_ROWS_SHOWN:,} of {len(st.session_state.all_service):,} rows. Download the Excel file for all rows.")
            st.dataframe(st.session_state.all_service.head(MAX_ROWS_SHOWN))

        if has_route:
            st.subheader("Route Data")
            if len(st.session_state.all_route) > MAX_ROWS_SHOWN:
                st.caption(f"Showing first {MAX_ROWS_SHOWN:,} of {len(st.session_state.all_route):,} rows. Download the Excel file for all rows.")
            st.dataframe(st.session_state.all_route.head(MAX_ROWS_SHOWN))

        if has_sample:
            st.subheader("Summary Data")
            st.dataframe(st.session_state.all_sample.head(MAX_ROWS_SHOWN))

        timestamp = datetime.now().strftime("%d-%m-%Y_%H%M%S")
        output_excel_name = f"ETM_Analysis_PS_{timestamp}.xlsx"

        try:
            excel_buffer = io.BytesIO()
            with pd.ExcelWriter(excel_buffer, engine='openpyxl') as writer:
                if has_service:
                    st.session_state.all_service.to_excel(writer, sheet_name='Service_Data', index=False)
                if has_route:
                    st.session_state.all_route.to_excel(writer, sheet_name='Route_Data', index=False)
                if has_sample:
                    st.session_state.all_sample.to_excel(writer, sheet_name='Summary_Data', index=False)

            st.download_button(
                label="📥 Download Full Analysis Results (.xlsx)",
                data=excel_buffer.getvalue(),
                file_name=output_excel_name,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
        except Exception as e:
            st.error(f"Error creating Excel report: {e}")

# Sidebar Instructions & Notes
st.sidebar.header("ℹ️ Instructions")
st.sidebar.info("""
1. **Select Data Source**: Test immediately with built-in sample data or upload your own `.csv` / `.bz2` files
2. **Select Analysis Method**:
   - All Files: Process every selected file
   - Date Range: Filter by start/end dates
   - Monthly: Filter by month
   - Weekday/Weekend: Filter Mon-Fri or Sat-Sun
3. **Run Analysis**: Generate Service, Route, and Summary tables
4. **Download**: Export the full multi-sheet Excel report
""")

st.sidebar.header("📝 Notes")
st.sidebar.warning("""
- MTC Fleet Lookup (3,390 vehicles) is built into the app; missing/unmatched fleet numbers default to a capacity of 65
- Works with any filename (`sample_etm_data.csv.bz2`, `sample_etm_data.csv`, etc.)
""")
