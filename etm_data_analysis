import streamlit as st
import pandas as pd
import os
import io
import zlib
import json
import base64
import calendar
import traceback
import numpy as np
from datetime import datetime

# Embedded MTC Fleet Capacity Lookup (from MTC _Fleet_List_08012025.xlsx - Sheet3, 3390 vehicles)
FLEET_LOOKUP_B64 = (
    "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"
    "kcZnOYzWE2hzls5bCVw1YOWzls5bCVw1YOWzlPQRyrqV1JdrjGYRWHVRypd6TekXpH6p1eHtMdbnTkxckSMAZvUs+d1PbceWjPnZ32BHkdpj03"
    "A9pzY7I9d0Dtadpt7m1u6K52V1Nj6ep1gs+/Of7KrTuuTuCphVm6Xh47Vy6PbbRt8Luspvnl4Idtn2Lbp4SGe1WcozIcNeBw+8PtD7c/3P600g"
    "Iatz8s/vDqw6sPrz68+vDqw6sPrz68+vDqw6sPrz6s+LDiw4oPKz6M6XDlw5UPVz5c+bC2w/AOrz588vDJwycPMz/M/DDzwycPnzx88vDJw/gP"
    "4z/F+BnpKb7O7Y8acEr0KRKH6R6me5juYbqH6R6me04xBTT2etjrYa+n2OuRQ4/kfeTQdHUXsymSEzwy9pGxj4x90KL4DlpwgoAIiNBuaDe027"
    "TbtNuKDUQ0LTQtdC10LXQtdC10nex61rXbtZsaS42lxlJjacSp3dRuanfo79Du0NjQ2CiurN2hsVlWXmMLYrlhlRuMYqEttIW2jOLmUAvLHYIg"
    "lK8QJSF2QkSFOAvRF2IyRGqI31A3Q90MGRAlL0q2HKM4xnaM+JiHUqWPiTqm75jUU/INbaNttI3m94Xhv8DGLP/t0r8e9C/sfj93WfHvf73Pzf"
    "n3uVn4Pjd53+dm7PvcFHmfm2/vcxPyfW62vM/Nofe5Wfg+d1fxPtdA3uduO97nJvr7XC95n2sV73O95H3ufud9bs6/z03T97m7lfe5Gfs+1wne"
    "5+bm+9zcfJ+bvO9zNznvc7P7fe5+531uFr7Pzc33ubug97nJ+z43pd/npv/73PR/n5v+73O3de9zTfd9rtO+zzXH97mbvfe57vk+d1v3Ptcn3"
    "+d63/tc73ufu8N7n7tdfJ9rbe9z93Lvczdw73P94X2uP7zP9Yf3uf7wPner9j7XKt7nWsX7XFN4n+sa73Nd433uvu997m7wfa7hvc91ufe5m8"
    "j3uZvI97m7wfe5Re19bg69z82s97n59j53n/o+dyP7Ptd03+c67fvcjez73N32+1xPfZ/rtO9zs/t9bs6/z3WC97n+8D7XNd7n2vb73Cr9PtfX"
    "3+fu7d/n7orf5+6K3+e60ftcj3qf61zvc/3sfa7Lvc/1vvcJrhGyMARtCNqQ6CF+Q3bHU9q9Ywu+E9I/xHqwqxD2IS8+v5DpHxlXyotgNsHEQr"
    "YElwsOE/wseEkwseByIbNCZoUsDFkY7PXzO5+/XU9ykJNc5CbPlb929VeahzQ7iZZoiZZoiZZoiTbQBtpAG2gDbaANiAE8y2NoE21CTOAJMbU7"
    "0ZZ2pwFNiAWxtLvQlk6O28LnY4q/N6BN9y60pWflsQ28gTfw1vVtmBtto22IDbzRDtpBO2gH7Vza3Ru9z90x/aNQ/aPO//PZVfz5O9k/n+L+R7"
    "6f/7L7aaGpvE2Nbap0U6WbEtqU0KaENkW4qbxNPW4KdlPGm1LXWHxj/E3FaYy/KT6NFbdRxuYxe6Nmx9RsXJqNS7MDaepxm0UWsAGpDE3FaepQ"
    "U9SamtXUrM+PAv4utzA6wugIo3PD6PPLvL8yyEZ2MslJDvLSPj84+3t1k8CBFmiBFmhxEcn7kvdllhsmuchN3j4k70vel7wveV/yvuR9n9+e/Z"
    "WGGYbZzEMz4lYQ+stIkzEl90w22A2+m5LPb8T+9sFaNJ1sRsGKk1fnLDeYEl79+bnYX2nEvQyzgG+7nXN1LteZWOdcnV111ta5fZ/lqhZMVC++"
    "blK7qe6WsFvjLgi6IOiipAujLri6kOsCsStfXfHpZQlNaldxcpW1sPKrrIWAMSWJlgpg1nY9tqxx6Zk61JW6rg519a2rb32XFspyQxSpWKYWUr"
    "upD6mxhMhVhmnwOpmKZT9lAUiT2s1DN6ndAvRZbnDV/HaT2ivYvVy5nxJ9JXlLd8qUmCgZkDVbyuBLOnmMrydfT76efD35ej5lo+XqKY8JoyhX"
    "C1jPlIPk9qkyJItPRposM1t5TGhEadcoiqfy3yyu7LTYnCGbk2VzhmyOns1BrDnzNmfe5qjcVqnd7j1F6oPjWX8eMshGdnKQi7RxcYhvDoPpLJ"
    "LOIukAkg4g6QCSDiDpAJKOZ9nLVWvhEJS9rCaEIpGtrLzH1JZU1KbGZumDMjMZ9GTQk0FPBj2V0JmlhUYGqb55DdB22cCVBSgBU+4VGqdI956y"
    "bjcepvmd5nea3+aVTFslqu0nvfVp3h60U8LoXu1RAnGSd5jJ2vIUf7CEDC+L/zoSdAeF7vjQ7ZW7fXW32+5OHd1uu9ukdxvv3op0r7N/b2XEwF"
    "5FdC8dutcWPcpjBuQVR29FusFbie5dRfcGo7ciyzyUFspEPdbixvrg9oPFj2L8pTL0khf36rQhmrZJ095o2htNG6JpmzQVtWkzMm1RJt+ZpQ98"
    "Z/aS6B5TmoeaNRSq8ZQbSh9KJ0tjBjTLvWUUpetGPMvsHLRFlquWRakbCuBQN4fyNdTCodTNVaT+2qJMG7hp/zBtUab9zlBj5y5Su7Zf0/Zr7o"
    "LwmC3VtBOb9nLTZm/uMvjSHdOn5I9W7tUus5nMZnKjyXeGMjNaQVgWZWYoHcP+YbTSWJmo0h1Tbe857c+mHdO0lxsCfHDwwdeH+jbUt1FqgDyep"
    "4CNTZItprCYwmIKq6SThBxK3chy7yIPeQe0otA2WbpjxGrs6GV2jEJpXhJnyZZVckhxH4r7yDIgw2R4w05hMLHB2gbDG6OAS8/K4A1TVC/xu4T"
    "9KrHOMpeQWzJgicnVyr0GxLYH2x5MbJS8EDtDDo1T7tWufBuCa5YyU2xwlqvlXkErzmYxaIE4hdEUMFPIzae0KwOiSDdYgFkMr3iqJZwiaoqoa"
    "YWm6JvFzFWRob4NVW8U01VbhoozFJShzAylbgqC2UrP9Fc8zGJXysFQJIbSMVSGoV4MZj5WaUxUtxK/ApxdrV6C1lq4d9ar1q0VhJhUh8YuVw1"
    "IzRqq0yhlZpdhWguFahRX3uVqSQaN8fXB11cv8yDRWfzq5V7Jy+0X71sccRXTzTK/JoqfrfoWsN339lFe7P/uHP/v/wGWR7LH"
)


@st.cache_data(show_spinner=False)
def get_embedded_fleet_lookup():
    """Loads the embedded MTC fleet capacity lookup table."""
    raw = zlib.decompress(base64.b64decode(FLEET_LOOKUP_B64)).decode("utf-8")
    mapping = json.loads(raw)
    return pd.DataFrame({
        "fleet_no": list(mapping.keys()),
        "total_capacity": list(mapping.values())
    })


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

# Load embedded fleet lookup
lookup_df = get_embedded_fleet_lookup()

# Sidebar Configuration
with st.sidebar:
    st.header("⚙️ Configuration")
    
    # Path Configuration (Lookup file is now built-in!)
    input_folder = st.text_input("Input Folder", r"please paste the path that contains Etm data in Csv or bz2 format")
    output_folder = st.text_input("Output Folder", r"please paste path of output folder")
    st.caption(f"✅ Fleet lookup built-in ({len(lookup_df):,} vehicles loaded; default capacity = 65)")
    
    # Verify paths button
    if st.button("Verify Paths", key="verify_paths"):
        errors = []
        valid = True
        
        if not os.path.exists(input_folder):
            errors.append(f"Input folder not found: {input_folder}")
            valid = False
        else:
            csv_files = [f for f in os.listdir(input_folder) if f.lower().endswith(('.csv', '.bz2'))]
            if not csv_files:
                errors.append(f"No CSV/BZ2 files found in input folder: {input_folder}")
                valid = False
        
        if output_folder and output_folder != r"please paste path of output folder":
            os.makedirs(output_folder, exist_ok=True)
        
        if not valid:
            st.error("Errors found in configuration:")
            for error in errors:
                st.error(error)
        else:
            st.success("All paths are valid!")


# Function to parse dates
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


@st.cache_data(show_spinner=False)
def get_file_dates_map(folder_path, filenames):
    """Maps {filename: datetime.date} irrespective of filename format or .bz2 compression."""
    date_map = {}
    for file in filenames:
        dt = None
        # Step 1: Fast check if date is embedded in filename
        try:
            if "Ticket-Data" in file:
                date_part = file.split('-')[2:5]
                date_str = '-'.join(date_part).split('.')[0]
                dt = datetime.strptime(date_str, '%Y-%m-%d').date()
            elif '_' in file:
                date_str = file.split('_')[1].split('.')[0]
                if len(date_str) == 8 and date_str.isdigit():
                    dt = datetime.strptime(date_str, '%d%m%Y').date()
        except Exception:
            dt = None

        # Step 2: Fallback — inspect first row inside the file using compression='infer'
        if dt is None:
            try:
                file_path = os.path.join(folder_path, file)
                sample_df = pd.read_csv(file_path, compression='infer', nrows=1)
                date_col = 'date' if 'date' in sample_df.columns else ('duty_date' if 'duty_date' in sample_df.columns else None)
                if date_col and not sample_df.empty:
                    parsed = parse_date(sample_df[date_col].iloc[0])
                    if parsed is not None and not pd.isna(parsed):
                        dt = parsed.date()
            except Exception:
                pass

        if dt is not None:
            date_map[file] = dt

    return date_map


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
    all_files = [f for f in os.listdir(input_folder) if f.lower().endswith(('.csv', '.bz2'))]
    st.info(f"Found {len(all_files)} data files in input folder")

# Build date map only when needed for date-based filtering
file_date_map = {}
if all_files and analysis_method != "Process all files in input folder":
    file_date_map = get_file_dates_map(input_folder, all_files)

# Date range selection
date_files = []
if analysis_method == "Process files by date range" and all_files:
    st.subheader("📅 Select Date Range")
    if file_date_map:
        min_date = min(file_date_map.values())
        max_date = max(file_date_map.values())
        
        col1, col2 = st.columns(2)
        with col1:
            start_date = st.date_input("Start Date", min_date)
        with col2:
            end_date = st.date_input("End Date", max_date)
        
        date_files = [f for f, dt in file_date_map.items() if start_date <= dt <= end_date]
        st.info(f"Found {len(date_files)} files in selected date range")
    else:
        st.warning("No valid dates found in the input files.")

# Monthly analysis selection
month_files = []
if analysis_method == "Process files by month" and all_files:
    st.subheader("🗓️ Select Month")
    if file_date_map:
        month_year_options = {}
        for file, dt in file_date_map.items():
            month_year = f"{calendar.month_name[dt.month]} {dt.year}"
            month_year_options.setdefault(month_year, []).append(file)
        
        sorted_months = sorted(month_year_options.keys(), key=lambda x: datetime.strptime(x, '%B %Y'))
        selected_month = st.selectbox("Available Months", sorted_months)
        month_files = month_year_options.get(selected_month, [])
        st.info(f"Found {len(month_files)} files for {selected_month}")
    else:
        st.warning("No valid monthly data found in the input files.")

# Weekday/Weekend analysis selection
day_type_files = []
if analysis_method == "Process files by weekday/weekend" and all_files:
    st.subheader("📅 Select Day Type")
    weekday_files = [f for f, dt in file_date_map.items() if dt.weekday() < 5]
    weekend_files = [f for f, dt in file_date_map.items() if dt.weekday() >= 5]
    
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


# Main processing function
def process_file(file_path, lookup_df):
    try:
        # Read the CSV/BZ2 file using compression='infer' (works regardless of file extension)
        try:
            df = pd.read_csv(file_path, compression='infer')
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
            split_result = df['schedule_no'].astype(str).str.split('-', expand=True)
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
        df = df[df['KM'].ne(0) | df['KM'].isna()].copy()

        # Total Stages per each route
        df['route_description'] = df['Trip_Origin'].astype(str) + " To " + df['Trip_Destination'].astype(str)
        df['route_desc'] = df['route_no'].astype(str) + "-" + df['route_description'].astype(str)
        
        # Calculate total stages correctly
        df['max_stage'] = df[['to_stage', 'from_stage']].max(axis=1)
        df['total_stages'] = df.groupby('route_desc')['max_stage'].transform('max')
        df.drop(columns=['max_stage'], inplace=True)

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
        
        # Aggregate metrics
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

        # Service DataFrame
        service_df = df[["service_no", "service_type"]].drop_duplicates()
        service_df['route_no'] = service_df['service_no'].str.split("-").str[0]
        service_df = service_df[['service_no', 'service_type', 'route_no']]
        
        # Parse date for service_df
        duty_date = parse_date(df["duty_date"].iloc[0])
        if duty_date is not None and not pd.isna(duty_date):
            service_df["date"] = duty_date.strftime('%d-%m-%y')
            service_df["day"] = duty_date.strftime('%A')
        else:
            service_df["date"] = str(df["duty_date"].iloc[0])
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
            ('TCS08', 'Women Ticket')
        ]
        
        for col, ticket_type in ticket_types:
            service_df[col] = service_df['service_no'].map(
                df[df['ticket_type'] == ticket_type].groupby('service_no').size()
            ).fillna(0)
        
        service_df['TTPS'] = service_df[[f'TCS{i:02d}' for i in range(1, 9)]].sum(axis=1)
        service_df = service_df.merge(df.groupby('service_no').size().rename('TTPS_Check'), on='service_no', how='left').fillna(0)

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

        # Revenue per service
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
        
        # Service metrics
        service_metrics = df.groupby("service_no").agg(trips_per_service=('trip_no', 'nunique')).reset_index()
        service_df = service_df.merge(service_metrics, on="service_no", how='left')

        # Service Demand and supply
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
        
        # Parse date for route_df
        duty_date = parse_date(df["duty_date"].iloc[0])
        if duty_date is not None and not pd.isna(duty_date):
            route_df["date"] = duty_date.strftime('%d-%m-%y')
            route_df["day"] = duty_date.strftime('%A')
        else:
            route_df["date"] = str(df["duty_date"].iloc[0])
            route_df["day"] = "Unknown"

        # Ticket Count for route
        for i, (col, ticket_type) in enumerate(ticket_types, 1):
            route_df[f'TCR{i:02d}'] = route_df['route_no'].map(
                df[df['ticket_type'] == ticket_type].groupby('route_no').size()
            ).fillna(0)
            
        route_df['TTPR'] = route_df[[f'TCR{i:02d}' for i in range(1, 9)]].sum(axis=1)
        route_df = route_df.merge(df.groupby('route_no').size().rename('TTPR_Check'), on='route_no', how='left').fillna(0)

        # Passenger count per route
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

        # Total no of passengers per Route
        route_df['PPRPD'] = route_df[['PCR01', 'PCR02', 'PCR03', 'PCR04', 'PCR06', 'PCR07', 'PCR08']].sum(axis=1)
        route_df = route_df.merge(df.groupby('route_no')['passenger_count'].sum().rename('PPRPD_Check'), on='route_no')

        # Revenue for Route
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

        # Demand and Supply Ratio Calculations
        route_demand = df.groupby('route_no')['demand'].sum().reset_index()
        route_df = route_df.merge(route_demand, on='route_no').fillna(0)
        route_df = pd.merge(route_df, route_help, on='route_no', how='left')
        route_df['route_DSRWOW'] = (route_df['demand'] / route_df['route_supply']).round(2).fillna(0)
        route_df['total_route_demand'] = (route_df["demand"] + route_df["zero_value_ticket_demand_route"]).round(2)
        route_df['route_DSRWW'] = (route_df['total_route_demand'] / route_df['route_supply']).round(2)
        
        route_df = route_df.sort_values('route_DSRWW', ascending=False).reset_index(drop=True)
        sample_size['route_count'] = len(route_df)

        return service_df, route_df, sample_size

    except Exception as e:
        st.error(f"Error processing {os.path.basename(file_path)}: {str(e)}")
        st.text(traceback.format_exc())
        return pd.DataFrame(), pd.DataFrame(), pd.DataFrame()


# Run Analysis Button
if st.button("🚀 Run Analysis", key="run_analysis"):
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
    
    if not os.path.exists(input_folder):
        st.error(f"Input folder not found: {input_folder}")
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
    
    # Export results to Excel buffer (and save to disk if output folder is set)
    timestamp = datetime.now().strftime("%d-%m-%Y_%H%M%S")
    output_excel_name = f"ETM_Analysis_PS_{timestamp}.xlsx"
    
    try:
        excel_buffer = io.BytesIO()
        with pd.ExcelWriter(excel_buffer, engine='openpyxl') as writer:
            if not st.session_state.all_service.empty:
                st.session_state.all_service.to_excel(writer, sheet_name='Service_Data', index=False)
            if not st.session_state.all_route.empty:
                st.session_state.all_route.to_excel(writer, sheet_name='Route_Data', index=False)
            if not st.session_state.all_sample.empty:
                st.session_state.all_sample.to_excel(writer, sheet_name='Summary_Data', index=False)
        excel_bytes = excel_buffer.getvalue()
        
        if output_folder and output_folder != r"please paste path of output folder":
            os.makedirs(output_folder, exist_ok=True)
            output_path = os.path.join(output_folder, output_excel_name)
            with open(output_path, "wb") as f:
                f.write(excel_bytes)
            st.session_state.output_path = output_path
            st.info(f"Results saved to: {output_path}")
        
        st.download_button(
            label="📥 Download Full Analysis Results",
            data=excel_bytes,
            file_name=output_excel_name,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
            
    except Exception as e:
        st.error(f"Error creating Excel report: {e}")

# Instructions Section
st.sidebar.header("ℹ️ Instructions")
st.sidebar.info("""
1. **Configure Paths**: Set input and output folders
2. **Select Analysis Method**:
  - All Files: Process every CSV/BZ2 in input folder
  - Date Range: Select start/end dates
  - Monthly: Choose month from dropdown
  - Weekday/Weekend: Analyze Mon-Fri or Sat-Sun
3. **Verify Paths**: Check configuration is correct
4. **Run Analysis**: Process files and generate results
5. **Download**: Get Excel report when complete
""")

st.sidebar.header("📝 Notes")
st.sidebar.warning("""
- MTC Fleet Lookup (3,390 vehicles) is built into the app; missing/unmatched fleet numbers default to a capacity of 65
- Works with any filename (.csv or .bz2), including renamed files
- Output folder will be created automatically if specified
""")
