import streamlit as st
import pandas as pd
import datetime as dt
import random
from io import BytesIO
import zipfile
import re
try:
    from vip_config import VIP_CONFIG
except Exception:
    VIP_CONFIG = {}
from reportlab.lib import colors
from reportlab.lib.pagesizes import letter, landscape, A4
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.enums import TA_CENTER, TA_LEFT, TA_RIGHT

st.set_page_config(page_title="Laureate Group Attendance System - Multi-Campus", layout="wide")
st.title("📊 Multi-Campus Staff Attendance & Payroll Analyzer")

# --- INITIALIZE SESSION STATE ---
if 'batch_results' not in st.session_state:
    st.session_state.batch_results = {}  # {campus_name: {'summary_df': ..., 'sorted_df': ..., 'valid_working_dates': ..., 'file_name': ...}}
if 'approved_leaves_df' not in st.session_state:
    st.session_state.approved_leaves_df = pd.DataFrame(columns=["Employee_ID", "Date"])
if 'custom_date_timings' not in st.session_state:
    st.session_state.custom_date_timings = []
if 'no_lc_dates' not in st.session_state:
    st.session_state.no_lc_dates = []

# --- STYLING FUNCTIONS ---
def style_attendance(val):
    val_str = str(val)
    if 'Absent' in val_str: return 'background-color: #ffcccc; color: #990000; font-weight: bold;'
    if 'Sandwich' in val_str: return 'background-color: #e06666; color: white; font-weight: bold;'
    if 'Approved' in val_str: return 'background-color: #c9daf8; color: #1155cc; font-weight: bold;'
    if 'Short Leave' in val_str: return 'background-color: #f4cccc; color: #cc0000; font-weight: bold;'
    if 'Double LC' in val_str: return 'background-color: #f9cb9c; color: #b45f06; font-weight: bold;'
    if 'LC' in val_str or 'Late Coming' in val_str: return 'background-color: #fff2cc; color: #b45f06; font-weight: bold;'
    if 'Missing Out' in val_str: return 'background-color: #fce5cd; color: #e69138; font-weight: bold;'
    if 'Left Early' in val_str: return 'background-color: #e8eaed; color: #5f6368;'
    if 'Sunday' in val_str or 'Holiday' in val_str: return 'background-color: #d9d9d9; color: #434343; font-style: italic;'
    if 'On Time' in val_str: return 'background-color: #d9ead3; color: #274e13;'
    return ''

def apply_style(styler, style_func, subset_cols):
    valid_cols = [c for c in subset_cols if c in styler.columns]
    if not valid_cols:
        return styler
    if hasattr(styler, 'map'):
        return styler.map(style_func, subset=valid_cols)
    return styler.applymap(style_func, subset=valid_cols)

# --- SUMMARY EXPORT BUILDER FUNCTION ---
def build_summary_export_df(dataframe):
    df_export = pd.DataFrame()
    df_export['Sr.'] = range(1, len(dataframe) + 1)
    df_export['Name'] = dataframe['Name'].values
    df_export['WD'] = dataframe['Work Days'].values
    df_export['LCs'] = dataframe['LCs'].values
    df_export['Double LCs'] = dataframe['Double LCs'].values
    df_export['Short Leaves'] = dataframe['Short Leaves'].values
    df_export['Missing Out Time'] = dataframe['Missing Out Time'].values
    df_export['Full Leaves'] = dataframe['Full Leaves'].values
    # Blank for manual entry by coordinators (as requested)
    df_export['Remit/Reason'] = ""
    return df_export

# --- EXCEL GENERATOR FUNCTIONS ---
def generate_excel(dataframe):
    output = BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        dataframe.to_excel(writer, index=False, sheet_name='Report Data')
    return output.getvalue()

def generate_summary_excel(dataframe):
    export_df = build_summary_export_df(dataframe)
    output = BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        export_df.to_excel(writer, index=False, sheet_name='Summary')
    return output.getvalue()

# --- SUMMARY PDF GENERATOR (A4 Portrait, Exact Fit to Print & Repeating Headings) ---
def generate_summary_pdf(dataframe, campus, month_str=""):
    buffer = BytesIO()
    # A4 Portrait: width = 595.28 pt, height = 841.89 pt. 
    # Left & Right margins = 18 pt each -> Total printable width = 595.28 - 36 = 559.28 pt
    doc = SimpleDocTemplate(
        buffer, 
        pagesize=A4, 
        leftMargin=18, 
        rightMargin=18, 
        topMargin=20, 
        bottomMargin=20
    )
    styles = getSampleStyleSheet()
    title_style = ParagraphStyle(
        name='MainTitlePortrait', 
        parent=styles['Title'], 
        alignment=TA_CENTER, 
        fontName='Helvetica-Bold',
        fontSize=13, 
        leading=16,
        textColor=colors.HexColor("#1155cc"),
        spaceAfter=2
    )
    subtitle_style = ParagraphStyle(
        name='SubTitlePortrait', 
        parent=styles['Normal'], 
        alignment=TA_CENTER, 
        fontName='Helvetica-Bold',
        fontSize=9.5, 
        leading=12,
        textColor=colors.HexColor("#333333"),
        spaceAfter=10
    )
    
    header_title = f"Laureate Group of Schools ({campus})"
    sub_title = "Monthly Staff Attendance & Payroll Summary Report"
    if month_str:
        sub_title += f" — {month_str}"
        
    elements = [
        Paragraph(f"<b>{header_title}</b>", title_style),
        Paragraph(sub_title, subtitle_style)
    ]
    
    df_export = build_summary_export_df(dataframe)
    
    # Exact Column widths summing to 558 pt (Exact fit for A4 Portrait 559 pt printable width)
    # Sr. (24), Name (126), WD (26), LCs (26), Double LCs (44), Short Leaves (50), Missing Out Time (56), Full Leaves (46), Remit/Reason (140)
    col_widths = [24, 126, 26, 26, 44, 50, 56, 46, 140]
    
    cell_style_left = ParagraphStyle('CellLeftP', fontName='Helvetica', fontSize=7.5, leading=9, alignment=TA_LEFT)
    cell_style_center = ParagraphStyle('CellCenterP', fontName='Helvetica', fontSize=7.5, leading=9, alignment=TA_CENTER)
    header_style = ParagraphStyle('HdrStyleP', fontName='Helvetica-Bold', fontSize=7.5, leading=9, alignment=TA_CENTER, textColor=colors.whitesmoke)
    
    headers = ['Sr.', 'Name', 'WD', 'LCs', 'Double LCs', 'Short Leaves', 'Missing Out Time', 'Full Leaves', 'Remit/Reason']
    table_data = [[Paragraph(h, header_style) for h in headers]]
    
    for row in df_export.itertuples(index=False):
        row_cells = [
            Paragraph(str(row[0]), cell_style_center), # Sr.
            Paragraph(str(row[1]), cell_style_left),   # Name
            Paragraph(str(row[2]), cell_style_center), # WD
            Paragraph(str(row[3]), cell_style_center), # LCs
            Paragraph(str(row[4]), cell_style_center), # Double LCs
            Paragraph(str(row[5]), cell_style_center), # Short Leaves
            Paragraph(str(row[6]), cell_style_center), # Missing Out Time
            Paragraph(str(row[7]), cell_style_center), # Full Leaves
            Paragraph("", cell_style_left)             # Remit/Reason (Blank for manual entry by coordinator)
        ]
        table_data.append(row_cells)
        
    t_style = [
        ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor("#1155cc")),
        ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
        ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
        ('GRID', (0, 0), (-1, -1), 0.5, colors.HexColor("#aaaaaa")),
        ('TOPPADDING', (0, 0), (-1, -1), 3),
        ('BOTTOMPADDING', (0, 0), (-1, -1), 3),
        ('LEFTPADDING', (0, 0), (-1, -1), 2.5),
        ('RIGHTPADDING', (0, 0), (-1, -1), 2.5),
    ]
    
    # Alternating row background for clean readability & highlight leaves
    for i, row in enumerate(df_export.itertuples(index=False)):
        row_idx = i + 1
        full_lvs = row[7]
        if full_lvs > 0:
            t_style.append(('BACKGROUND', (0, row_idx), (-1, row_idx), colors.HexColor("#fff0f0")))
        elif i % 2 == 1:
            t_style.append(('BACKGROUND', (0, row_idx), (-1, row_idx), colors.HexColor("#f8f9fa")))
            
    # repeatRows=1 repeats the table header cleanly on every new page without overlapping!
    table = Table(table_data, colWidths=col_widths, repeatRows=1)
    table.setStyle(TableStyle(t_style))
    elements.append(table)
    
    doc.build(elements)
    return buffer.getvalue()

# --- MASTER DETAILED PDF GENERATOR ---
def generate_pdf(dataframe, report_title, campus):
    buffer = BytesIO()
    doc = SimpleDocTemplate(buffer, pagesize=landscape(A4), topMargin=22, bottomMargin=22, leftMargin=20, rightMargin=20)
    styles = getSampleStyleSheet()
    title_style = ParagraphStyle(name='CenterTitle', parent=styles['Title'], alignment=TA_CENTER, fontSize=16, spaceAfter=4)
    subtitle_style = ParagraphStyle(name='CenterSubtitle', parent=styles['Heading2'], alignment=TA_CENTER, spaceAfter=15)
    
    elements = [Paragraph(f"<b>Laureate Group of Schools ({campus})</b>", title_style), Paragraph(report_title, subtitle_style)]
    clean_df = dataframe.fillna("").astype(str)
    data = [clean_df.columns.tolist()] + clean_df.values.tolist()
    
    table_style = [
        ('BACKGROUND', (0,0), (-1,0), colors.HexColor("#1155cc")),
        ('TEXTCOLOR', (0,0), (-1,0), colors.whitesmoke),
        ('ALIGN', (0,0), (-1,-1), 'CENTER'),
        ('FONTNAME', (0,0), (-1,0), 'Helvetica-Bold'),
        ('GRID', (0,0), (-1,-1), 0.5, colors.grey),
        ('FONTSIZE', (0,0), (-1,-1), 7.5),
        ('TOPPADDING', (0,0), (-1,-1), 2.5),
        ('BOTTOMPADDING', (0,0), (-1,-1), 2.5),
    ]
    
    if 'Status' in dataframe.columns:
        for i, row in enumerate(clean_df.values):
            s = str(row[-1])
            bg = colors.white
            if 'Absent' in s: bg = colors.HexColor("#ffcccc")
            elif 'Sandwich' in s: bg = colors.HexColor("#e06666")
            elif 'Approved' in s: bg = colors.HexColor("#c9daf8")
            elif 'Short Leave' in s: bg = colors.HexColor("#f4cccc")
            elif 'Double LC' in s: bg = colors.HexColor("#f9cb9c")
            elif 'LC' in s or 'Late' in s: bg = colors.HexColor("#fff2cc")
            elif 'Missing' in s: bg = colors.HexColor("#fce5cd")
            elif 'Sunday' in s or 'Holiday' in s: bg = colors.HexColor("#d9d9d9")
            elif 'On Time' in s: bg = colors.HexColor("#d9ead3")
            table_style.append(('BACKGROUND', (0, i+1), (-1, i+1), bg))

    table = Table(data, repeatRows=1)
    table.setStyle(TableStyle(table_style))
    elements.append(table)
    doc.build(elements)
    return buffer.getvalue()

def extract_campus_name(filename):
    clean = re.sub(r'\.(xlsx|xls)$', '', filename, flags=re.IGNORECASE)
    clean = re.sub(r'[\-_]', ' ', clean)
    clean = re.sub(r'\b(october|oct|november|nov|december|dec|january|jan|february|feb|march|mar|april|apr|may|june|jun|july|jul|august|aug|september|sep|sept|report|transaction|attendance|monthly|biometric|data|sheet|202[4-9]|copy|\(\d+\))\b', '', clean, flags=re.IGNORECASE)
    clean = re.sub(r'\s+', ' ', clean).strip()
    return clean.title() if clean else filename

# --- SIDEBAR CONFIGURATION ---
st.sidebar.header("⚙️ Standard Shift Rules")
shift_in = st.sidebar.time_input("Standard In Time", dt.time(7, 30))
shift_out = st.sidebar.time_input("Standard Out Time", dt.time(14, 0))
grace_mins = st.sidebar.number_input("Grace Period (Minutes)", min_value=0, max_value=60, value=5, help="Arrivals within this limit are 'On Time'.")

st.sidebar.header("🕌 Friday Timing Rules")
friday_shift_in = st.sidebar.time_input("Friday In Time", dt.time(7, 30))
friday_shift_out = st.sidebar.time_input("Friday Out Time", dt.time(12, 30))

st.sidebar.header("🔄 Mid-Month Shift Change")
shift2_date = st.sidebar.date_input("Shift 2 Effective Date", value=None)
shift2_in = st.sidebar.time_input("Shift 2 In Time", dt.time(8, 0))
shift2_out = st.sidebar.time_input("Shift 2 Out Time", dt.time(14, 30))
shift2_friday_in = st.sidebar.time_input("Shift 2 Friday In Time", dt.time(8, 0))
shift2_friday_out = st.sidebar.time_input("Shift 2 Friday Out Time", dt.time(12, 30))

st.sidebar.header("📅 Specific Date Overrides")
with st.sidebar.expander("Add Specific Date Timing", expanded=False):
    c_date_input = st.date_input("Override Date", value=None, key="override_date_picker")
    c_col_in, c_col_out = st.columns(2)
    with c_col_in:
        c_in_input = st.time_input("In Time", dt.time(8, 0), key="override_in_time")
    with c_col_out:
        c_out_input = st.time_input("Out Time", dt.time(13, 0), key="override_out_time")
    
    if st.button("➕ Add Override", use_container_width=True):
        if c_date_input is not None:
            existing = [item for item in st.session_state.custom_date_timings if item['date'] == c_date_input]
            if existing:
                existing[0]['in'] = c_in_input
                existing[0]['out'] = c_out_input
                st.success(f"Updated timing for {c_date_input.strftime('%d-%b-%Y')}")
            else:
                st.session_state.custom_date_timings.append({
                    'date': c_date_input,
                    'in': c_in_input,
                    'out': c_out_input
                })
                st.success(f"Added override for {c_date_input.strftime('%d-%b-%Y')}")
            st.rerun()

    if st.session_state.custom_date_timings:
        st.markdown("**Active Overrides:**")
        to_delete = None
        for i, item in enumerate(st.session_state.custom_date_timings):
            c_row1, c_row2 = st.columns([3, 1])
            with c_row1:
                st.caption(f"📌 **{item['date'].strftime('%d-%b-%Y')}**: {item['in'].strftime('%I:%M %p')} - {item['out'].strftime('%I:%M %p')}")
            with c_row2:
                if st.button("❌", key=f"del_override_{i}"):
                    to_delete = i
        if to_delete is not None:
            st.session_state.custom_date_timings.pop(to_delete)
            st.rerun()

st.sidebar.header("🚫 No LC Dates (For All Staff)")
with st.sidebar.expander("Select No LC Dates", expanded=False):
    no_lc_date_input = st.date_input("Pick Date(s) — No LC will apply for everyone", value=None, key="no_lc_date_picker")
    if st.button("➕ Add No LC Date", use_container_width=True, key="add_no_lc_btn"):
        if no_lc_date_input is not None:
            if no_lc_date_input not in st.session_state.no_lc_dates:
                st.session_state.no_lc_dates.append(no_lc_date_input)
                st.success(f"Added No LC for {no_lc_date_input.strftime('%d-%b-%Y')}")
            else:
                st.info(f"{no_lc_date_input.strftime('%d-%b-%Y')} is already in the list.")
            st.rerun()
    
    if st.session_state.no_lc_dates:
        st.markdown("**Active No LC Dates:**")
        no_lc_to_delete = None
        for i, nlc_date in enumerate(sorted(st.session_state.no_lc_dates)):
            nlc_col1, nlc_col2 = st.columns([3, 1])
            with nlc_col1:
                st.caption(f"🚫 **{nlc_date.strftime('%d-%b-%Y (%A)')}**")
            with nlc_col2:
                if st.button("❌", key=f"del_no_lc_{i}"):
                    no_lc_to_delete = nlc_date
        if no_lc_to_delete is not None:
            st.session_state.no_lc_dates.remove(no_lc_to_delete)
            st.rerun()

st.sidebar.header("🛡️ Approved Staff Leaves")
st.sidebar.write("Add IDs and Dates for approved leave. *Acts as a shield against the Sandwich rule.*")
edited_leaves_df = st.sidebar.data_editor(
    st.session_state.approved_leaves_df, 
    num_rows="dynamic",
    column_config={
        "Date": st.column_config.DateColumn("Date", required=True),
        "Employee_ID": st.column_config.TextColumn("Employee ID", required=True)
    }
)

st.sidebar.header("🏖️ Extra Off Days")
extra_off_days_count = st.sidebar.number_input("Extra Off Days (Raw Number)", min_value=0, value=0)

def get_vip_ids_for_campus(campus_name):
    merged_config = dict(VIP_CONFIG)
    try:
        if hasattr(st, "secrets") and "VIP_CONFIG" in st.secrets:
            merged_config.update(dict(st.secrets["VIP_CONFIG"]))
    except Exception:
        pass
        
    if not merged_config:
        return []
        
    general_ids = []
    for g_key in ["General", "general", "All", "all"]:
        if g_key in merged_config:
            general_ids.extend([str(x).strip().replace(".0", "") for x in merged_config[g_key]])
            
    matched_ids = list(general_ids)
    current_lower = campus_name.strip().lower()
    
    for c_key, c_ids in merged_config.items():
        k_lower = c_key.strip().lower()
        if k_lower in ["general", "all"]:
            continue
        if k_lower in current_lower or current_lower in k_lower:
            matched_ids.extend([str(x).strip().replace(".0", "") for x in c_ids])
            
    return list(set(matched_ids))

# --- SINGLE CAMPUS PROCESSOR FUNCTION ---
def process_single_file(file_content, campus_name, global_config):
    df_raw = pd.read_excel(BytesIO(file_content), header=None)
    
    is_new_format = False
    col0_val = str(df_raw.iloc[0, 0]).strip().lower()
    if 'emp id' in col0_val or 'employee' in col0_val:
        is_new_format = True

    if is_new_format:
        if len(df_raw.columns) >= 4 and 'time' in str(df_raw.iloc[0, 3]).strip().lower():
            df_clean = df_raw.iloc[1:, [0, 1, 2, 3]].copy()
            df_clean.columns = ['Employee_ID', 'Full_Name', 'Date_Str', 'Time_Str']
            df_clean['DateTime'] = pd.to_datetime(df_clean['Date_Str'].astype(str) + ' ' + df_clean['Time_Str'].astype(str), errors='coerce')
            df_clean = df_clean.drop(columns=['Date_Str', 'Time_Str'])
        else:
            df_clean = df_raw.iloc[1:, [0, 1, 2]].copy()
            df_clean.columns = ['Employee_ID', 'Full_Name', 'DateTime']
    else:
        df_clean = df_raw.iloc[8:, [1, 2, 5]].copy()
        df_clean.columns = ['Employee_ID', 'Full_Name', 'DateTime']
    
    df_clean['Employee_ID'] = df_clean['Employee_ID'].astype(str).str.strip().str.replace(r'\.0$', '', regex=True)
    df_clean = df_clean.dropna(subset=['DateTime'])

    extracted_records = []
    for _, row in df_clean.iterrows():
        if pd.api.types.is_datetime64_any_dtype(row['DateTime']) or isinstance(row['DateTime'], pd.Timestamp):
            if pd.notnull(row['DateTime']):
                extracted_records.append({'Employee_ID': row['Employee_ID'], 'Full_Name': row['Full_Name'], 'DateTime': row['DateTime']})
        else:
            dt_val = str(row['DateTime']).strip()
            punches = re.split(r'\n|\s{2,}', dt_val) 
            for p in punches:
                try:
                    extracted_records.append({'Employee_ID': row['Employee_ID'], 'Full_Name': row['Full_Name'], 'DateTime': pd.to_datetime(p.strip())})
                except:
                    pass
    
    df_clean = pd.DataFrame(extracted_records)
    if df_clean.empty:
        raise ValueError(f"No valid attendance punches found in file for {campus_name}")
        
    df_clean['Date'] = df_clean['DateTime'].dt.date
    
    # Calendar setup for this campus
    min_date, max_date = df_clean['Date'].min(), df_clean['Date'].max()
    all_dates = pd.date_range(start=min_date, end=max_date).date
    sundays = [d for d in all_dates if d.weekday() == 6]
    possible_working_days = [d for d in all_dates if d.weekday() != 6]
    
    # Reconcile off days with global rules
    selected_holidays_dates = list(global_config['selected_holidays_dates'])
    valid_working_dates = [d for d in possible_working_days if d not in selected_holidays_dates]
    
    approved_leaves_dict = global_config['approved_leaves_dict']
    date_overrides = global_config['date_overrides']
    no_lc_dates_set = global_config['no_lc_dates_set']
    vip_ids = global_config['vip_ids']

    def get_shift_rules(current_date):
        if current_date in date_overrides:
            return date_overrides[current_date]
        is_friday = (current_date.weekday() == 4)
        if shift2_date is not None and current_date >= shift2_date:
            if is_friday: return shift2_friday_in, shift2_friday_out
            return shift2_in, shift2_out
        else:
            if is_friday: return friday_shift_in, friday_shift_out
            return shift_in, shift_out

    daily_in = df_clean.groupby(['Employee_ID', 'Full_Name', 'Date'])['DateTime'].min().reset_index()
    daily_out = df_clean.groupby(['Employee_ID', 'Full_Name', 'Date'])['DateTime'].max().reset_index()
    daily_counts = df_clean.groupby(['Employee_ID', 'Full_Name', 'Date'])['DateTime'].count().reset_index()
    
    daily_attendance = pd.merge(daily_in, daily_out, on=['Employee_ID', 'Full_Name', 'Date'], suffixes=('_In', '_Out'))
    daily_attendance = pd.merge(daily_attendance, daily_counts, on=['Employee_ID', 'Full_Name', 'Date'])
    daily_attendance.rename(columns={'DateTime': 'Punch_Count'}, inplace=True)

    # Ghost Protocol (VIP Protection)
    if vip_ids:
        for idx, row in daily_attendance.iterrows():
            if str(row['Employee_ID']).strip() in vip_ids:
                c_date = row['Date']
                active_in, active_out = get_shift_rules(c_date)
                
                # Protect In-Time
                active_in_dt = dt.datetime.combine(c_date, active_in)
                grace_dt = active_in_dt + dt.timedelta(minutes=grace_mins)
                if row['DateTime_In'] > pd.Timestamp(grace_dt):
                    s_min = active_in.hour * 60 + active_in.minute
                    e_min = grace_dt.hour * 60 + grace_dt.minute
                    if e_min >= s_min:
                        r_min = random.randint(s_min, e_min)
                        daily_attendance.at[idx, 'DateTime_In'] = pd.Timestamp(dt.datetime.combine(c_date, dt.time(r_min // 60, r_min % 60)))
                        
                # Protect Out-Time
                missing_out = (row['Punch_Count'] == 1 or row['DateTime_In'] == row['DateTime_Out'])
                if not missing_out:
                    active_out_dt = pd.Timestamp(dt.datetime.combine(c_date, active_out))
                    out_delta_mins = (active_out_dt - row['DateTime_Out']).total_seconds() / 60.0
                    if out_delta_mins >= 15:
                        s_out_min = active_out.hour * 60 + active_out.minute
                        e_out_min = s_out_min + 10  
                        r_out_min = random.randint(s_out_min, e_out_min)
                        new_out = pd.Timestamp(dt.datetime.combine(c_date, dt.time(r_out_min // 60, r_out_min % 60)))
                        if new_out > daily_attendance.at[idx, 'DateTime_In']:
                            daily_attendance.at[idx, 'DateTime_Out'] = new_out

    # Pre-evaluation rules
    def evaluate_day(row):
        if row['Date'] in selected_holidays_dates or row['Date'] in sundays:
            return pd.Series(["On Time"], index=['Daily_Status'])
        
        active_in, active_out = get_shift_rules(row['Date'])
        t_in = row['DateTime_In']
        t_out = row['DateTime_Out']
        p_count = row['Punch_Count']
        
        active_in_dt = pd.Timestamp(dt.datetime.combine(row['Date'], active_in))
        active_out_dt = pd.Timestamp(dt.datetime.combine(row['Date'], active_out))
        
        in_delta_mins = (t_in - active_in_dt).total_seconds() / 60.0
        missing_out = (p_count == 1 or t_in == t_out)
        out_delta_mins = 0 if missing_out else (active_out_dt - t_out).total_seconds() / 60.0
        
        statuses = []
        if in_delta_mins >= 90: statuses.append("Short Leave")
        elif in_delta_mins >= 30: statuses.append("Double LC")
        elif in_delta_mins > grace_mins: statuses.append("LC")
            
        if missing_out:
            statuses.append("Missing Out Time")
        else:
            if out_delta_mins >= 90: statuses.append("Short Leave")
            elif out_delta_mins >= 30: statuses.append("Double LC")
            elif out_delta_mins >= 15: statuses.append("LC")
                
        if row['Date'] in no_lc_dates_set:
            statuses = [s for s in statuses if s != "LC"]
        
        final_status = " + ".join(statuses) if statuses else "On Time"
        return pd.Series([final_status], index=['Daily_Status'])

    eval_res = daily_attendance.apply(evaluate_day, axis=1)
    daily_attendance = pd.concat([daily_attendance, eval_res], axis=1)

    employees = df_clean[['Employee_ID', 'Full_Name']].drop_duplicates()
    employees = employees.sort_values(by='Employee_ID', key=lambda col: pd.to_numeric(col, errors='coerce'))

    summary_results = []
    sorted_records = []

    for _, emp in employees.iterrows():
        eid = emp['Employee_ID']
        emp_data = daily_attendance[daily_attendance['Employee_ID'] == eid]
        
        raw_status_map = {}
        for d in all_dates:
            match = emp_data[emp_data['Date'] == d]
            if not match.empty:
                raw_status_map[d] = match.iloc[0]['Daily_Status']
            else:
                if eid in approved_leaves_dict and d in approved_leaves_dict[eid]: 
                    raw_status_map[d] = "Approved Leave"
                elif d in sundays: 
                    raw_status_map[d] = "Sunday"
                elif d in selected_holidays_dates: 
                    raw_status_map[d] = "Holiday"
                else: 
                    raw_status_map[d] = "Absent"

        # Sandwich Logic
        final_status_map = raw_status_map.copy()
        sandwichable_days = ["Sunday", "Holiday"]
        
        for d in all_dates:
            if raw_status_map[d] in sandwichable_days:
                prev_days = [pd for pd in all_dates if pd < d and raw_status_map[pd] not in sandwichable_days]
                next_days = [nd for nd in all_dates if nd > d and raw_status_map[nd] not in sandwichable_days]
                
                prev_work = prev_days[-1] if prev_days else None
                next_work = next_days[0] if next_days else None
                
                prev_absent = (prev_work and raw_status_map[prev_work] == "Absent")
                next_absent = (next_work and raw_status_map[next_work] == "Absent")
                
                if prev_absent or next_absent:
                    final_status_map[d] = "Sandwich Deducted"

        days_present_valid, sl_count, lc_count, double_lc_count, missing_out_count, full_leaves = 0, 0, 0, 0, 0, 0
        
        for d in all_dates:
            stat = final_status_map[d]
            t_in_str, t_out_str = "—", "—"
            
            match = emp_data[emp_data['Date'] == d]
            if not match.empty:
                t_in_str = match.iloc[0]['DateTime_In'].strftime('%I:%M %p')
                t_out_str = "Missing" if "Missing Out Time" in stat else match.iloc[0]['DateTime_Out'].strftime('%I:%M %p')
                if d in valid_working_dates:
                    days_present_valid += 1
                    parts = stat.split(" + ")
                    for p in parts:
                        if p == "LC": lc_count += 1
                        if p == "Double LC": double_lc_count += 1
                        if p == "Short Leave": sl_count += 1
                        if p == "Missing Out Time": missing_out_count += 1

            if stat in ["Absent", "Sandwich Deducted"]:
                full_leaves += 1

            sorted_records.append({
                'ID': eid,
                'Name': emp['Full_Name'],
                'Date': d.strftime('%d-%b-%Y'),
                'Time In': t_in_str,
                'Time Out': t_out_str,
                'Status': stat
            })

        summary_results.append({
            'ID': eid, 'Name': emp['Full_Name'],
            'Work Days': days_present_valid,
            'LCs': lc_count, 'Double LCs': double_lc_count, 'Short Leaves': sl_count,
            'Missing Out Time': missing_out_count,
            'Full Leaves': max(0, full_leaves)
        })

    summary_df = pd.DataFrame(summary_results)
    sorted_report_df = pd.DataFrame(sorted_records)
    summary_df = summary_df.sort_values(by='ID', key=lambda col: pd.to_numeric(col, errors='coerce'))

    return {
        'campus_name': campus_name,
        'summary_df': summary_df,
        'sorted_report_df': sorted_report_df,
        'valid_working_dates': valid_working_dates,
        'min_date': min_date,
        'max_date': max_date
    }


# ==============================================================================
# MAIN PAGE INTERFACE: MULTI-FILE UPLOAD & BATCH PROCESSING
# ==============================================================================
st.markdown("### Step 1: Upload Biometric Excel Files (Single or Multi-Campus)")
uploaded_files = st.file_uploader(
    "📁 Upload Biometric Files (.xlsx) — Select up to 24+ campus files at once", 
    type=["xlsx"], 
    accept_multiple_files=True
)

if uploaded_files:
    st.info(f"📁 **{len(uploaded_files)} file(s) selected.** Verify detected Campus Names below before running:")
    
    # Campus Name Review Table
    col_heads = st.columns([3, 3])
    with col_heads[0]: st.markdown("**File Name**")
    with col_heads[1]: st.markdown("**Assigned Campus Name**")
    
    campus_name_mapping = {}
    for idx, f in enumerate(uploaded_files):
        row_cols = st.columns([3, 3])
        with row_cols[0]:
            st.caption(f"📄 `{f.name}`")
        with row_cols[1]:
            default_c_name = extract_campus_name(f.name)
            user_c_name = st.text_input(
                f"Campus for file {idx+1}", 
                value=default_c_name, 
                key=f"c_name_{idx}", 
                label_visibility="collapsed"
            )
            campus_name_mapping[f.name] = user_c_name.strip() if user_c_name.strip() else default_c_name

    st.markdown("---")
    st.markdown("### Step 2: Global Calendar & Off-Days Configuration")
    
    # Global calendar selection based on current month range
    all_sats_off = st.checkbox("All Saturdays Off Across Campuses", value=False, help="Check this if Saturdays are general off days.")
    
    col_cal1, col_cal2 = st.columns(2)
    with col_cal1:
        manual_off_saturdays = st.date_input(
            "Select Specific Off Saturday(s):",
            value=[],
            key="batch_off_saturdays"
        )
    with col_cal2:
        manual_holidays = st.date_input(
            "Select General Holidays / Extra Off Dates:",
            value=[],
            key="batch_general_holidays"
        )
        
    # Standardize off-days
    selected_holidays_dates = []
    
    # If a single date or list of dates was returned
    for m_item in [manual_off_saturdays, manual_holidays]:
        if isinstance(m_item, (list, tuple)):
            for d in m_item:
                if isinstance(d, dt.datetime): selected_holidays_dates.append(d.date())
                elif isinstance(d, dt.date): selected_holidays_dates.append(d)
        elif isinstance(m_item, dt.datetime):
            selected_holidays_dates.append(m_item.date())
        elif isinstance(m_item, dt.date):
            selected_holidays_dates.append(m_item)

    st.markdown("---")
    
    # Run Batch Button
    run_col1, run_col2 = st.columns([1, 3])
    with run_col1:
        start_processing = st.button("🚀 Process All Campuses", type="primary", use_container_width=True)
        
    if start_processing:
        st.session_state.batch_results = {}
        progress_bar = st.progress(0.0)
        status_box = st.status(f"Processing 0 of {len(uploaded_files)} campuses...", expanded=True)
        
        # Prepare approved leaves dict
        approved_leaves_dict = {}
        for _, row in edited_leaves_df.iterrows():
            d_val = row.get('Date')
            e_val = row.get('Employee_ID')
            if pd.notna(d_val) and pd.notna(e_val):
                e_str = str(e_val).strip()
                if e_str.endswith('.0'): e_str = e_str[:-2]
                if e_str:
                    if e_str not in approved_leaves_dict: approved_leaves_dict[e_str] = []
                    try:
                        if isinstance(d_val, dt.date): c_d = d_val
                        elif isinstance(d_val, dt.datetime): c_d = d_val.date()
                        else: c_d = pd.to_datetime(d_val).date()
                        approved_leaves_dict[e_str].append(c_d)
                    except: pass

        date_overrides = {item['date']: (item['in'], item['out']) for item in st.session_state.custom_date_timings}
        no_lc_dates_set = set(st.session_state.no_lc_dates)
        
        errors = []
        for i, file_obj in enumerate(uploaded_files):
            c_name = campus_name_mapping.get(file_obj.name, extract_campus_name(file_obj.name))
            status_box.write(f"⏳ ({i+1}/{len(uploaded_files)}) Analyzing attendance for: **{c_name}**...")
            
            # Read file in memory
            file_bytes = file_obj.getvalue()
            
            # Determine Saturdays if All Saturdays Off is selected
            campus_holidays = list(selected_holidays_dates)
                
            c_vips = get_vip_ids_for_campus(c_name)
            
            global_config = {
                'selected_holidays_dates': campus_holidays,
                'approved_leaves_dict': approved_leaves_dict,
                'date_overrides': date_overrides,
                'no_lc_dates_set': no_lc_dates_set,
                'vip_ids': c_vips
            }
            
            try:
                res = process_single_file(file_bytes, c_name, global_config)
                # If all_sats_off was ticked, ensure saturdays in campus range were excluded
                if all_sats_off:
                    sat_range = [d for d in pd.date_range(res['min_date'], res['max_date']).date if d.weekday() == 5]
                    if not set(sat_range).issubset(set(campus_holidays)):
                        global_config['selected_holidays_dates'] = list(set(campus_holidays + sat_range))
                        res = process_single_file(file_bytes, c_name, global_config)
                        
                st.session_state.batch_results[c_name] = res
                status_box.write(f"✅ **{c_name}**: Successfully processed ({len(res['summary_df'])} staff).")
            except Exception as e:
                errors.append(f"{c_name} ({file_obj.name}): {str(e)}")
                status_box.write(f"❌ **{c_name}**: Error - {str(e)}")
                
            progress_bar.progress((i + 1) / len(uploaded_files))
            
        if errors:
            status_box.update(label=f"Completed with {len(errors)} error(s)", state="error", expanded=True)
        else:
            status_box.update(label=f"🎉 All {len(uploaded_files)} Campuses Processed Successfully!", state="complete", expanded=False)


# ==============================================================================
# STEP 3: RESULTS VIEWER & BULK DOWNLOADS
# ==============================================================================
if st.session_state.batch_results:
    st.markdown("---")
    st.markdown("### Step 3: View Reports & Download")
    
    total_processed = len(st.session_state.batch_results)
    st.success(f"📊 **{total_processed} Campus Report(s) Ready.** You can inspect any campus directly or download individual/all reports.")
    
    # --- BULK ZIP DOWNLOAD ---
    def create_batch_zip(batch_dict):
        zip_buffer = BytesIO()
        with zipfile.ZipFile(zip_buffer, "w", zipfile.ZIP_DEFLATED) as zip_file:
            for c_name, data in batch_dict.items():
                safe_name = re.sub(r'[^\w\s-]', '', c_name).strip().replace(' ', '_')
                # 1. Summary Excel (Sample Format)
                excel_bytes = generate_summary_excel(data['summary_df'])
                zip_file.writestr(f"{safe_name}/{safe_name}_Summary.xlsx", excel_bytes)
                # 2. Master Excel
                sorted_excel_bytes = generate_excel(data['sorted_report_df'])
                zip_file.writestr(f"{safe_name}/{safe_name}_Master_Report.xlsx", sorted_excel_bytes)
                # 3. Summary PDF (Exact A4 Landscape Format)
                try:
                    month_label = f"{data['min_date'].strftime('%B %Y')}" if 'min_date' in data else ""
                    summary_pdf = generate_summary_pdf(data['summary_df'], c_name, month_label)
                    zip_file.writestr(f"{safe_name}/{safe_name}_Summary.pdf", summary_pdf)
                except:
                    pass
                # 4. Master PDF
                try:
                    master_pdf = generate_pdf(data['sorted_report_df'], "Sorted Master Report", c_name)
                    zip_file.writestr(f"{safe_name}/{safe_name}_Master_Report.pdf", master_pdf)
                except:
                    pass
        return zip_buffer.getvalue()

    col_bulk1, col_bulk2 = st.columns([1, 2])
    with col_bulk1:
        zip_bytes = create_batch_zip(st.session_state.batch_results)
        st.download_button(
            label="📦 Download ALL Campuses Reports (.ZIP)",
            data=zip_bytes,
            file_name="All_Campuses_Attendance_Reports.zip",
            mime="application/zip",
            use_container_width=True
        )
    with col_bulk2:
        st.caption("✨ Contains **Monthly Summaries & Master Reports in both Excel (.xlsx) and PDF** organized in campus folders.")

    st.markdown("---")
    
    # --- CAMPUS SELECTOR FOR DIRECT INSPECTION ---
    st.subheader("👁️ Inspect Campus Report Directly")
    campus_list = list(st.session_state.batch_results.keys())
    selected_campus = st.selectbox("Select Campus to View:", campus_list)
    
    if selected_campus and selected_campus in st.session_state.batch_results:
        camp_data = st.session_state.batch_results[selected_campus]
        summary_df = camp_data['summary_df']
        sorted_report_df = camp_data['sorted_report_df']
        valid_dates = camp_data['valid_working_dates']
        
        tab1, tab2 = st.tabs(["📋 Monthly Payroll Summary", "📑 Sorted Master Report"])
        
        with tab1:
            st.subheader(f"Payroll Summary: {selected_campus} | Expected Valid Days: {len(valid_dates) - extra_off_days_count}")
            styled_summary = apply_style(
                summary_df.style, 
                style_attendance, 
                ['LCs', 'Double LCs', 'Short Leaves', 'Missing Out Time', 'Full Leaves']
            )
            st.dataframe(styled_summary, use_container_width=True)
            
            c1, c2 = st.columns(2)
            with c1:
                excel_data = generate_summary_excel(summary_df)
                st.download_button(
                    f"📥 Download Summary (Excel) - {selected_campus}", 
                    data=excel_data, 
                    file_name=f"{selected_campus}_summary.xlsx", 
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    key=f"dl_sum_xl_{selected_campus}"
                )
            with c2:
                m_label = f"{camp_data['min_date'].strftime('%B %Y')}" if 'min_date' in camp_data else ""
                pdf_data = generate_summary_pdf(summary_df, selected_campus, m_label)
                st.download_button(
                    f"📄 Download Summary (PDF) - {selected_campus}", 
                    data=pdf_data, 
                    file_name=f"{selected_campus}_summary.pdf", 
                    mime="application/pdf",
                    key=f"dl_sum_pdf_{selected_campus}"
                )

        with tab2:
            st.subheader(f"Master Detailed Report: {selected_campus}")
            st.dataframe(sorted_report_df.style.map(lambda v: style_attendance(v), subset=['Status']), use_container_width=True)
            
            c3, c4 = st.columns(2)
            with c3:
                excel_sorted = generate_excel(sorted_report_df)
                st.download_button(
                    f"📥 Download Master Report (Excel) - {selected_campus}", 
                    data=excel_sorted, 
                    file_name=f"{selected_campus}_master.xlsx", 
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    key=f"dl_mst_xl_{selected_campus}"
                )
            with c4:
                pdf_sorted = generate_pdf(sorted_report_df, "Sorted Master Report", selected_campus)
                st.download_button(
                    f"📄 Download Master Report (PDF) - {selected_campus}", 
                    data=pdf_sorted, 
                    file_name=f"{selected_campus}_master.pdf", 
                    mime="application/pdf",
                    key=f"dl_mst_pdf_{selected_campus}"
                )
