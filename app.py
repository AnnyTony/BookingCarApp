import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import numpy as np

# --- 1. CẤU HÌNH GIAO DIỆN CHUẨN DASHBOARD ---
st.set_page_config(page_title="Executive Fleet Dashboard", page_icon="📊", layout="wide")

# CSS để giống Power BI (Nền xám nhạt, Card trắng nổi, Font chuẩn)
st.markdown("""
<style>
    /* Tổng thể nền */
    .stApp {background-color: #f0f2f5;}
    
    /* Sidebar */
    [data-testid="stSidebar"] {background-color: #ffffff; border-right: 1px solid #e0e0e0;}
    
    /* Metric Cards */
    div[data-testid="stMetricValue"] {font-size: 28px; color: #0078d4; font-weight: 700;}
    div[data-testid="stMetricLabel"] {font-size: 14px; color: #605e5c;}
    
    /* Header */
    .dashboard-title {font-size: 32px; font-weight: bold; color: #201f1e; margin-bottom: 5px;}
    .dashboard-subtitle {font-size: 16px; color: #8a8886; margin-bottom: 20px;}
</style>
""", unsafe_allow_html=True)

# --- 2. HÀM XỬ LÝ DỮ LIỆU (MERGE 3 TAB) ---
@st.cache_data
def load_data_powerbi(file):
    try:
        xl = pd.ExcelFile(file, engine='openpyxl')
        
        # Tìm tên các Sheet
        sheet_booking = next((s for s in xl.sheet_names if 'booking' in s.lower()), None)
        sheet_cbnv = next((s for s in xl.sheet_names if 'cbnv' in s.lower() or 'staff' in s.lower()), None)
        
        if not sheet_booking:
            return "❌ Lỗi: Không tìm thấy Sheet 'Booking car'."

        # A. LOAD BOOKING
        df_bk = xl.parse(sheet_booking)
        df_bk.columns = df_bk.columns.str.strip()
        
        # Xử lý ngày giờ
        df_bk['Start_Datetime'] = pd.to_datetime(df_bk['Ngày khởi hành'].astype(str) + ' ' + df_bk['Giờ khởi hành'].astype(str), errors='coerce')
        df_bk['End_Datetime'] = pd.to_datetime(df_bk['Ngày khởi hành'].astype(str) + ' ' + df_bk['Giờ kết thúc'].astype(str), errors='coerce')
        
        mask_overnight = df_bk['End_Datetime'] < df_bk['Start_Datetime']
        df_bk.loc[mask_overnight, 'End_Datetime'] += pd.Timedelta(days=1)
        
        df_bk['Duration_Hours'] = (df_bk['End_Datetime'] - df_bk['Start_Datetime']).dt.total_seconds() / 3600
        df_bk['Month'] = df_bk['Start_Datetime'].dt.month
        df_bk['Year'] = df_bk['Start_Datetime'].dt.year
        df_bk['Weekday'] = df_bk['Start_Datetime'].dt.day_name()
        
        # Logic 1: Nửa ngày / Cả ngày
        df_bk['Session_Type'] = df_bk['Duration_Hours'].apply(lambda x: 'Nửa ngày (≤4h)' if x <= 4 else 'Cả ngày (>4h)')
        
        # Logic 2: Tỉnh / Nội thành (Dựa trên keyword Lộ trình)
        def classify_scope(route):
            if pd.isna(route): return "Không xác định"
            route = str(route).lower()
            keywords = ['tỉnh', 'tp.', 'bình dương', 'đồng nai', 'vũng tàu', 'long an', 'hà nội', 'bắc ninh', 'hải phòng']
            # Nếu lộ trình chứa từ khóa tỉnh -> Đi Tỉnh
            if any(k in route for k in keywords): return "Đi Tỉnh"
            return "Nội thành"
        
        if 'Lộ trình' in df_bk.columns:
            df_bk['Scope'] = df_bk['Lộ trình'].apply(classify_scope)
        else:
            df_bk['Scope'] = "Nội thành" # Mặc định

        # B. LOAD CBNV & MERGE (VLOOKUP)
        if sheet_cbnv:
            df_staff = xl.parse(sheet_cbnv)
            df_staff.columns = df_staff.columns.str.strip()
            
            # Mapping tên cột cho chuẩn
            col_map = {}
            for c in df_staff.columns:
                c_low = c.lower()
                if 'name' in c_low: col_map[c] = 'Full Name'
                if 'công ty' in c_low or 'company' in c_low: col_map[c] = 'Company_Lookup'
                if 'bu' in c_low or 'bộ phận' in c_low: col_map[c] = 'Dept_Lookup'
                if 'location' in c_low or 'site' in c_low: col_map[c] = 'Location_Lookup'
            
            df_staff = df_staff.rename(columns=col_map)
            
            # Merge (Left Join)
            df_final = pd.merge(df_bk, df_staff, left_on='Người sử dụng xe', right_on='Full Name', how='left')
            
            # Fillna cho các trường hợp không tìm thấy nhân viên
            df_final['Company'] = df_final['Company_Lookup'].fillna('Khác / Ngoài DS')
            df_final['Department'] = df_final['Dept_Lookup'].fillna('Khác')
            
            # Logic 3: Phân vùng Bắc/Nam từ Location
            def get_region(loc):
                if pd.isna(loc): return 'Unknown'
                loc = str(loc).upper()
                if 'HN' in loc or 'BẮC' in loc or 'HANOI' in loc: return 'Miền Bắc'
                if 'HCM' in loc or 'NAM' in loc: return 'Miền Nam'
                return 'Khác'
            
            df_final['Region'] = df_final['Location_Lookup'].apply(get_region)
        else:
            # Fallback nếu không có sheet CBNV
            df_final = df_bk
            df_final['Company'] = "Unknown"
            df_final['Department'] = "Unknown"
            df_final['Region'] = "Miền Nam" # Mặc định

        return df_final

    except Exception as e:
        return f"Lỗi xử lý file: {str(e)}"

# --- 3. GIAO DIỆN CHÍNH ---
st.markdown("<div class='dashboard-title'>📊 Fleet Analytics Dashboard</div>", unsafe_allow_html=True)
st.markdown("<div class='dashboard-subtitle'>Hệ thống báo cáo quản trị đội xe tập trung</div>", unsafe_allow_html=True)

# UPLOAD
uploaded_file = st.sidebar.file_uploader("📂 Tải file Excel báo cáo", type=['xlsx'])

if uploaded_file:
    df = load_data_powerbi(uploaded_file)
    if isinstance(df, str):
        st.error(df)
        st.stop()

    # --- 4. CASCADING FILTERS (BỘ LỌC THÔNG MINH KIỂU POWER BI) ---
    st.sidebar.header("🎛️ Bộ Lọc (Slicers)")

    # 1. Lọc Năm & Tháng (Cao nhất)
    years = sorted(df['Year'].dropna().unique())
    selected_year = st.sidebar.selectbox("📅 Chọn Năm", years, index=len(years)-1)
    
    df_y = df[df['Year'] == selected_year]
    
    # 2. Lọc Vùng Miền (Ảnh hưởng bởi Năm)
    regions = ['Tất cả'] + sorted(list(df_y['Region'].unique()))
    selected_region = st.sidebar.selectbox("🌍 Chọn Vùng Miền", regions)
    
    if selected_region != 'Tất cả':
        df_r = df_y[df_y['Region'] == selected_region]
    else:
        df_r = df_y
        
    # 3. Lọc Công Ty (Ảnh hưởng bởi Vùng)
    companies = ['Tất cả'] + sorted(list(df_r['Company'].unique()))
    selected_company = st.sidebar.selectbox("🏢 Chọn Công Ty", companies)
    
    if selected_company != 'Tất cả':
        df_c = df_r[df_r['Company'] == selected_company]
    else:
        df_c = df_r

    # Dữ liệu cuối cùng để vẽ (df_final)
    df_final = df_c

    # --- 5. TÍNH TOÁN KPI (OCCUPANCY CHUẨN) ---
    # Logic xe: Nam 16, Bắc 5. Tổng 21.
    if selected_region == 'Miền Nam': total_cars = 16
    elif selected_region == 'Miền Bắc': total_cars = 5
    else: total_cars = 21 
    
    # Số ngày lọc được
    if not df_final.empty:
        num_days = (df_final['Start_Datetime'].max() - df_final['Start_Datetime'].min()).days + 1
        num_days = max(1, num_days)
    else:
        num_days = 1
        
    total_trips = len(df_final)
    total_hours = df_final['Duration_Hours'].sum()
    capacity = total_cars * num_days * 9 # 9 tiếng/ngày
    occupancy = (total_hours / capacity * 100) if capacity > 0 else 0
    
    # Đếm trạng thái
    if 'Tình trạng đơn yêu cầu' in df_final.columns:
        cancel_count = df_final[df_final['Tình trạng đơn yêu cầu'].str.contains('CANCEL|REJECT', case=False, na=False)].shape[0]
        completed_count = df_final[df_final['Tình trạng đơn yêu cầu'].str.contains('CLOSED|APPROVED', case=False, na=False)].shape[0]
    else:
        cancel_count = 0
        completed_count = 0

    # --- 6. HIỂN THỊ KPI CARDS ---
    c1, c2, c3, c4 = st.columns(4)
    c1.metric("Tổng Số Chuyến", f"{total_trips}", f"{completed_count} hoàn thành")
    c2.metric("Tỷ Lệ Lấp Đầy (Occupancy)", f"{occupancy:.1f}%", f"Capacity: {total_cars} xe")
    c3.metric("Số Giờ Vận Hành", f"{total_hours:,.0f}h")
    c4.metric("Chuyến Hủy/Từ Chối", f"{cancel_count}", delta_color="inverse")
    
    st.markdown("---")

    # --- 7. BIỂU ĐỒ DASHBOARD (POWER BI STYLE) ---
    
    # HÀNG 1: PHÂN BỐ CÔNG TY (Sunburst & Bar)
    col_row1_1, col_row1_2 = st.columns([1, 1])
    
    with col_row1_1:
        st.subheader("🏢 Cơ Cấu Chuyến Đi Theo Công Ty & Bộ Phận")
        # Sunburst Chart: Biểu đồ tròn phân cấp (Công ty -> Bộ phận)
        # Đây là biểu đồ xịn nhất để thể hiện Drill-down
        df_sunburst = df_final.groupby(['Company', 'Department']).size().reset_index(name='Count')
        fig_sun = px.sunburst(df_sunburst, path=['Company', 'Department'], values='Count',
                              color='Count', color_continuous_scale='Blues')
        st.plotly_chart(fig_sun, use_container_width=True)
        
    with col_row1_2:
        st.subheader("📊 Tỷ Trọng Trạng Thái Theo Công Ty")
        # Stacked Bar Chart: Trạng thái (Approved/Cancel) theo Công ty
        if 'Tình trạng đơn yêu cầu' in df_final.columns:
            df_status = df_final.groupby(['Company', 'Tình trạng đơn yêu cầu']).size().reset_index(name='Count')
            fig_bar = px.bar(df_status, x='Company', y='Count', color='Tình trạng đơn yêu cầu',
                             title="Trạng thái chuyến đi từng Công ty",
                             color_discrete_map={'CLOSED': '#00CC96', 'APPROVED': '#636EFA', 'CANCELLED': '#EF553B', 'REJECTED': '#AB63FA'})
            st.plotly_chart(fig_bar, use_container_width=True)

    st.markdown("---")

    # HÀNG 2: PHẠM VI & LOẠI CHUYẾN
    col_row2_1, col_row2_2, col_row2_3 = st.columns(3)
    
    with col_row2_1:
        st.subheader("🗺️ Tỉnh vs Nội Thành")
        scope_counts = df_final['Scope'].value_counts().reset_index()
        scope_counts.columns = ['Phạm vi', 'Số chuyến']
        fig_pie1 = px.pie(scope_counts, values='Số chuyến', names='Phạm vi', hole=0.6, color_discrete_sequence=px.colors.qualitative.Prism)
        st.plotly_chart(fig_pie1, use_container_width=True)
        
    with col_row2_2:
        st.subheader("⏱️ Nửa Ngày vs Cả Ngày")
        sess_counts = df_final['Session_Type'].value_counts().reset_index()
        sess_counts.columns = ['Loại', 'Số chuyến']
        fig_pie2 = px.pie(sess_counts, values='Số chuyến', names='Loại', hole=0.6, color_discrete_sequence=px.colors.qualitative.Pastel)
        st.plotly_chart(fig_pie2, use_container_width=True)
        
    with col_row2_3:
        st.subheader("🚗 Top 5 Xe Hoạt Động Cao Nhất")
        if 'Biển số xe' in df_final.columns:
            car_top = df_final['Biển số xe'].value_counts().head(5).reset_index()
            car_top.columns = ['Xe', 'Số chuyến']
            fig_car = px.bar(car_top, x='Số chuyến', y='Xe', orientation='h', text_auto=True)
            st.plotly_chart(fig_car, use_container_width=True)

    # HÀNG 3: XU HƯỚNG THỜI GIAN
    st.subheader("📈 Xu Hướng Occupancy Rate Theo Tháng")
    monthly_stats = df_final.groupby('Month').agg(
        Total_Hours=('Duration_Hours', 'sum'),
    ).reset_index()
    
    # Tính Capacity cố định theo tháng (26 ngày làm việc)
    monthly_cap = total_cars * 26 * 9
    monthly_stats['Occupancy'] = (monthly_stats['Total_Hours'] / monthly_cap * 100)
    
    fig_line = go.Figure()
    fig_line.add_trace(go.Bar(x=monthly_stats['Month'], y=monthly_stats['Total_Hours'], name='Giờ chạy thực tế', opacity=0.4))
    fig_line.add_trace(go.Scatter(x=monthly_stats['Month'], y=monthly_stats['Occupancy'], name='Tỷ lệ lấp đầy (%)', yaxis='y2', mode='lines+markers', line=dict(color='firebrick', width=3)))
    
    fig_line.update_layout(
        xaxis=dict(title='Tháng'),
        yaxis=dict(title='Giờ chạy'),
        yaxis2=dict(title='Tỷ lệ %', overlaying='y', side='right', range=[0, 100]),
        legend=dict(x=0, y=1.1, orientation='h')
    )
    st.plotly_chart(fig_line, use_container_width=True)

else:
    st.info("👋 Chào mừng! Hãy tải file Excel (có tab Booking & CBNV) để xem Dashboard.")