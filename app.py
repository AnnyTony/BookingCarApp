import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go

# --- 1. CẤU HÌNH TRANG ---
st.set_page_config(page_title="Power BI Style Dashboard", page_icon="📊", layout="wide")

# CSS: Giao diện sạch, giống Dashboard doanh nghiệp
st.markdown("""
<style>
    .main-header {font-size: 26px; font-weight: bold; color: #2c3e50;}
    div[data-testid="stMetricValue"] {font-size: 22px; color: #2980b9;}
    [data-testid="stSidebar"] {background-color: #f1f3f6;}
    /* Chỉnh màu cho các Tab */
    .stTabs [data-baseweb="tab-list"] {gap: 10px;}
    .stTabs [data-baseweb="tab"] {height: 50px; white-space: pre-wrap; background-color: white; border-radius: 4px; box-shadow: 0px 1px 3px rgba(0,0,0,0.1);}
    .stTabs [aria-selected="true"] {background-color: #e3f2fd; color: #1976d2;}
</style>
""", unsafe_allow_html=True)

st.markdown("<div class='main-header'>📊 Fleet Management Intelligence (Power BI Style)</div>", unsafe_allow_html=True)
st.markdown("---")

# --- 2. LOAD DATA (GIỮ NGUYÊN LOGIC ĐA SHEET CŨ) ---
@st.cache_data
def load_data_pro(file):
    try:
        xl = pd.ExcelFile(file, engine='openpyxl')
        
        # Tìm sheet
        sheet_booking = next((s for s in xl.sheet_names if 'booking' in s.lower()), None)
        sheet_cbnv = next((s for s in xl.sheet_names if 'cbnv' in s.lower() or 'staff' in s.lower()), None)
        
        if not sheet_booking: return "Thiếu sheet Booking"

        # 1. Booking
        df_bk = xl.parse(sheet_booking)
        df_bk.columns = df_bk.columns.str.strip()
        
        # Xử lý datetime
        df_bk['Start_Datetime'] = pd.to_datetime(df_bk['Ngày khởi hành'].astype(str) + ' ' + df_bk['Giờ khởi hành'].astype(str), errors='coerce')
        df_bk['End_Datetime'] = pd.to_datetime(df_bk['Ngày khởi hành'].astype(str) + ' ' + df_bk['Giờ kết thúc'].astype(str), errors='coerce')
        mask_overnight = df_bk['End_Datetime'] < df_bk['Start_Datetime']
        df_bk.loc[mask_overnight, 'End_Datetime'] += pd.Timedelta(days=1)
        
        df_bk['Duration_Hours'] = (df_bk['End_Datetime'] - df_bk['Start_Datetime']).dt.total_seconds() / 3600
        df_bk['Month_Year'] = df_bk['Start_Datetime'].dt.to_period('M').astype(str)
        df_bk['Year'] = df_bk['Start_Datetime'].dt.year
        df_bk['Loại Chuyến'] = df_bk['Duration_Hours'].apply(lambda x: 'Nửa ngày' if x <= 4 else 'Cả ngày')

        # Logic Đi Tỉnh / Nội Thành
        def check_scope(route):
            s = str(route).lower()
            return "Đi Tỉnh" if any(x in s for x in ['tỉnh', 'tp.', 'bình dương', 'đồng nai', 'vũng tàu']) else "Nội thành"
        df_bk['Phạm Vi'] = df_bk['Lộ trình'].apply(check_scope) if 'Lộ trình' in df_bk.columns else "Unknown"

        # 2. Merge CBNV
        if sheet_cbnv:
            df_staff = xl.parse(sheet_cbnv)
            df_staff.columns = df_staff.columns.str.strip()
            
            # Map cột
            col_map = {}
            for c in df_staff.columns:
                if 'full name' in c.lower() or 'họ tên' in c.lower(): col_map[c] = 'Full Name'
                if 'công ty' in c.lower(): col_map[c] = 'Công ty_L'
                if 'bu' in c.lower() or 'bộ phận' in c.lower(): col_map[c] = 'BoPhan_L'
                if 'location' in c.lower(): col_map[c] = 'Location_L'
            
            df_staff = df_staff.rename(columns=col_map)
            
            # Merge
            df_final = pd.merge(df_bk, df_staff[['Full Name', 'Công ty_L', 'BoPhan_L', 'Location_L']], 
                                left_on='Người sử dụng xe', right_on='Full Name', how='left')
            
            # Fillna
            df_final['Công ty'] = df_final['Công ty_L'].fillna('Chưa xác định')
            df_final['Bộ phận'] = df_final['BoPhan_L'].fillna('Chưa xác định')
            
            # Logic Bắc/Nam
            def get_region(loc):
                loc = str(loc).upper()
                if 'HCM' in loc or 'NAM' in loc: return 'Miền Nam'
                if 'HN' in loc or 'BẮC' in loc: return 'Miền Bắc'
                return 'Khác'
            df_final['Vùng Miền'] = df_final['Location_L'].apply(get_region)
            
        else:
            df_final = df_bk
            df_final['Công ty'] = "No Data"
            df_final['Bộ phận'] = "No Data"
            df_final['Vùng Miền'] = "Khác"
            
        return df_final

    except Exception as e:
        return f"Error: {str(e)}"

# --- 3. UPLOAD ---
uploaded_file = st.file_uploader("📂 Kéo thả file Excel (Booking + CBNV)", type=['xlsx'])
if not uploaded_file:
    st.info("👋 Chờ file dữ liệu...")
    st.stop()

df = load_data_pro(uploaded_file)
if isinstance(df, str):
    st.error(df)
    st.stop()

# --- 4. SIDEBAR "CASCADING" (BỘ LỌC PHÂN CẤP THÔNG MINH) ---
with st.sidebar:
    st.header("🎛️ Bộ lọc Điều khiển")
    
    # 1. Chọn Năm (Gốc)
    years = sorted(df['Year'].dropna().unique())
    selected_years = st.multiselect("Năm:", years, default=years)
    df_lv1 = df[df['Year'].isin(selected_years)]
    
    # 2. Chọn Vùng Miền (Lọc theo Năm)
    regions = ['Tất cả'] + sorted(list(df_lv1['Vùng Miền'].unique()))
    selected_region = st.selectbox("Vùng Miền:", regions)
    
    if selected_region != 'Tất cả':
        df_lv2 = df_lv1[df_lv1['Vùng Miền'] == selected_region]
    else:
        df_lv2 = df_lv1
        
    # 3. Chọn Công Ty (Lọc theo Vùng Miền đã chọn) -> ĐÂY LÀ CHỖ THÔNG MINH
    avail_companies = sorted(df_lv2['Công ty'].astype(str).unique())
    selected_companies = st.multiselect("Công ty:", avail_companies, default=avail_companies)
    
    # 4. Chọn Bộ Phận (Lọc theo Công ty đã chọn)
    if selected_companies:
        df_lv3 = df_lv2[df_lv2['Công ty'].isin(selected_companies)]
    else:
        df_lv3 = df_lv2
        
    avail_depts = sorted(df_lv3['Bộ phận'].astype(str).unique())
    selected_depts = st.multiselect("Phòng ban/Bộ phận:", avail_depts, default=avail_depts)

    # --- ÁP DỤNG FILTER CUỐI CÙNG ---
    if selected_depts:
        df_final_filtered = df_lv3[df_lv3['Bộ phận'].isin(selected_depts)]
    else:
        df_final_filtered = df_lv3
        
    st.success(f"🔍 Dữ liệu: {len(df_final_filtered)} chuyến")

# --- 5. TÍNH KPI OCCUPANCY ---
# Logic xe như cũ
if selected_region == 'Miền Nam': total_cars = 16
elif selected_region == 'Miền Bắc': total_cars = 5
else: total_cars = 21

if 'Start_Datetime' in df_final_filtered.columns and not df_final_filtered.empty:
    days = (df_final_filtered['Start_Datetime'].max() - df_final_filtered['Start_Datetime'].min()).days + 1
    days = max(days, 1)
    cap_hours = total_cars * days * 9
    used_hours = df_final_filtered['Duration_Hours'].sum()
    occupancy = (used_hours / cap_hours * 100) if cap_hours > 0 else 0
else:
    occupancy = 0
    days = 0
    used_hours = 0

# --- 6. DASHBOARD CHÍNH ---

# ROW 1: KPI
c1, c2, c3, c4 = st.columns(4)
c1.metric("Tổng Số Chuyến", len(df_final_filtered))
c2.metric("Tổng Giờ Chạy", f"{used_hours:,.0f}h")
c3.metric("Tỷ lệ Lấp Đầy (Occupancy)", f"{occupancy:.1f}%")
c4.metric("Số Xe Khả Dụng", f"{total_cars} xe")

st.markdown("---")

# ROW 2: BIỂU ĐỒ PHÂN CẤP (SUNBURST) - GIỐNG POWER BI NHẤT
t1, t2 = st.tabs(["🏢 Cấu Trúc & Phân Bổ (Hierarchy)", "📈 Xu Hướng & Hiệu Suất"])

with t1:
    col_sun, col_tree = st.columns([1, 1])
    
    with col_sun:
        st.subheader("Phân bổ: Vùng -> Công Ty -> Bộ Phận")
        # Nhóm dữ liệu để vẽ Sunburst
        df_sun = df_final_filtered.groupby(['Vùng Miền', 'Công ty', 'Bộ phận']).size().reset_index(name='Số chuyến')
        # Xử lý dữ liệu bằng 0 hoặc nhỏ để biểu đồ đẹp hơn
        df_sun = df_sun[df_sun['Số chuyến'] > 0]
        
        fig_sun = px.sunburst(df_sun, path=['Vùng Miền', 'Công ty', 'Bộ phận'], values='Số chuyến',
                              color='Số chuyến', color_continuous_scale='RdBu')
        st.plotly_chart(fig_sun, use_container_width=True)
        st.caption("💡 Mẹo: Click vào vòng tròn để đi sâu (Drill-down) vào từng Công ty/Bộ phận.")

    with col_tree:
        st.subheader("Tỷ lệ Trạng thái chuyến đi")
        if 'Tình trạng đơn yêu cầu' in df_final_filtered.columns:
            status_df = df_final_filtered['Tình trạng đơn yêu cầu'].fillna('Unknown').value_counts().reset_index()
            status_df.columns = ['Status', 'Count']
            color_map = {'CLOSED': 'green', 'APPROVED': 'blue', 'CANCELLED': 'red', 'REJECTED': 'darkred'}
            
            fig_pie = px.pie(status_df, values='Count', names='Status', hole=0.5, 
                             color='Status', color_discrete_map=color_map)
            st.plotly_chart(fig_pie, use_container_width=True)
            
            # Thêm bảng nhỏ bên dưới để xem số reject
            st.dataframe(status_df.set_index('Status').T, use_container_width=True)

    # Biểu đồ cột chồng: Công ty vs Loại chuyến (Nửa ngày/Cả ngày)
    st.subheader("Phân tích Loại chuyến theo Công ty")
    df_type = df_final_filtered.groupby(['Công ty', 'Loại Chuyến']).size().reset_index(name='Count')
    fig_bar_stack = px.bar(df_type, x='Công ty', y='Count', color='Loại Chuyến', 
                           title="Số chuyến Nửa ngày vs Cả ngày theo từng Công ty", barmode='group')
    st.plotly_chart(fig_bar_stack, use_container_width=True)

with t2:
    col_trend, col_map = st.columns([2, 1])
    
    with col_trend:
        st.subheader("Biểu đồ Xu Hướng (Timeline)")
        monthly = df_final_filtered.groupby('Month_Year')['Duration_Hours'].sum().reset_index()
        fig_line = px.area(monthly, x='Month_Year', y='Duration_Hours', title="Tổng giờ vận hành theo Tháng", markers=True)
        st.plotly_chart(fig_line, use_container_width=True)
        
    with col_map:
        st.subheader("Nội thành vs Đi Tỉnh")
        loc_counts = df_final_filtered['Phạm Vi'].value_counts().reset_index()
        loc_counts.columns = ['Phạm Vi', 'Số chuyến']
        fig_donut = px.pie(loc_counts, values='Số chuyến', names='Phạm Vi', hole=0.6, color_discrete_sequence=['#3498db', '#f1c40f'])
        st.plotly_chart(fig_donut, use_container_width=True)

    # Heatmap Xe
    st.subheader("Hiệu suất sử dụng từng xe (Top 15)")
    if 'Biển số xe' in df_final_filtered.columns:
        car_ usage = df_final_filtered.groupby('Biển số xe')['Duration_Hours'].sum().reset_index().sort_values('Duration_Hours', ascending=False).head(15)
        fig_car = px.bar(car_usage, x='Biển số xe', y='Duration_Hours', color='Duration_Hours', title="Top 15 xe hoạt động nhiều nhất (Giờ)", color_continuous_scale='Viridis')
        st.plotly_chart(fig_car, use_container_width=True)