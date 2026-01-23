import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import numpy as np

# --- 1. CẤU HÌNH TRANG ---
st.set_page_config(page_title="Fleet Management System", page_icon="🏢", layout="wide")

# CSS làm đẹp
st.markdown("""
<style>
    .main-header {font-size: 28px; font-weight: bold; color: #2c3e50;}
    .kpi-card {background-color: #f8f9fa; padding: 15px; border-radius: 10px; border: 1px solid #e9ecef;}
    [data-testid="stSidebar"] {background-color: #f0f2f6;}
</style>
""", unsafe_allow_html=True)

st.markdown("<div class='main-header'>🏢 Dashboard Quản Lý Đội Xe (Multi-Tab)</div>", unsafe_allow_html=True)
st.markdown("---")

# --- 2. HÀM XỬ LÝ DỮ LIỆU ĐA TAB ---
@st.cache_data
def load_data_multisheet(file):
    try:
        # Đọc file Excel (Load cả 2 sheet cần thiết)
        # Lưu ý: Tên sheet phải khớp với file Excel của bạn
        xl = pd.ExcelFile(file, engine='openpyxl')
        
        # Kiểm tra tên sheet (phòng trường hợp user đặt tên khác chút xíu)
        sheet_names = xl.sheet_names
        sheet_booking = next((s for s in sheet_names if 'booking' in s.lower()), None)
        sheet_cbnv = next((s for s in sheet_names if 'cbnv' in s.lower() or 'staff' in s.lower()), None)
        
        if not sheet_booking:
            return "Không tìm thấy Sheet 'Booking car' (hoặc tên tương tự)."
            
        # 1. Load Booking Data
        df_bk = xl.parse(sheet_booking)
        df_bk.columns = df_bk.columns.str.strip() # Xóa khoảng trắng tên cột
        
        # Xử lý ngày giờ
        df_bk['Start_Datetime'] = pd.to_datetime(df_bk['Ngày khởi hành'].astype(str) + ' ' + df_bk['Giờ khởi hành'].astype(str), errors='coerce')
        df_bk['End_Datetime'] = pd.to_datetime(df_bk['Ngày khởi hành'].astype(str) + ' ' + df_bk['Giờ kết thúc'].astype(str), errors='coerce')
        
        # Xử lý qua đêm
        mask_overnight = df_bk['End_Datetime'] < df_bk['Start_Datetime']
        df_bk.loc[mask_overnight, 'End_Datetime'] += pd.Timedelta(days=1)
        
        df_bk['Duration_Hours'] = (df_bk['End_Datetime'] - df_bk['Start_Datetime']).dt.total_seconds() / 3600
        df_bk['Month_Year'] = df_bk['Start_Datetime'].dt.to_period('M').astype(str)
        df_bk['Year'] = df_bk['Start_Datetime'].dt.year
        
        # Phân loại Nửa ngày / Cả ngày
        df_bk['Loại Chuyến'] = df_bk['Duration_Hours'].apply(lambda x: 'Nửa ngày (<=4h)' if x <= 4 else 'Cả ngày (>4h)')
        
        # Phân loại Tỉnh / Nội thành (Dựa trên Lộ trình - Heuristic cơ bản)
        def classify_location(route):
            route = str(route).lower()
            keywords_tinh = ['tỉnh', 'tp.', 'bình dương', 'đồng nai', 'vũng tàu', 'long an', 'hà nội', 'bắc ninh', 'hải phòng']
            # Logic: Nếu lộ trình chứa từ khóa tỉnh -> Tỉnh, ngược lại Nội thành
            # Lưu ý: Đây là logic tương đối, chính xác nhất là cần cột dữ liệu chuẩn từ user
            if any(k in route for k in keywords_tinh):
                return "Đi Tỉnh"
            return "Nội thành"
        
        if 'Lộ trình' in df_bk.columns:
            df_bk['Phạm Vi'] = df_bk['Lộ trình'].apply(classify_location)
        else:
            df_bk['Phạm Vi'] = "Không xác định"

        # 2. Load CBNV Data & Merge (Vlookup)
        if sheet_cbnv:
            df_staff = xl.parse(sheet_cbnv)
            df_staff.columns = df_staff.columns.str.strip()
            
            # Chọn các cột cần thiết từ bảng CBNV để merge
            # Giả sử bảng CBNV có cột: 'Full Name', 'Công ty', 'BU', 'Location'
            # Cần chuẩn hóa tên cột CBNV cho khớp code
            col_mapping = {}
            for c in df_staff.columns:
                if 'name' in c.lower(): col_mapping[c] = 'Full Name'
                if 'công ty' in c.lower() or 'company' in c.lower(): col_mapping[c] = 'Công ty_Lookup'
                if 'bu' in c.lower() or 'bộ phận' in c.lower(): col_mapping[c] = 'BoPhan_Lookup'
                if 'location' in c.lower() or 'site' in c.lower(): col_mapping[c] = 'Location_Lookup'
            
            df_staff = df_staff.rename(columns=col_mapping)
            
            # Merge: Booking join với Staff qua tên người dùng
            # Left join để giữ lại toàn bộ booking dù không tìm thấy nhân viên
            df_final = pd.merge(df_bk, df_staff[['Full Name', 'Công ty_Lookup', 'BoPhan_Lookup', 'Location_Lookup']], 
                                left_on='Người sử dụng xe', right_on='Full Name', how='left')
            
            # Ưu tiên lấy dữ liệu từ Lookup, nếu không có thì lấy từ file Booking gốc (nếu có)
            df_final['Công ty'] = df_final['Công ty_Lookup'].fillna('Khác')
            df_final['Bộ phận'] = df_final['BoPhan_Lookup'].fillna('Khác')
            
            # Xử lý Vùng miền (Bắc / Nam) dựa trên Location
            # Giả định: HCM -> Nam, HN -> Bắc
            def get_region(loc):
                loc = str(loc).upper()
                if 'HCM' in loc or 'NAM' in loc: return 'Miền Nam'
                if 'HN' in loc or 'BẮC' in loc or 'HANOI' in loc: return 'Miền Bắc'
                return 'Khác'
            
            df_final['Vùng Miền'] = df_final['Location_Lookup'].apply(get_region)
            
        else:
            df_final = df_bk
            df_final['Công ty'] = "Không có dữ liệu CBNV"
            df_final['Bộ phận'] = "Không có dữ liệu CBNV"
            df_final['Vùng Miền'] = "Khác"

        return df_final

    except Exception as e:
        return f"Lỗi chi tiết: {str(e)}"

# --- 3. UPLOAD ---
uploaded_file = st.file_uploader("📂 Upload file Excel (Chứa cả tab Booking và CBNV)", type=['xlsx'])

if uploaded_file:
    df = load_data_multisheet(uploaded_file)
    
    if isinstance(df, str): # Nếu trả về chuỗi là lỗi
        st.error(df)
        st.stop()
        
    # --- 4. SIDEBAR FILTERS ---
    with st.sidebar:
        st.header("🔍 Bộ Lọc Dữ Liệu")
        
        # Lọc Vùng Miền (Quan trọng để tính tổng xe)
        all_regions = ['Tất cả'] + sorted(list(df['Vùng Miền'].unique()))
        region_filter = st.selectbox("🌍 Vùng Miền:", all_regions, index=0)
        
        # Lọc Năm
        all_years = sorted(df['Year'].dropna().unique())
        year_filter = st.multiselect("📅 Năm:", all_years, default=all_years)
        
        # Lọc Công Ty
        all_companies = sorted(df['Công ty'].astype(str).unique())
        comp_filter = st.multiselect("🏢 Công ty:", all_companies, default=all_companies)
        
        # Áp dụng lọc
        df_filtered = df.copy()
        
        # Logic lọc vùng
        if region_filter != 'Tất cả':
            df_filtered = df_filtered[df_filtered['Vùng Miền'] == region_filter]
            
        # Logic lọc năm & công ty
        if year_filter:
            df_filtered = df_filtered[df_filtered['Year'].isin(year_filter)]
        if comp_filter:
            df_filtered = df_filtered[df_filtered['Công ty'].isin(comp_filter)]

        st.success(f"Hiển thị: {len(df_filtered)} chuyến")

    # --- 5. TÍNH TOÁN KPI OCCUPANCY (TỶ LỆ LẤP ĐẦY) ---
    # Logic: 
    # Miền Nam: 16 xe, Miền Bắc: 5 xe. 
    # Nếu chọn Tất cả: 21 xe.
    # Số giờ khả dụng (Capacity) = Số xe * Số ngày trong khoảng lọc * 9 tiếng/ngày (Giả định)
    
    if region_filter == 'Miền Nam': total_cars = 16
    elif region_filter == 'Miền Bắc': total_cars = 5
    else: total_cars = 21 # Tổng cả 2 miền

    # Tính số ngày trong dữ liệu lọc (để tính mẫu số)
    if not df_filtered.empty and 'Start_Datetime' in df_filtered.columns:
        date_min = df_filtered['Start_Datetime'].min()
        date_max = df_filtered['Start_Datetime'].max()
        days_diff = (date_max - date_min).days + 1
        if days_diff <= 0: days_diff = 1
        
        # Capacity (Giờ) = Số xe * Số ngày * 9h (Giờ hành chính)
        capacity_hours = total_cars * days_diff * 9
        used_hours = df_filtered['Duration_Hours'].sum()
        
        occupancy_rate = (used_hours / capacity_hours * 100) if capacity_hours > 0 else 0
    else:
        occupancy_rate = 0
        days_diff = 0

    # --- 6. DASHBOARD CHÍNH ---
    
    # KPI Cards
    c1, c2, c3, c4 = st.columns(4)
    c1.metric("Tổng số chuyến", len(df_filtered))
    c2.metric("Tổng giờ vận hành", f"{used_hours:,.0f}h")
    c3.metric("Số xe khả dụng", f"{total_cars} xe")
    c4.metric("Tỷ lệ Lấp đầy (Occupancy)", f"{occupancy_rate:.1f}%", help=f"Tính trên {total_cars} xe trong {days_diff} ngày (9h/ngày)")

    st.markdown("---")

    # TAB 1: TRẠNG THÁI & HIỆU SUẤT
    t1, t2, t3, t4 = st.tabs(["📊 Trạng Thái & Loại Chuyến", "🏢 Công Ty & Phòng Ban", "🗺️ Lộ Trình & Xe", "📈 Xu Hướng (Time)"])
    
    with t1:
        col_st1, col_st2 = st.columns(2)
        with col_st1:
            # 1. Tổng số chuyến hoàn thành, cancel, reject
            if 'Tình trạng đơn yêu cầu' in df_filtered.columns:
                st.subheader("Tỷ lệ Trạng thái chuyến đi")
                status_counts = df_filtered['Tình trạng đơn yêu cầu'].fillna('Unknown').value_counts().reset_index()
                status_counts.columns = ['Status', 'Count']
                
                # Tô màu đặc biệt cho Cancel/Reject
                color_map = {'CLOSED': '#2ecc71', 'APPROVED': '#3498db', 'CANCELLED': '#e74c3c', 'REJECTED': '#c0392b'}
                fig_status = px.pie(status_counts, values='Count', names='Status', hole=0.4, 
                                    color='Status', color_discrete_map=color_map)
                st.plotly_chart(fig_status, use_container_width=True)
                
                # Hiển thị số liệu chi tiết
                st.dataframe(status_counts, use_container_width=True)
                
        with col_st2:
            # 2. Tỷ lệ Nửa ngày vs Cả ngày
            st.subheader("Loại chuyến (Nửa ngày vs Cả ngày)")
            type_counts = df_filtered['Loại Chuyến'].value_counts().reset_index()
            type_counts.columns = ['Loại', 'Số chuyến']
            fig_type = px.bar(type_counts, x='Loại', y='Số chuyến', text_auto=True, color='Loại')
            st.plotly_chart(fig_type, use_container_width=True)

    with t2:
        # 3. Tỷ lệ theo Công ty
        st.subheader("Phân bổ chuyến đi theo Công ty")
        comp_counts = df_filtered['Công ty'].value_counts().reset_index()
        comp_counts.columns = ['Công ty', 'Số chuyến']
        fig_comp = px.bar(comp_counts, x='Số chuyến', y='Công ty', orientation='h', 
                          text_auto=True, color='Số chuyến', color_continuous_scale='Viridis')
        fig_comp.update_layout(yaxis={'categoryorder':'total ascending'})
        st.plotly_chart(fig_comp, use_container_width=True)
        
        st.divider()
        
        # 4. Tỷ lệ theo Bộ phận (Drill down)
        st.subheader("Chi tiết theo Bộ phận")
        dept_counts = df_filtered.groupby(['Công ty', 'Bộ phận']).size().reset_index(name='Số chuyến')
        fig_sun = px.sunburst(dept_counts, path=['Công ty', 'Bộ phận'], values='Số chuyến', 
                              title="Biểu đồ Sunburst: Công ty > Bộ phận")
        st.plotly_chart(fig_sun, use_container_width=True)

    with t3:
        col_loc1, col_loc2 = st.columns(2)
        with col_loc1:
            # 5. Nội thành vs Tỉnh
            st.subheader("Phạm vi di chuyển")
            scope_counts = df_filtered['Phạm Vi'].value_counts().reset_index()
            scope_counts.columns = ['Phạm Vi', 'Số chuyến']
            fig_scope = px.pie(scope_counts, values='Số chuyến', names='Phạm Vi', title="Nội thành vs Đi Tỉnh")
            st.plotly_chart(fig_scope, use_container_width=True)
            
        with col_loc2:
            # 6. Tỷ lệ xe sử dụng
            st.subheader("Tần suất sử dụng các xe")
            if 'Biển số xe' in df_filtered.columns:
                car_stats = df_filtered['Biển số xe'].value_counts().reset_index().head(10)
                car_stats.columns = ['Biển số xe', 'Số chuyến']
                fig_car = px.bar(car_stats, x='Biển số xe', y='Số chuyến', color='Số chuyến')
                st.plotly_chart(fig_car, use_container_width=True)

    with t4:
        # 7. Occupancy Rate theo thời gian
        st.subheader("Tỷ lệ sử dụng xe theo Tháng")
        
        # Gom nhóm theo tháng
        monthly_stats = df_filtered.groupby('Month_Year').agg(
            Total_Hours=('Duration_Hours', 'sum'),
            Days_Count=('Start_Datetime', lambda x: x.dt.day.nunique()) # Số ngày có chạy trong tháng
        ).reset_index()
        
        # Tính Capacity tháng đó (Số xe * 26 ngày công chuẩn * 9h) - Hoặc tính theo ngày thực tế
        # Ở đây lấy ước lượng 26 ngày làm việc/tháng cho đơn giản
        monthly_capacity = total_cars * 26 * 9 
        
        monthly_stats['Occupancy_%'] = (monthly_stats['Total_Hours'] / monthly_capacity * 100).clip(upper=100)
        
        fig_occ = px.line(monthly_stats, x='Month_Year', y='Occupancy_%', markers=True, 
                          title=f"Tỷ lệ lấp đầy theo tháng (Giả định {total_cars} xe, 26 ngày công/tháng)",
                          labels={'Occupancy_%': 'Tỷ lệ lấp đầy (%)'})
        
        # Thêm biểu đồ cột số chuyến chồng bên dưới
        fig_occ.add_bar(x=monthly_stats['Month_Year'], y=monthly_stats['Total_Hours'], name='Tổng giờ chạy', opacity=0.3, yaxis='y2')
        
        st.plotly_chart(fig_occ, use_container_width=True)
        
        st.subheader("Heatmap: Mật độ sử dụng trong tuần")
        df_filtered['Weekday'] = df_filtered['Start_Datetime'].dt.day_name()
        df_filtered['Hour'] = df_filtered['Start_Datetime'].dt.hour
        
        heat_data = df_filtered.groupby(['Weekday', 'Hour']).size().reset_index(name='Count')
        days_order = ['Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Saturday', 'Sunday']
        
        fig_heat = px.density_heatmap(heat_data, x='Hour', y='Weekday', z='Count', 
                                      category_orders={'Weekday': days_order},
                                      color_continuous_scale='RdBu_r')
        st.plotly_chart(fig_heat, use_container_width=True)

else:
    st.info("👋 Vui lòng upload file Excel chứa sheet 'Booking car' và 'CBNV'.")