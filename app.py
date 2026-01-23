import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go

# --- 1. CẤU HÌNH TRANG ---
st.set_page_config(page_title="Power BI Style Dashboard", page_icon="📊", layout="wide")

st.markdown("""
<style>
    .main-header {font-size: 26px; font-weight: bold; color: #2c3e50;}
    div[data-testid="stMetricValue"] {font-size: 22px; color: #2980b9;}
    [data-testid="stSidebar"] {background-color: #f1f3f6;}
    .stTabs [data-baseweb="tab-list"] {gap: 10px;}
    .stTabs [data-baseweb="tab"] {height: 50px; background-color: white; border-radius: 4px; box-shadow: 0px 1px 3px rgba(0,0,0,0.1);}
    .stTabs [aria-selected="true"] {background-color: #e3f2fd; color: #1976d2;}
</style>
""", unsafe_allow_html=True)

st.markdown("<div class='main-header'>📊 Fleet Management Intelligence (Power BI Style)</div>", unsafe_allow_html=True)
st.markdown("---")

# --- 2. HÀM LOAD DATA (ĐÃ SỬA LỖI HEADER) ---
@st.cache_data
def load_data_pro(file):
    try:
        xl = pd.ExcelFile(file, engine='openpyxl')
        
        # Tìm tên sheet linh hoạt
        sheet_booking = next((s for s in xl.sheet_names if 'booking' in s.lower()), None)
        sheet_cbnv = next((s for s in xl.sheet_names if 'cbnv' in s.lower() or 'staff' in s.lower()), None)
        
        if not sheet_booking: return "Lỗi: Không tìm thấy sheet chứa dữ liệu Booking."

        # === HÀM PHỤ: TÌM DÒNG TIÊU ĐỀ ===
        def find_header_and_read(excel_file, sheet_name, keywords):
            # Đọc thử 5 dòng đầu để tìm header
            df_temp = excel_file.parse(sheet_name, header=None, nrows=10)
            header_idx = 0
            found = False
            
            for i, row in df_temp.iterrows():
                row_str = row.astype(str).str.lower().tolist()
                # Nếu dòng này chứa từ khóa quan trọng (vd: 'full name', 'biển số xe')
                if any(k in ' '.join(row_str) for k in keywords):
                    header_idx = i
                    found = True
                    break
            
            # Đọc lại với header đúng
            return excel_file.parse(sheet_name, header=header_idx)

        # 1. XỬ LÝ SHEET BOOKING
        df_bk = find_header_and_read(xl, sheet_booking, ['ngày khởi hành', 'biển số', 'date'])
        df_bk.columns = df_bk.columns.str.strip() # Xóa khoảng trắng thừa
        
        # Xử lý ngày giờ
        try:
            df_bk['Start_Datetime'] = pd.to_datetime(df_bk['Ngày khởi hành'].astype(str) + ' ' + df_bk['Giờ khởi hành'].astype(str), errors='coerce')
            df_bk['End_Datetime'] = pd.to_datetime(df_bk['Ngày khởi hành'].astype(str) + ' ' + df_bk['Giờ kết thúc'].astype(str), errors='coerce')
            
            # Xử lý qua đêm
            mask_overnight = df_bk['End_Datetime'] < df_bk['Start_Datetime']
            df_bk.loc[mask_overnight, 'End_Datetime'] += pd.Timedelta(days=1)
            
            df_bk['Duration_Hours'] = (df_bk['End_Datetime'] - df_bk['Start_Datetime']).dt.total_seconds() / 3600
            df_bk['Month_Year'] = df_bk['Start_Datetime'].dt.to_period('M').astype(str)
            df_bk['Year'] = df_bk['Start_Datetime'].dt.year
        except:
            pass # Bỏ qua nếu lỗi ngày tháng để app vẫn chạy

        df_bk['Loại Chuyến'] = df_bk['Duration_Hours'].apply(lambda x: 'Nửa ngày' if x <= 4 else 'Cả ngày')

        # Logic Đi Tỉnh / Nội Thành
        def check_scope(route):
            s = str(route).lower()
            if any(x in s for x in ['tỉnh', 'tp.', 'bình dương', 'đồng nai', 'vũng tàu', 'long an', 'tiền giang', 'bắc ninh']):
                return "Đi Tỉnh"
            return "Nội thành"
        
        if 'Lộ trình' in df_bk.columns:
            df_bk['Phạm Vi'] = df_bk['Lộ trình'].apply(check_scope) 
        else:
            df_bk['Phạm Vi'] = "Unknown"

        # 2. XỬ LÝ SHEET CBNV & MERGE
        if sheet_cbnv:
            # Tự tìm header có chữ 'Full Name' hoặc 'Họ tên'
            df_staff = find_header_and_read(xl, sheet_cbnv, ['full name', 'họ tên', 'email', 'công ty'])
            df_staff.columns = df_staff.columns.str.strip()
            
            # Map tên cột (Chuẩn hóa)
            col_map = {}
            for c in df_staff.columns:
                c_lower = c.lower()
                if 'full name' in c_lower or 'họ tên' in c_lower: col_map[c] = 'Full Name'
                elif 'công ty' in c_lower or 'company' in c_lower: col_map[c] = 'Công ty_L'
                elif 'bu' in c_lower or 'bộ phận' in c_lower: col_map[c] = 'BoPhan_L'
                elif 'location' in c_lower or 'site' in c_lower: col_map[c] = 'Location_L'
            
            df_staff = df_staff.rename(columns=col_map)
            
            # Kiểm tra xem đã map đủ cột chưa, nếu thiếu thì tạo cột rỗng để không bị lỗi Key Error
            for req_col in ['Full Name', 'Công ty_L', 'BoPhan_L', 'Location_L']:
                if req_col not in df_staff.columns:
                    df_staff[req_col] = "Unknown"

            # Merge Booking với Staff
            df_final = pd.merge(df_bk, df_staff[['Full Name', 'Công ty_L', 'BoPhan_L', 'Location_L']], 
                                left_on='Người sử dụng xe', right_on='Full Name', how='left')
            
            # Điền dữ liệu
            df_final['Công ty'] = df_final['Công ty_L'].fillna('Chưa xác định')
            df_final['Bộ phận'] = df_final['BoPhan_L'].fillna('Chưa xác định')
            
            # Logic Bắc/Nam
            def get_region(loc):
                loc = str(loc).upper()
                if 'HCM' in loc or 'NAM' in loc or 'HO CHI MINH' in loc: return 'Miền Nam'
                if 'HN' in loc or 'BẮC' in loc or 'HANOI' in loc: return 'Miền Bắc'
                return 'Khác'
            
            df_final['Vùng Miền'] = df_final['Location_L'].apply(get_region)
            
        else:
            df_final = df_bk
            df_final['Công ty'] = "No Data"
            df_final['Bộ phận'] = "No Data"
            df_final['Vùng Miền'] = "Khác"
            
        return df_final

    except Exception as e:
        return f"Lỗi xử lý file: {str(e)}"

# --- 3. UPLOAD ---
uploaded_file = st.file_uploader("📂 Upload file Excel (Booking + CBNV)", type=['xlsx'])
if not uploaded_file:
    st.info("👋 Vui lòng tải file dữ liệu để bắt đầu.")
    st.stop()

df = load_data_pro(uploaded_file)
if isinstance(df, str):
    st.error(df)
    st.stop()

# --- 4. SIDEBAR "CASCADING" (BỘ LỌC) ---
with st.sidebar:
    st.header("🎛️ Bộ lọc Điều khiển")
    
    # 1. Chọn Năm
    if 'Year' in df.columns:
        years = sorted(df['Year'].dropna().unique())
        selected_years = st.multiselect("Năm:", years, default=years)
        df_lv1 = df[df['Year'].isin(selected_years)]
    else:
        df_lv1 = df
    
    # 2. Chọn Vùng Miền
    if 'Vùng Miền' in df_lv1.columns:
        regions = ['Tất cả'] + sorted(list(df_lv1['Vùng Miền'].unique()))
        selected_region = st.selectbox("Vùng Miền:", regions)
        if selected_region != 'Tất cả':
            df_lv2 = df_lv1[df_lv1['Vùng Miền'] == selected_region]
        else:
            df_lv2 = df_lv1
    else:
        df_lv2 = df_lv1
        selected_region = 'Khác'
        
    # 3. Chọn Công Ty
    avail_companies = sorted(df_lv2['Công ty'].astype(str).unique())
    selected_companies = st.multiselect("Công ty:", avail_companies, default=avail_companies)
    
    # 4. Chọn Bộ Phận
    if selected_companies:
        df_lv3 = df_lv2[df_lv2['Công ty'].isin(selected_companies)]
    else:
        df_lv3 = df_lv2
        
    avail_depts = sorted(df_lv3['Bộ phận'].astype(str).unique())
    selected_depts = st.multiselect("Phòng ban/Bộ phận:", avail_depts, default=avail_depts)

    # Filter Final
    if selected_depts:
        df_final_filtered = df_lv3[df_lv3['Bộ phận'].isin(selected_depts)]
    else:
        df_final_filtered = df_lv3
        
    st.success(f"🔍 Dữ liệu: {len(df_final_filtered)} chuyến")

# --- 5. TÍNH KPI ---
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
    occupancy = 0; used_hours = 0

# --- 6. DASHBOARD ---
c1, c2, c3, c4 = st.columns(4)
c1.metric("Tổng Số Chuyến", len(df_final_filtered))
c2.metric("Tổng Giờ Chạy", f"{used_hours:,.0f}h")
c3.metric("Tỷ lệ Lấp Đầy", f"{occupancy:.1f}%")
c4.metric("Số Xe Khả Dụng", f"{total_cars} xe")

st.markdown("---")

t1, t2 = st.tabs(["🏢 Cấu Trúc & Phân Bổ", "📈 Xu Hướng & Hiệu Suất"])

with t1:
    col_sun, col_tree = st.columns([1, 1])
    with col_sun:
        st.subheader("Phân bổ: Vùng > Công Ty > Bộ Phận")
        df_sun = df_final_filtered.groupby(['Vùng Miền', 'Công ty', 'Bộ phận']).size().reset_index(name='Số chuyến')
        df_sun = df_sun[df_sun['Số chuyến'] > 0]
        fig_sun = px.sunburst(df_sun, path=['Vùng Miền', 'Công ty', 'Bộ phận'], values='Số chuyến', color='Số chuyến', color_continuous_scale='RdBu')
        st.plotly_chart(fig_sun, use_container_width=True)
        st.caption("💡 Click vào vòng tròn để xem chi tiết.")

    with col_tree:
        st.subheader("Trạng thái chuyến đi")
        if 'Tình trạng đơn yêu cầu' in df_final_filtered.columns:
            status_df = df_final_filtered['Tình trạng đơn yêu cầu'].fillna('Unknown').value_counts().reset_index()
            status_df.columns = ['Status', 'Count']
            color_map = {'CLOSED': 'green', 'APPROVED': 'blue', 'CANCELLED': 'red', 'REJECTED': 'darkred'}
            fig_pie = px.pie(status_df, values='Count', names='Status', hole=0.5, color='Status', color_discrete_map=color_map)
            st.plotly_chart(fig_pie, use_container_width=True)

    st.subheader("Nửa ngày vs Cả ngày theo Công ty")
    df_type = df_final_filtered.groupby(['Công ty', 'Loại Chuyến']).size().reset_index(name='Count')
    fig_bar = px.bar(df_type, x='Công ty', y='Count', color='Loại Chuyến', barmode='group')
    st.plotly_chart(fig_bar, use_container_width=True)

with t2:
    col_trend, col_map = st.columns([2, 1])
    with col_trend:
        st.subheader("Xu Hướng theo Tháng")
        if 'Month_Year' in df_final_filtered.columns:
            monthly = df_final_filtered.groupby('Month_Year')['Duration_Hours'].sum().reset_index()
            fig_line = px.area(monthly, x='Month_Year', y='Duration_Hours', markers=True)
            st.plotly_chart(fig_line, use_container_width=True)
    
    with col_map:
        st.subheader("Nội thành vs Đi Tỉnh")
        loc_counts = df_final_filtered['Phạm Vi'].value_counts().reset_index()
        loc_counts.columns = ['Phạm Vi', 'Số chuyến']
        fig_donut = px.pie(loc_counts, values='Số chuyến', names='Phạm Vi', hole=0.6)
        st.plotly_chart(fig_donut, use_container_width=True)

    st.subheader("Top 15 Xe hoạt động nhiều nhất")
    if 'Biển số xe' in df_final_filtered.columns:
        car_usage = df_final_filtered.groupby('Biển số xe')['Duration_Hours'].sum().reset_index().sort_values('Duration_Hours', ascending=False).head(15)
        fig_car = px.bar(car_usage, x='Biển số xe', y='Duration_Hours', color='Duration_Hours', color_continuous_scale='Viridis')
        st.plotly_chart(fig_car, use_container_width=True)