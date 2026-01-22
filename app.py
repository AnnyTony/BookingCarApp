import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go

# --- 1. CẤU HÌNH TRANG & CSS (Làm đẹp giao diện) ---
st.set_page_config(page_title="Fleet Management Dashboard", page_icon="🚘", layout="wide")

# CSS tùy chỉnh: Chỉnh màu nền Sidebar, làm bo tròn các khung
st.markdown("""
<style>
    /* Chỉnh giao diện Sidebar */
    [data-testid="stSidebar"] {
        background-color: #f0f2f6;
    }
    [data-testid="stSidebar"] h1 {
        font-size: 20px;
        color: #1f77b4;
    }
    
    /* Chỉnh Metric Cards */
    div[data-testid="stMetricValue"] {
        font-size: 24px;
        color: #007bff;
        font-weight: bold;
    }
    
    /* Header chính */
    .main-header {
        font-family: 'Helvetica Neue', sans-serif;
        color: #2c3e50;
    }
</style>
""", unsafe_allow_html=True)

# --- 2. HEADER ---
# Logo và Tiêu đề
col_head1, col_head2 = st.columns([4, 1])
with col_head1:
    st.markdown("<h1 class='main-header'>🚘 Fleet Operations Center</h1>", unsafe_allow_html=True)
    st.markdown("Dashboard phân tích hiệu suất và điều phối đội xe")
with col_head2:
    # Bạn có thể thay link ảnh logo công ty bạn vào đây
    st.image("https://cdn-icons-png.flaticon.com/512/3097/3097180.png", width=70)

st.divider()

# --- 3. UPLOAD DATA ---
uploaded_file = st.file_uploader("📂 Import Dữ liệu Booking (Kéo thả file vào đây)", type=['xlsx', 'csv'])

if uploaded_file is None:
    st.info("👈 Vui lòng tải file dữ liệu lên để bắt đầu.")
    st.stop()

# --- XỬ LÝ DỮ LIỆU (Cache để chạy nhanh) ---
@st.cache_data 
def load_data(file):
    try:
        if file.name.endswith('.csv'):
            df = pd.read_csv(file)
        else:
            df = pd.read_excel(file, engine='openpyxl')
            
        # Xử lý ngày giờ
        df['Start_Datetime'] = pd.to_datetime(df['Ngày khởi hành'].astype(str) + ' ' + df['Giờ khởi hành'].astype(str), errors='coerce')
        df['End_Datetime'] = pd.to_datetime(df['Ngày khởi hành'].astype(str) + ' ' + df['Giờ kết thúc'].astype(str), errors='coerce')
        
        # Xử lý qua đêm
        mask_overnight = df['End_Datetime'] < df['Start_Datetime']
        df.loc[mask_overnight, 'End_Datetime'] += pd.Timedelta(days=1)
        
        df['Duration_Hours'] = (df['End_Datetime'] - df['Start_Datetime']).dt.total_seconds() / 3600
        df['Month_Year'] = df['Start_Datetime'].dt.to_period('M').astype(str)
        df['Day_Name'] = df['Start_Datetime'].dt.day_name()
        
        return df
    except Exception as e:
        return str(e)

df = load_data(uploaded_file)
if isinstance(df, str): 
    st.error(f"Lỗi dữ liệu: {df}")
    st.stop()

df_assigned = df.dropna(subset=['Biển số xe'])

# --- 4. SIDEBAR "XỊN" (ĐÃ NÂNG CẤP) ---
with st.sidebar:
    st.markdown("## 🎛️ Bảng Điều Khiển")
    
    # Gom nhóm 1: Thời gian
    with st.expander("📆 Lọc Thời Gian", expanded=True):
        min_date = df_assigned['Start_Datetime'].min().date()
        max_date = df_assigned['End_Datetime'].max().date()
        
        date_range = st.date_input(
            "Chọn khoảng ngày:",
            value=(min_date, max_date),
            min_value=min_date,
            max_value=max_date
        )

    # Gom nhóm 2: Xe (Có nút chọn tất cả)
    with st.expander("🚗 Lọc Theo Xe", expanded=True):
        all_cars = sorted(df_assigned['Biển số xe'].astype(str).unique())
        
        # Nút gạt chọn tất cả (Tiện lợi hơn multiselect thường)
        select_all_cars = st.toggle("Chọn tất cả xe", value=True)
        
        if select_all_cars:
            selected_cars = all_cars
            st.caption(f"Đang chọn toàn bộ {len(all_cars)} xe")
        else:
            selected_cars = st.multiselect(
                "Chọn xe cụ thể:",
                options=all_cars,
                default=all_cars[:5] # Mặc định chọn 5 xe đầu nếu bỏ tick all
            )

    # Nút Reset (Thực ra là reload trang)
    if st.button("🔄 Reset Bộ Lọc", type="primary", use_container_width=True):
        st.rerun()
    
    # Footer nhỏ
    st.markdown("---")
    st.markdown(f"**Dữ liệu gốc:** {len(df_assigned)} dòng")


# --- XỬ LÝ LOGIC LỌC ---
# 1. Lọc ngày
if isinstance(date_range, tuple) and len(date_range) == 2:
    start_d, end_d = date_range
    mask_date = (df_assigned['Start_Datetime'].dt.date >= start_d) & (df_assigned['Start_Datetime'].dt.date <= end_d)
elif isinstance(date_range, tuple) and len(date_range) == 1:
    mask_date = (df_assigned['Start_Datetime'].dt.date == date_range[0])
else:
    mask_date = pd.Series([True] * len(df_assigned)) # Fallback

# 2. Lọc xe
mask_car = df_assigned['Biển số xe'].isin(selected_cars)

# DataFrame cuối cùng
df_filtered = df_assigned[mask_date & mask_car]

# Hiển thị thông báo trạng thái ở Sidebar (Feedback loop)
st.sidebar.success(f"🔍 Tìm thấy: **{len(df_filtered)}** chuyến")

if df_filtered.empty:
    st.warning("⚠️ Không có dữ liệu nào khớp với bộ lọc hiện tại.")
    st.stop()

# --- 5. TÍNH TOÁN KPI ---
total_trips = len(df_filtered)
total_hours = df_filtered['Duration_Hours'].sum()
avg_duration = df_filtered['Duration_Hours'].mean()

# Overlap logic
df_sorted = df_filtered.sort_values(by=['Biển số xe', 'Start_Datetime'])
df_sorted['Prev_End'] = df_sorted.groupby('Biển số xe')['End_Datetime'].shift(1)
overlaps = df_sorted[df_sorted['Start_Datetime'] < df_sorted['Prev_End']]
overlap_count = len(overlaps)
overlap_rate = (overlap_count / total_trips * 100) if total_trips > 0 else 0

# --- 6. DASHBOARD CONTENT ---

# A. Metrics
col1, col2, col3, col4 = st.columns(4)
col1.metric("Tổng Số Chuyến", f"{total_trips}")
col2.metric("Tổng Giờ Vận Hành", f"{total_hours:,.0f}h")
col3.metric("TB Một Chuyến", f"{avg_duration:.1f}h")
col4.metric("Trùng Lịch (Overlap)", f"{overlap_count}", f"{overlap_rate:.1f}%", delta_color="inverse")

st.markdown("---")

# B. Tabs Biểu đồ
tab1, tab2, tab3 = st.tabs(["📊 Hiệu Suất Vận Hành", "👥 Phân Tích User", "⚠️ Cảnh Báo Trùng"])

with tab1:
    c1, c2 = st.columns([7, 3])
    with c1:
        # Biểu đồ diễn biến theo tháng
        monthly = df_filtered.groupby('Month_Year')['Duration_Hours'].sum().reset_index()
        fig_month = px.bar(monthly, x='Month_Year', y='Duration_Hours', 
                           title="Tổng giờ hoạt động theo Tháng",
                           text_auto='.0f',
                           color='Duration_Hours', color_continuous_scale='Blues')
        fig_month.update_layout(height=400, xaxis_title="", yaxis_title="")
        st.plotly_chart(fig_month, use_container_width=True)
    
    with c2:
        # Tỷ trọng xe
        car_counts = df_filtered['Biển số xe'].value_counts().reset_index().head(10)
        car_counts.columns = ['Xe', 'Số chuyến']
        fig_pie = px.pie(car_counts, values='Số chuyến', names='Xe', title="Top Xe hoạt động", hole=0.5)
        fig_pie.update_layout(height=400, showlegend=False)
        st.plotly_chart(fig_pie, use_container_width=True)

    # Heatmap
    st.subheader("Bản đồ nhiệt: Mật độ đặt xe")
    df_filtered['Hour'] = df_filtered['Start_Datetime'].dt.hour
    heatmap_data = df_filtered.groupby(['Day_Name', 'Hour']).size().reset_index(name='Count')
    days_order = ['Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Saturday', 'Sunday']
    
    fig_heat = px.density_heatmap(heatmap_data, x='Hour', y='Day_Name', z='Count',
                                  color_continuous_scale='Viridis',
                                  category_orders={'Day_Name': days_order})
    st.plotly_chart(fig_heat, use_container_width=True)

with tab2:
    if 'Người sử dụng xe' in df_filtered.columns:
        user_stats = df_filtered.groupby('Người sử dụng xe')['Duration_Hours'].sum().nlargest(15).sort_values()
        fig_user = px.bar(user_stats, x='Duration_Hours', y=user_stats.index, orientation='h',
                          title="Top 15 Người sử dụng nhiều nhất (Giờ)",
                          text_auto='.0f',
                          color='Duration_Hours', color_continuous_scale='Sunset')
        fig_user.update_layout(height=600, yaxis_title="")
        st.plotly_chart(fig_user, use_container_width=True)
    else:
        st.info("File dữ liệu không có cột 'Người sử dụng xe'")

with tab3:
    if overlap_count > 0:
        st.error(f"Phát hiện {overlap_count} trường hợp trùng lịch xe:")
        # Format lại bảng cho đẹp
        display_cols = ['Ngày khởi hành', 'Biển số xe', 'Tên tài xế', 'Start_Datetime', 'End_Datetime', 'Prev_End']
        st.dataframe(
            overlaps[display_cols].style.background_gradient(cmap='Reds', subset=['Start_Datetime']),
            use_container_width=True
        )
    else:
        st.success("✅ Không có chuyến xe nào bị trùng giờ trong bộ lọc hiện tại.")