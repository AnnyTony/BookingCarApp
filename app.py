import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go

# --- 1. CẤU HÌNH TRANG & CSS (Làm đẹp giao diện) ---
st.set_page_config(page_title="Fleet Management Dashboard", page_icon="🚘", layout="wide")

# CSS tùy chỉnh để ẩn menu mặc định và làm đẹp metrics
st.markdown("""
<style>
    .main {background-color: #f8f9fa;}
    .stMetric {
        background-color: #ffffff;
        padding: 15px;
        border-radius: 10px;
        border: 1px solid #e0e0e0;
        box-shadow: 2px 2px 5px rgba(0,0,0,0.05);
    }
    div[data-testid="stMetricValue"] {font-size: 24px; color: #007bff;}
    h1, h2, h3 {font-family: 'Segoe UI', sans-serif; color: #2c3e50;}
</style>
""", unsafe_allow_html=True)

# --- 2. HEADER ---
col_head1, col_head2 = st.columns([3, 1])
with col_head1:
    st.title("🚘 Fleet Operations Dashboard")
    st.markdown("Hệ thống báo cáo & Giám sát hoạt động đội xe")
with col_head2:
    st.image("https://cdn-icons-png.flaticon.com/512/741/741407.png", width=80) # Logo giả lập
    st.caption("Last updated: Live")

st.divider()

# --- 3. UPLOAD DATA ---
uploaded_file = st.file_uploader("📂 Import Dữ liệu Booking (Excel/CSV)", type=['xlsx', 'csv'])

if uploaded_file is None:
    st.info("👈 Vui lòng tải file dữ liệu lên để xem báo cáo.")
    st.stop()

# --- XỬ LÝ DỮ LIỆU ---
@st.cache_data # Cache để tăng tốc độ load khi filter
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
if isinstance(df, str): # Nếu trả về chuỗi lỗi
    st.error(f"Lỗi dữ liệu: {df}")
    st.stop()

df_assigned = df.dropna(subset=['Biển số xe'])

# --- 4. SIDEBAR & FILTERS ---
with st.sidebar:
    st.header("🎛️ Bộ Lọc Điều Khiển")
    
    # Filter Thời gian
    min_date = df_assigned['Start_Datetime'].min().date()
    max_date = df_assigned['End_Datetime'].max().date()
    
    date_range = st.date_input("📅 Khoảng thời gian", value=(min_date, max_date), min_value=min_date, max_value=max_date)
    
    # Filter Xe
    all_cars = sorted(df_assigned['Biển số xe'].astype(str).unique())
    selected_cars = st.multiselect("🚘 Chọn xe", options=all_cars, default=all_cars, placeholder="Chọn biển số...")
    
    st.markdown("---")
    st.caption("Developed with Streamlit & Plotly")

# ÁP DỤNG FILTER
if len(date_range) == 2:
    start_d, end_d = date_range
    mask_date = (df_assigned['Start_Datetime'].dt.date >= start_d) & (df_assigned['Start_Datetime'].dt.date <= end_d)
else:
    mask_date = (df_assigned['Start_Datetime'].dt.date == date_range[0])

mask_car = df_assigned['Biển số xe'].isin(selected_cars)
df_filtered = df_assigned[mask_date & mask_car]

if df_filtered.empty:
    st.warning("⚠️ Không có dữ liệu nào khớp với bộ lọc hiện tại.")
    st.stop()

# --- 5. TÍNH TOÁN KPI ---
total_trips = len(df_filtered)
total_hours = df_filtered['Duration_Hours'].sum()
avg_duration = df_filtered['Duration_Hours'].mean()

# Tính Overlap
df_sorted = df_filtered.sort_values(by=['Biển số xe', 'Start_Datetime'])
df_sorted['Prev_End'] = df_sorted.groupby('Biển số xe')['End_Datetime'].shift(1)
overlaps = df_sorted[df_sorted['Start_Datetime'] < df_sorted['Prev_End']]
overlap_count = len(overlaps)
overlap_rate = (overlap_count / total_trips * 100) if total_trips > 0 else 0

# --- 6. DASHBOARD CHÍNH ---

# A. Hàng KPI Metrics
col1, col2, col3, col4 = st.columns(4)
col1.metric("Tổng Số Chuyến", f"{total_trips}", "chuyến")
col2.metric("Tổng Giờ Vận Hành", f"{total_hours:,.0f}", "giờ")
col3.metric("Thời Gian TB/Chuyến", f"{avg_duration:.1f}", "giờ")
col4.metric("Cảnh Báo Trùng (Overlap)", f"{overlap_count}", f"{overlap_rate:.1f}%", delta_color="inverse")

st.markdown("### 📈 Phân Tích Hiệu Suất")

# B. Hàng Biểu đồ 1 (Timeline & Xe)
c1, c2 = st.columns([2, 1])

with c1:
    # Biểu đồ cột theo tháng (Dùng Plotly)
    monthly_data = df_filtered.groupby('Month_Year')['Duration_Hours'].sum().reset_index()
    fig_month = px.bar(monthly_data, x='Month_Year', y='Duration_Hours', 
                       title="Tổng giờ hoạt động theo Tháng",
                       labels={'Month_Year': 'Tháng', 'Duration_Hours': 'Số giờ'},
                       color='Duration_Hours', color_continuous_scale='Blues')
    fig_month.update_layout(xaxis_title="", yaxis_title="Giờ", height=350)
    st.plotly_chart(fig_month, use_container_width=True)

with c2:
    # Biểu đồ Pie/Donut tỷ lệ xe
    car_counts = df_filtered['Biển số xe'].value_counts().reset_index()
    car_counts.columns = ['Biển số xe', 'Số chuyến']
    fig_pie = px.pie(car_counts.head(10), values='Số chuyến', names='Biển số xe', 
                     title="Top 10 Xe hoạt động nhiều nhất",
                     hole=0.4, color_discrete_sequence=px.colors.qualitative.Pastel)
    fig_pie.update_layout(height=350, showlegend=False)
    st.plotly_chart(fig_pie, use_container_width=True)

# C. Hàng Biểu đồ 2 (Heatmap & User)
st.markdown("### 👥 Phân Tích Người Dùng & Thời Điểm")
c3, c4 = st.columns([1, 1])

with c3:
    # Heatmap Ngày trong tuần vs Giờ
    # Tạo cột Giờ bắt đầu (làm tròn)
    df_filtered['Hour_Start'] = df_filtered['Start_Datetime'].dt.hour
    heatmap_data = df_filtered.groupby(['Day_Name', 'Hour_Start']).size().reset_index(name='Counts')
    # Sắp xếp thứ tự ngày
    days_order = ['Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Saturday', 'Sunday']
    
    fig_heat = px.density_heatmap(heatmap_data, x='Hour_Start', y='Day_Name', z='Counts',
                                  title="Bản đồ nhiệt: Mật độ đặt xe (Thứ vs Giờ)",
                                  category_orders={'Day_Name': days_order},
                                  color_continuous_scale='Viridis')
    fig_heat.update_layout(height=400)
    st.plotly_chart(fig_heat, use_container_width=True)

with c4:
    # Top User (Horizontal Bar)
    if 'Người sử dụng xe' in df_filtered.columns:
        user_data = df_filtered.groupby('Người sử dụng xe')['Duration_Hours'].sum().nlargest(10).reset_index()
        fig_user = px.bar(user_data, x='Duration_Hours', y='Người sử dụng xe', orientation='h',
                          title="Top 10 Người sử dụng (Theo giờ)",
                          text_auto='.0f',
                          color='Duration_Hours', color_continuous_scale='Sunset')
        fig_user.update_layout(yaxis={'categoryorder':'total ascending'}, height=400)
        st.plotly_chart(fig_user, use_container_width=True)
    else:
        st.warning("Thiếu cột 'Người sử dụng xe'")

# --- 7. CHI TIẾT OVERLAP (EXPANDER) ---
with st.expander("⚠️ Xem chi tiết Danh sách Xe bị trùng lịch (Overlap)", expanded=False):
    if overlap_count > 0:
        st.dataframe(
            overlaps[['Ngày khởi hành', 'Biển số xe', 'Tên tài xế', 'Start_Datetime', 'End_Datetime', 'Prev_End']]
            .style.background_gradient(cmap='Reds', subset=['Start_Datetime']),
            use_container_width=True
        )
    else:
        st.success("Không có trường hợp nào bị trùng lịch.")