import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
from datetime import datetime

# --- 1. CẤU HÌNH TRANG ---
st.set_page_config(
    page_title="Dashboard Quản Trị Đội Xe (Booking Car)",
    page_icon="🚘",
    layout="wide"
)

# CSS Styling để làm đẹp giao diện
st.markdown("""
<style>
    .kpi-card {
        background-color: #ffffff;
        border-radius: 10px;
        padding: 20px;
        box-shadow: 0 4px 6px rgba(0,0,0,0.1);
        text-align: center;
        border: 1px solid #e0e0e0;
    }
    .kpi-title {
        font-size: 14px;
        color: #6c757d;
        font-weight: 600;
        text-transform: uppercase;
        margin-bottom: 5px;
    }
    .kpi-value {
        font-size: 28px;
        font-weight: 800;
        color: #2c3e50;
    }
    .kpi-unit {
        font-size: 12px;
        color: #999;
    }
</style>
""", unsafe_allow_html=True)

# --- 2. HÀM XỬ LÝ DỮ LIỆU ---
@st.cache_data
def load_data(file):
    try:
        # 1. ĐỌC FILE VÀ TÌM SHEET 'BOOKING CAR'
        # Nếu là file Excel, cố gắng tìm sheet có tên chứa chữ "Booking"
        if file.name.endswith('.xlsx'):
            xl = pd.ExcelFile(file)
            sheet_names = xl.sheet_names
            
            # Tìm tên sheet phù hợp (không phân biệt hoa thường)
            target_sheet = next((s for s in sheet_names if "booking" in s.lower() and "car" in s.lower()), None)
            
            if target_sheet:
                # Quan trọng: header=3 để bỏ qua 3 dòng trống đầu tiên
                df = pd.read_excel(file, sheet_name=target_sheet, header=3)
            else:
                # Nếu không tìm thấy sheet tên Booking Car, đọc sheet đầu tiên và cảnh báo
                st.warning(f"Không tìm thấy Sheet 'Booking Car'. Đang đọc sheet đầu tiên: '{sheet_names[0]}'. Hãy kiểm tra lại cấu trúc file nếu dữ liệu sai.")
                df = pd.read_excel(file, sheet_name=0, header=3)
        
        elif file.name.endswith('.csv'):
            # Đọc CSV với header ở dòng 4 (index 3)
            df = pd.read_csv(file, header=3)
        else:
            return None

        # 2. CHUẨN HÓA TÊN CỘT (Xóa khoảng trắng, xuống dòng trong tên cột)
        df.columns = [str(c).strip().replace('\n', ' ') for c in df.columns]

        # 3. MAPPING CỘT (Đổi tên cột tiếng Việt sang tên biến tiếng Anh dễ xử lý)
        # Kiểm tra xem các cột quan trọng có tồn tại không
        col_map = {
            'Ngày Tháng Năm': 'Date',
            'Biển số xe': 'Car_Plate',
            'Tên tài xế': 'Driver',
            'Bộ phận': 'Department',
            'Km sử dụng': 'Km_Used',
            'Tổng chi phí': 'Total_Cost',
            'Lộ trình': 'Route',
            'Người sử dụng xe': 'User',
            'Giờ khởi hành': 'Start_Time',
            'Giờ kết thúc': 'End_Time'
        }
        
        # Chỉ lấy các cột có trong dữ liệu
        available_cols = [c for c in col_map.keys() if c in df.columns]
        df = df[available_cols].rename(columns=col_map)
        
        # Loại bỏ các dòng hoàn toàn trống
        df.dropna(how='all', inplace=True)
        # Loại bỏ các dòng mà ngày tháng bị rỗng (thường là dòng tổng cộng hoặc rác ở cuối)
        if 'Date' in df.columns:
            df = df.dropna(subset=['Date'])

        # 4. XỬ LÝ DỮ LIỆU CHI TIẾT
        
        # A. Xử lý Ngày Tháng
        if 'Date' in df.columns:
            df['Date'] = pd.to_datetime(df['Date'], errors='coerce')
            df = df.dropna(subset=['Date']) # Bỏ dòng nếu không convert được ngày
            df['Month_Str'] = df['Date'].dt.strftime('%m-%Y') # Dạng chuỗi cho bộ lọc
            df['Year_Month'] = df['Date'].dt.to_period('M')   # Dạng Period để sort đúng

        # B. Xử lý Cột Bộ Phận (Quan trọng: Xóa khoảng trắng thừa)
        if 'Department' in df.columns:
            df['Department'] = df['Department'].astype(str).str.strip()
            # Có thể thêm bước viết hoa chữ cái đầu hoặc viết hoa toàn bộ để đồng nhất
            # df['Department'] = df['Department'].str.upper() 

        # C. Xử lý Số Liệu (Chi phí & KM) - Chuyển text sang số
        for col in ['Total_Cost', 'Km_Used']:
            if col in df.columns:
                # Chuyển về dạng số, nếu lỗi biến thành NaN, sau đó fill bằng 0
                df[col] = pd.to_numeric(df[col], errors='coerce').fillna(0)

        return df

    except Exception as e:
        st.error(f"Có lỗi khi xử lý file: {e}")
        return pd.DataFrame()

# --- 3. GIAO DIỆN CHÍNH ---

st.title("📊 Dashboard Quản Lý Đội Xe")
st.markdown("Hệ thống phân tích dữ liệu từ Tab Booking Car")

# Upload File
uploaded_file = st.file_uploader("📂 Tải lên file Excel quản lý xe (File có Tab 'Booking Car')", type=['xlsx', 'csv'])

if uploaded_file is not None:
    # Load dữ liệu
    with st.spinner('Đang xử lý dữ liệu...'):
        df = load_data(uploaded_file)

    if df is not None and not df.empty:
        # --- SIDEBAR: BỘ LỌC ---
        st.sidebar.header("🔍 Bộ Lọc Dữ Liệu")
        
        # 1. Lọc theo Tháng
        all_months = sorted(df['Month_Str'].unique())
        selected_months = st.sidebar.multiselect("Chọn Tháng", all_months, default=all_months)
        
        # 2. Lọc theo Bộ Phận
        all_depts = sorted(df['Department'].unique())
        selected_depts = st.sidebar.multiselect("Chọn Bộ Phận", all_depts, default=all_depts)
        
        # Áp dụng lọc
        mask = (df['Month_Str'].isin(selected_months)) & (df['Department'].isin(selected_depts))
        df_filtered = df[mask]
        
        if df_filtered.empty:
            st.warning("Không có dữ liệu phù hợp với bộ lọc!")
        else:
            # --- PHẦN 1: KPI CARDS ---
            st.markdown("### 1. Tổng Quan Hoạt Động")
            
            # Tính toán chỉ số
            total_trips = len(df_filtered)
            total_km = df_filtered['Km_Used'].sum()
            total_cost = df_filtered['Total_Cost'].sum()
            avg_cost_per_km = (total_cost / total_km) if total_km > 0 else 0
            active_cars = df_filtered['Car_Plate'].nunique()
            
            # Hiển thị 4 cột
            c1, c2, c3, c4 = st.columns(4)
            
            with c1:
                st.markdown(f"""
                <div class="kpi-card">
                    <div class="kpi-title">Tổng Số Chuyến</div>
                    <div class="kpi-value">{total_trips:,}</div>
                    <div class="kpi-unit">Chuyến xe</div>
                </div>
                """, unsafe_allow_html=True)
                
            with c2:
                st.markdown(f"""
                <div class="kpi-card">
                    <div class="kpi-title">Tổng Quãng Đường</div>
                    <div class="kpi-value">{total_km:,.0f}</div>
                    <div class="kpi-unit">Km</div>
                </div>
                """, unsafe_allow_html=True)
                
            with c3:
                st.markdown(f"""
                <div class="kpi-card">
                    <div class="kpi-title">Tổng Chi Phí</div>
                    <div class="kpi-value">{total_cost:,.0f}</div>
                    <div class="kpi-unit">VNĐ</div>
                </div>
                """, unsafe_allow_html=True)
                
            with c4:
                st.markdown(f"""
                <div class="kpi-card">
                    <div class="kpi-title">Chi Phí / KM</div>
                    <div class="kpi-value">{avg_cost_per_km:,.0f}</div>
                    <div class="kpi-unit">VNĐ/Km</div>
                </div>
                """, unsafe_allow_html=True)
            
            st.markdown("---")

            # --- PHẦN 2: BIỂU ĐỒ PHÂN TÍCH ---
            
            # Row 1: Xu hướng & Bộ phận
            col_left, col_right = st.columns(2)
            
            with col_left:
                st.subheader("📈 Xu Hướng Chi Phí & Km Theo Ngày")
                # Group by Date
                daily_stats = df_filtered.groupby('Date')[['Total_Cost', 'Km_Used']].sum().reset_index()
                
                # Vẽ biểu đồ 2 trục (Chi phí và Km)
                fig_trend = go.Figure()
                fig_trend.add_trace(go.Bar(
                    x=daily_stats['Date'], 
                    y=daily_stats['Total_Cost'], 
                    name='Chi Phí (VNĐ)',
                    marker_color='#3498db'
                ))
                fig_trend.add_trace(go.Scatter(
                    x=daily_stats['Date'], 
                    y=daily_stats['Km_Used'], 
                    name='Quãng Đường (Km)',
                    yaxis='y2',
                    line=dict(color='#e74c3c', width=3)
                ))
                
                fig_trend.update_layout(
                    yaxis=dict(title="Chi Phí (VNĐ)"),
                    yaxis2=dict(title="Quãng Đường (Km)", overlaying='y', side='right'),
                    legend=dict(orientation="h", y=1.1),
                    hovermode="x unified"
                )
                st.plotly_chart(fig_trend, use_container_width=True)

            with col_right:
                st.subheader("🏢 Top Bộ Phận Sử Dụng Nhiều Nhất")
                # Group by Dept
                dept_stats = df_filtered.groupby('Department')['Total_Cost'].sum().reset_index()
                dept_stats = dept_stats.sort_values(by='Total_Cost', ascending=True).tail(10) # Top 10
                
                fig_dept = px.bar(
                    dept_stats, 
                    x='Total_Cost', 
                    y='Department', 
                    orientation='h',
                    text_auto='.2s',
                    title="Top 10 Bộ Phận theo Chi Phí",
                    color='Total_Cost',
                    color_continuous_scale='Blues'
                )
                st.plotly_chart(fig_dept, use_container_width=True)

            # Row 2: Xe & Tài xế
            col_car, col_driver = st.columns(2)
            
            with col_car:
                st.subheader("🚗 Hiệu Suất Từng Xe")
                car_stats = df_filtered.groupby('Car_Plate')[['Km_Used', 'Total_Cost']].sum().reset_index()
                fig_car = px.scatter(
                    car_stats,
                    x='Km_Used',
                    y='Total_Cost',
                    size='Total_Cost',
                    color='Car_Plate',
                    hover_name='Car_Plate',
                    title="Tương quan Km & Chi Phí từng xe"
                )
                st.plotly_chart(fig_car, use_container_width=True)
                
            with col_driver:
                st.subheader("👮 Top Tài Xế Chạy Nhiều Nhất (Km)")
                driver_stats = df_filtered.groupby('Driver')['Km_Used'].sum().reset_index().sort_values(by='Km_Used', ascending=False).head(10)
                fig_driver = px.bar(
                    driver_stats,
                    x='Driver',
                    y='Km_Used',
                    color='Km_Used',
                    color_continuous_scale='Greens'
                )
                st.plotly_chart(fig_driver, use_container_width=True)

            # --- PHẦN 3: BẢNG DỮ LIỆU ---
            with st.expander("📄 Xem Dữ Liệu Chi Tiết"):
                st.dataframe(df_filtered.style.format({
                    "Total_Cost": "{:,.0f}",
                    "Km_Used": "{:,.0f}"
                }))
    else:
        st.info("Hãy tải lên file Excel để bắt đầu phân tích.")