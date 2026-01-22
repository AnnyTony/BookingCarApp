import streamlit as st
import pandas as pd
import plotly.express as px
import numpy as np

# --- 1. CẤU HÌNH TRANG & CSS (Làm đẹp - Lấy từ bản Pro) ---
st.set_page_config(page_title="Fleet Management Dashboard", page_icon="🚘", layout="wide")

# CSS: Tùy chỉnh màu Sidebar, Metric, Header
st.markdown("""
<style>
    /* Chỉnh màu nền Sidebar */
    [data-testid="stSidebar"] {
        background-color: #f0f2f6;
    }
    /* Chỉnh Tiêu đề Sidebar */
    [data-testid="stSidebar"] h1 {
        font-size: 20px;
        color: #1f77b4;
    }
    /* Chỉnh các thẻ chỉ số (KPI Card) */
    div[data-testid="stMetricValue"] {
        font-size: 24px;
        color: #007bff;
        font-weight: bold;
    }
    /* Tiêu đề chính đẹp hơn */
    .main-header {
        font-family: 'Helvetica Neue', sans-serif;
        color: #2c3e50;
        font-size: 32px;
        font-weight: 700;
    }
    .sub-header {
        font-size: 16px; 
        color: #7f8c8d;
    }
</style>
""", unsafe_allow_html=True)

# --- 2. HEADER ---
col_head1, col_head2 = st.columns([4, 1])
with col_head1:
    st.markdown("<div class='main-header'>🚘 Fleet Operations Center</div>", unsafe_allow_html=True)
    st.markdown("<div class='sub-header'>Hệ thống báo cáo thông minh & Tự động hóa tính toán</div>", unsafe_allow_html=True)
with col_head2:
    st.image("https://cdn-icons-png.flaticon.com/512/3097/3097180.png", width=70)

st.divider()

# --- 3. HÀM XỬ LÝ DỮ LIỆU ---
@st.cache_data
def load_and_process_data(file):
    try:
        if file.name.endswith('.csv'):
            df = pd.read_csv(file)
        else:
            df = pd.read_excel(file, engine='openpyxl')
        
        # Chuẩn hóa tên cột
        df.columns = df.columns.str.strip()
        
        # Xử lý Ngày Giờ (Cố gắng ép kiểu, nếu lỗi thì bỏ qua)
        try:
            df['Start_Datetime'] = pd.to_datetime(df['Ngày khởi hành'].astype(str) + ' ' + df['Giờ khởi hành'].astype(str), errors='coerce')
            df['End_Datetime'] = pd.to_datetime(df['Ngày khởi hành'].astype(str) + ' ' + df['Giờ kết thúc'].astype(str), errors='coerce')
            
            mask_overnight = df['End_Datetime'] < df['Start_Datetime']
            df.loc[mask_overnight, 'End_Datetime'] += pd.Timedelta(days=1)
            
            df['Thời lượng (Giờ)'] = (df['End_Datetime'] - df['Start_Datetime']).dt.total_seconds() / 3600
            df['Tháng'] = df['Start_Datetime'].dt.to_period('M').astype(str)
            df['Day_Name'] = df['Start_Datetime'].dt.day_name()
        except:
            pass 
            
        return df
    except Exception as e:
        return str(e)

# --- 4. UPLOAD DATA ---
uploaded_file = st.file_uploader("📂 Import Dữ liệu Booking (Kéo thả file vào đây)", type=['xlsx', 'csv'])

if uploaded_file is None:
    st.info("👈 Vui lòng tải file dữ liệu lên để bắt đầu.")
    st.stop()

df = load_and_process_data(uploaded_file)
if isinstance(df, str): 
    st.error(f"Lỗi dữ liệu: {df}")
    st.stop()

# --- 5. SIDEBAR XỊN (Lấy lại từ bản Pro) ---
with st.sidebar:
    st.markdown("## 🎛️ Bảng Điều Khiển")
    
    # Gom nhóm 1: Thời gian
    with st.expander("📆 Lọc Thời Gian", expanded=True):
        if 'Start_Datetime' in df.columns:
            df_valid = df.dropna(subset=['Start_Datetime'])
            if not df_valid.empty:
                min_d = df_valid['Start_Datetime'].min().date()
                max_d = df_valid['End_Datetime'].max().date()
                
                date_range = st.date_input("Chọn khoảng ngày:", value=(min_d, max_d), min_value=min_d, max_value=max_d)
    
    # Gom nhóm 2: Xe (Có nút Select All xịn xò)
    with st.expander("🚗 Lọc Theo Xe", expanded=False): # Mặc định đóng cho gọn
        if 'Biển số xe' in df.columns:
            all_cars = sorted(df['Biển số xe'].dropna().astype(str).unique())
            
            select_all_cars = st.toggle("Chọn tất cả xe", value=True)
            if select_all_cars:
                selected_cars = all_cars
            else:
                selected_cars = st.multiselect("Chọn xe cụ thể:", options=all_cars, default=all_cars[:5])
        else:
            selected_cars = []

    # Nút Reset
    if st.button("🔄 Reset Bộ Lọc", type="primary", use_container_width=True):
        st.rerun()
    
    st.markdown("---")
    st.caption(f"Dữ liệu gốc: {len(df)} dòng")

# --- XỬ LÝ LOGIC LỌC ---
df_filtered = df.copy()

# 1. Lọc ngày
if 'Start_Datetime' in df.columns and isinstance(date_range, tuple) and len(date_range) == 2:
    mask_date = (df_filtered['Start_Datetime'].dt.date >= date_range[0]) & (df_filtered['Start_Datetime'].dt.date <= date_range[1])
    df_filtered = df_filtered[mask_date]

# 2. Lọc xe
if 'Biển số xe' in df.columns and selected_cars:
    df_filtered = df_filtered[df_filtered['Biển số xe'].astype(str).isin(selected_cars)]

st.sidebar.success(f"🔍 Hiển thị: **{len(df_filtered)}** chuyến")

# --- 6. DASHBOARD CHÍNH ---

# TABS
tab1, tab2, tab3, tab4 = st.tabs(["📊 Tổng Quan Hiệu Suất", "🏢 Đơn Vị & User", "⚠️ Kiểm Tra Trùng", "🧮 Máy Tính Thông Minh"])

# --- TAB 1: TỔNG QUAN (Giao diện Pro) ---
with tab1:
    if 'Thời lượng (Giờ)' in df_filtered.columns:
        total_trips = len(df_filtered)
        total_hours = df_filtered['Thời lượng (Giờ)'].sum()
        avg_duration = df_filtered['Thời lượng (Giờ)'].mean()
        
        # 3 Metrics đẹp
        c1, c2, c3 = st.columns(3)
        c1.metric("Tổng Số Chuyến", f"{total_trips}")
        c2.metric("Tổng Giờ Vận Hành", f"{total_hours:,.0f}h")
        c3.metric("TB Một Chuyến", f"{avg_duration:.1f}h")
        
        st.markdown("---")
        
        # Biểu đồ cột
        col_chart1, col_chart2 = st.columns([2, 1])
        with col_chart1:
            daily_usage = df_filtered.groupby('Tháng')['Thời lượng (Giờ)'].sum().reset_index()
            fig = px.bar(daily_usage, x='Tháng', y='Thời lượng (Giờ)', 
                         title="Tổng giờ hoạt động theo Tháng",
                         text_auto='.0f', color='Thời lượng (Giờ)', color_continuous_scale='Blues')
            st.plotly_chart(fig, use_container_width=True)
            
        with col_chart2:
             if 'Biển số xe' in df_filtered.columns:
                car_counts = df_filtered['Biển số xe'].value_counts().reset_index().head(8)
                car_counts.columns = ['Xe', 'Số chuyến']
                fig_pie = px.pie(car_counts, values='Số chuyến', names='Xe', title="Top Xe hoạt động", hole=0.5)
                fig_pie.update_layout(showlegend=False)
                st.plotly_chart(fig_pie, use_container_width=True)
    else:
        st.warning("Dữ liệu thiếu cột ngày giờ, không vẽ được biểu đồ tổng quan.")

# --- TAB 2: ĐƠN VỊ ---
with tab2:
    # Tự động tìm cột
    cols_to_plot = [c for c in df_filtered.columns if c in ['Bộ phận', 'Công ty', 'Cost center', 'Người sử dụng xe']]
    
    if cols_to_plot:
        selected_col = st.selectbox("Chọn tiêu chí thống kê:", cols_to_plot)
        # Fillna
        df_plot = df_filtered.copy()
        df_plot[selected_col] = df_plot[selected_col].fillna("Unknown")
        
        counts = df_plot[selected_col].value_counts().reset_index().head(15)
        counts.columns = [selected_col, 'Số chuyến']
        
        fig2 = px.bar(counts, x='Số chuyến', y=selected_col, orientation='h', 
                      title=f"Top 15 {selected_col} có lượt đặt nhiều nhất",
                      text_auto=True, color='Số chuyến', color_continuous_scale='Sunset')
        fig2.update_layout(yaxis={'categoryorder':'total ascending'})
        st.plotly_chart(fig2, use_container_width=True)
    else:
        st.info("Không tìm thấy các cột thông tin đơn vị (Bộ phận/Công ty...). Hãy kiểm tra tên cột trong file Excel.")

# --- TAB 3: CHECK TRÙNG ---
with tab3:
    if 'Biển số xe' in df_filtered.columns and 'Start_Datetime' in df_filtered.columns:
        df_s = df_filtered.dropna(subset=['Biển số xe']).sort_values(['Biển số xe', 'Start_Datetime'])
        df_s['Prev_End'] = df_s.groupby('Biển số xe')['End_Datetime'].shift(1)
        overlaps = df_s[df_s['Start_Datetime'] < df_s['Prev_End']]
        
        if not overlaps.empty:
            st.error(f"⚠️ CẢNH BÁO: Phát hiện {len(overlaps)} trường hợp trùng lịch xe!")
            
            # Format String để tránh lỗi JSON NaN
            display_cols = ['Ngày khởi hành', 'Biển số xe', 'Tên tài xế', 'Start_Datetime', 'End_Datetime', 'Prev_End']
            df_display = overlaps[display_cols].copy()
            for col in ['Start_Datetime', 'End_Datetime', 'Prev_End']:
                df_display[col] = df_display[col].dt.strftime('%Y-%m-%d %H:%M')
            
            st.dataframe(df_display, use_container_width=True)
        else:
            st.success("✅ Tuyệt vời! Không có chuyến xe nào bị trùng giờ trong dữ liệu lọc.")

# --- TAB 4: MÁY TÍNH THÔNG MINH (Giữ nguyên logic sửa lỗi NaN) ---
with tab4:
    st.markdown("### 🛠️ Công cụ Tự Tạo Công Thức (AI Calculator)")
    st.info("💡 Chọn 2 cột số bất kỳ để thực hiện phép tính. Hệ thống sẽ tự động xử lý lỗi chia cho 0.")
    
    numeric_cols = df_filtered.select_dtypes(include=[np.number]).columns.tolist()
    
    if len(numeric_cols) < 2:
        st.warning("⚠️ File không đủ cột dữ liệu số để tính toán.")
    else:
        c1, c2, c3, c4 = st.columns([3, 1, 3, 2])
        
        with c1:
            col_a = st.selectbox("Cột A:", numeric_cols, index=0)
        with c2:
            operator = st.selectbox("Phép tính:", ["+", "-", "*", "/"])
        with c3:
            input_mode = st.radio("Cột B là:", ["Một Cột Khác", "Số Cố Định"], horizontal=True)
            if input_mode == "Một Cột Khác":
                col_b = st.selectbox("Cột B:", numeric_cols, index=1 if len(numeric_cols)>1 else 0)
                val_b = None
            else:
                col_b = None
                val_b = st.number_input("Nhập số:", value=1.0)
        
        with c4:
            st.write("") 
            st.write("")
            calc_btn = st.button("🚀 Tính Ngay", type="primary", use_container_width=True)

        if calc_btn:
            new_col_name = f"Kết quả ({col_a} {operator} {col_b if col_b else val_b})"
            try:
                # Tính toán
                series_a = pd.to_numeric(df_filtered[col_a], errors='coerce').fillna(0)
                if col_b:
                    series_b = pd.to_numeric(df_filtered[col_b], errors='coerce').fillna(0)
                else:
                    series_b = val_b

                if operator == "+": res = series_a + series_b
                elif operator == "-": res = series_a - series_b
                elif operator == "*": res = series_a * series_b
                elif operator == "/": res = series_a / series_b.replace(0, np.nan)
                
                # --- FIX LỖI NaN/Inf ---
                res = res.replace([np.inf, -np.inf], 0)
                res = res.fillna(0)
                
                df_filtered[new_col_name] = res
                
                st.success(f"✅ Đã tạo cột mới: **{new_col_name}**")
                
                # Thống kê nhanh
                m1, m2 = st.columns(2)
                m1.metric("Tổng cộng", f"{res.sum():,.2f}")
                m2.metric("Trung bình", f"{res.mean():,.2f}")
                
                # Vẽ biểu đồ kết quả
                st.markdown("#### 📊 Biểu đồ phân bố kết quả")
                x_axis_options = [c for c in df_filtered.columns if df_filtered[c].dtype == 'object'] 
                if not x_axis_options: x_axis_options = ['index']
                
                x_axis = st.selectbox("Gom nhóm theo:", x_axis_options, index=0)
                
                chart_data = df_filtered.groupby(x_axis)[new_col_name].sum().reset_index()
                fig_calc = px.bar(chart_data, x=x_axis, y=new_col_name, 
                                  title=f"Biểu đồ {new_col_name} theo {x_axis}",
                                  color=new_col_name, color_continuous_scale='Viridis')
                st.plotly_chart(fig_calc, use_container_width=True)

            except Exception as e:
                st.error(f"Lỗi tính toán: {e}")