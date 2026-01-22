import streamlit as st
import pandas as pd
import matplotlib.pyplot as plt
import seaborn as sns

# --- CẤU HÌNH TRANG ---
st.set_page_config(page_title="Dashboard Đội Xe", page_icon="🚗", layout="wide")

# Tiêu đề chính
st.title("🚗 Dashboard Thống Kê & Quản Lý Đội Xe")
st.markdown("---")

# --- 1. UPLOAD FILE ---
uploaded_file = st.file_uploader("📂 Bước 1: Kéo thả file Excel/CSV dữ liệu vào đây", type=['xlsx', 'csv'])

if uploaded_file is not None:
    # Đọc file
    try:
        if uploaded_file.name.endswith('.csv'):
            df = pd.read_csv(uploaded_file)
        else:
            df = pd.read_excel(uploaded_file, engine='openpyxl')
            
        st.success("✅ Đã tải dữ liệu thành công!")
    except Exception as e:
        st.error(f"❌ Lỗi đọc file: {e}")
        st.stop()

    # --- 2. XỬ LÝ DỮ LIỆU (DATA CLEANING) ---
    # Gộp ngày giờ
    try:
        df['Start_Datetime'] = pd.to_datetime(df['Ngày khởi hành'].astype(str) + ' ' + df['Giờ khởi hành'].astype(str), errors='coerce')
        df['End_Datetime'] = pd.to_datetime(df['Ngày khởi hành'].astype(str) + ' ' + df['Giờ kết thúc'].astype(str), errors='coerce')
        
        # Xử lý qua đêm
        mask_overnight = df['End_Datetime'] < df['Start_Datetime']
        df.loc[mask_overnight, 'End_Datetime'] += pd.Timedelta(days=1)
        
        # Tính thời lượng & cột tháng
        df['Duration_Hours'] = (df['End_Datetime'] - df['Start_Datetime']).dt.total_seconds() / 3600
        df['Month_Year'] = df['Start_Datetime'].dt.to_period('M').astype(str)
        
        # Lọc chỉ lấy các dòng đã gán xe (có biển số)
        df_assigned = df.dropna(subset=['Biển số xe'])

    except Exception as e:
        st.error(f"⚠️ Lỗi cấu trúc dữ liệu: {e}. Vui lòng kiểm tra tên cột Ngày/Giờ khởi hành.")
        st.stop()

    # --- 3. TẠO SIDEBAR BỘ LỌC (FILTER) ---
    st.sidebar.header("🔍 Bộ Lọc Dữ Liệu")
    st.sidebar.info("Chọn điều kiện bên dưới để lọc biểu đồ")

    # A. Lọc theo thời gian
    min_date = df_assigned['Start_Datetime'].min().date()
    max_date = df_assigned['End_Datetime'].max().date()

    date_range = st.sidebar.date_input(
        "📅 Chọn khoảng thời gian:",
        value=(min_date, max_date),
        min_value=min_date,
        max_value=max_date
    )

    # B. Lọc theo Biển số xe
    all_cars = sorted(df_assigned['Biển số xe'].unique().astype(str))
    selected_cars = st.sidebar.multiselect(
        "giao 🚘 Chọn xe hiển thị:",
        options=all_cars,
        default=all_cars
    )

    # --- 4. ÁP DỤNG BỘ LỌC ---
    # Xử lý logic lọc ngày (đề phòng user chỉ chọn 1 ngày)
    if len(date_range) == 2:
        start_date, end_date = date_range
        mask_date = (df_assigned['Start_Datetime'].dt.date >= start_date) & (df_assigned['Start_Datetime'].dt.date <= end_date)
    elif len(date_range) == 1:
        mask_date = (df_assigned['Start_Datetime'].dt.date == date_range[0])
    else:
        mask_date = pd.Series([True] * len(df_assigned)) # Không lọc nếu lỗi

    mask_car = df_assigned['Biển số xe'].isin(selected_cars)
    
    # DATAFRAME SAU KHI LỌC (Dùng cái này để vẽ biểu đồ)
    df_filtered = df_assigned[mask_date & mask_car]

    if df_filtered.empty:
        st.warning("⚠️ Không tìm thấy dữ liệu phù hợp với bộ lọc!")
        st.stop()

    # --- 5. TÍNH TOÁN CHỈ SỐ (KPIs) ---
    
    # Tính Overlap trên dữ liệu đã lọc
    df_sorted = df_filtered.sort_values(by=['Biển số xe', 'Start_Datetime'])
    df_sorted['Prev_End'] = df_sorted.groupby('Biển số xe')['End_Datetime'].shift(1)
    overlaps = df_sorted[df_sorted['Start_Datetime'] < df_sorted['Prev_End']]
    
    num_overlaps = len(overlaps)
    total_bookings = len(df_filtered)
    total_hours = df_filtered['Duration_Hours'].sum()
    overlap_rate = (num_overlaps / total_bookings * 100) if total_bookings > 0 else 0

    # Hiển thị KPI
    col1, col2, col3, col4 = st.columns(4)
    col1.metric("Tổng chuyến đi", f"{total_bookings} chuyến")
    col2.metric("Tổng giờ vận hành", f"{total_hours:,.1f} giờ")
    col3.metric("Số chuyến bị trùng", f"{num_overlaps}", delta_color="inverse")
    col4.metric("Tỷ lệ trùng lặp", f"{overlap_rate:.2f}%", delta_color="inverse")

    st.markdown("---")

    # --- 6. VẼ BIỂU ĐỒ (TABS) ---
    tab1, tab2, tab3 = st.tabs(["📊 Hiệu suất Xe", "👥 Người Dùng & Chi Phí", "⚠️ Danh sách Trùng Lịch"])

    with tab1:
        st.subheader("Thời gian sử dụng xe theo tháng")
        monthly_usage = df_filtered.groupby('Month_Year')['Duration_Hours'].sum().sort_index()
        st.bar_chart(monthly_usage)
        
        st.subheader("Tần suất sử dụng theo Biển số xe")
        car_usage = df_filtered['Biển số xe'].value_counts().head(15)
        st.bar_chart(car_usage)

    with tab2:
        col_left, col_right = st.columns(2)
        with col_left:
            st.subheader("Top 10 Người sử dụng nhiều nhất")
            if 'Người sử dụng xe' in df_filtered.columns:
                user_usage = df_filtered.groupby('Người sử dụng xe')['Duration_Hours'].sum().nlargest(10).sort_values()
                st.bar_chart(user_usage, color="#ffaa00", horizontal=True) # Vẽ ngang cho dễ đọc tên
            else:
                st.info("Không có cột 'Người sử dụng xe'")

        with col_right:
            st.subheader("Chi phí vận hành theo Bộ phận")
            if 'Bộ phận' in df_filtered.columns and 'Tổng chi phí' in df_filtered.columns:
                 # Check xem có dữ liệu chi phí không
                if df_filtered['Tổng chi phí'].sum() > 0:
                    dept_cost = df_filtered.groupby('Bộ phận')['Tổng chi phí'].sum().sort_values(ascending=False)
                    st.bar_chart(dept_cost)
                else:
                    st.info("Dữ liệu 'Tổng chi phí' đang trống hoặc bằng 0.")
            else:
                st.info("File thiếu cột 'Bộ phận' hoặc 'Tổng chi phí'.")

    with tab3:
        st.subheader(f"Chi tiết {num_overlaps} trường hợp bị trùng lịch")
        if num_overlaps > 0:
            st.error("Cảnh báo: Các chuyến xe dưới đây có giờ Khởi hành sớm hơn giờ Kết thúc của chuyến trước đó trên cùng 1 xe.")
            st.dataframe(
                overlaps[['Ngày khởi hành', 'Biển số xe', 'Tên tài xế', 'Start_Datetime', 'End_Datetime', 'Prev_End']]
                .style.format({"Start_Datetime": lambda t: t.strftime("%H:%M"), "End_Datetime": lambda t: t.strftime("%H:%M"), "Prev_End": lambda t: t.strftime("%H:%M")})
            )
        else:
            st.success("Tuyệt vời! Dữ liệu lọc hiện tại không có chuyến nào bị trùng.")

else:
    st.info("👋 Chào bạn! Hãy upload file Excel Booking Car để bắt đầu.")