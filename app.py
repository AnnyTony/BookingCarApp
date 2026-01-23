import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go

# --- 1. CẤU HÌNH TRANG ---
st.set_page_config(page_title="Fleet Management Dashboard Pro", page_icon="🚗", layout="wide")

# CSS cho giao diện đẹp như Power BI
st.markdown("""
<style>
    .main-header {font-size: 26px; font-weight: bold; color: #2c3e50; margin-bottom: 20px;}
    .kpi-card {
        background-color: white; 
        padding: 20px; 
        border-radius: 10px; 
        border-left: 5px solid #3498db; 
        box-shadow: 2px 2px 10px rgba(0,0,0,0.05);
        text-align: center;
    }
    .kpi-value {font-size: 28px; font-weight: bold; color: #2c3e50;}
    .kpi-label {font-size: 14px; color: #7f8c8d; text-transform: uppercase;}
    [data-testid="stSidebar"] {background-color: #f8f9fa;}
</style>
""", unsafe_allow_html=True)

st.markdown("<div class='main-header'>🚗 Fleet Management Intelligence (Full Option)</div>", unsafe_allow_html=True)

# --- 2. HÀM LOAD & XỬ LÝ DỮ LIỆU (ĐÃ FIX LỖI DUPLICATE) ---
@st.cache_data
def load_data():
    try:
        # A. ĐỌC DỮ LIỆU
        # Driver: Tìm header đúng (thường ở dòng thứ 3 - index 2)
        df_driver_raw = pd.read_csv("Booking car.xlsx - Driver.csv", header=None)
        # Tìm dòng chứa chữ 'Biển số xe' để làm header
        try:
            header_idx = df_driver_raw[df_driver_raw.eq("Biển số xe").any(axis=1)].index[0]
        except IndexError:
            header_idx = 2 # Fallback nếu không tìm thấy
            
        df_driver = pd.read_csv("Booking car.xlsx - Driver.csv", header=header_idx)
        df_cbnv = pd.read_csv("Booking car.xlsx - CBNV.csv", header=1)
        df_booking = pd.read_csv("Booking car.xlsx - Booking car.csv", header=0)

        # B. LÀM SẠCH & KHỬ TRÙNG LẶP (FIX LỖI CANNOT REINDEX)
        
        # --- Xử lý Driver ---
        # Chuẩn hóa tên cột (xóa xuống dòng, khoảng trắng thừa)
        df_driver.columns = df_driver.columns.str.replace('\n', ' ').str.strip()
        if 'Cost center' in df_driver.columns: 
            df_driver.rename(columns={'Cost center': 'Cost Center Driver'}, inplace=True)
            
        # QUAN TRỌNG: Loại bỏ xe trùng lặp. Giữ dòng cuối cùng (thường là cập nhật mới nhất)
        df_driver = df_driver.drop_duplicates(subset=['Biển số xe'], keep='last')
        
        # --- Xử lý CBNV ---
        # QUAN TRỌNG: Loại bỏ nhân viên trùng tên.
        df_cbnv = df_cbnv.drop_duplicates(subset=['Full Name'], keep='first')

        # C. MERGE DỮ LIỆU
        # Merge Booking với Driver
        df_final = df_booking.merge(df_driver[['Biển số xe', 'Loại nhiên liệu', 'Tên tài xế']], on='Biển số xe', how='left', suffixes=('', '_Driver'))
        
        # Merge Booking với CBNV (User -> Full Name)
        df_final = df_final.merge(df_cbnv[['Full Name', 'Location', 'Công ty', 'BU', 'Position EN']], left_on='Người sử dụng xe', right_on='Full Name', how='left')

        # D. XỬ LÝ THỜI GIAN & PHÂN LOẠI
        df_final['Ngày khởi hành'] = pd.to_datetime(df_final['Ngày khởi hành'], errors='coerce')
        df_final['Tháng'] = df_final['Ngày khởi hành'].dt.strftime('%Y-%m')
        
        # Xử lý dữ liệu thiếu cho biểu đồ Sunburst (không được để trống)
        df_final['Location'] = df_final['Location'].fillna('Unknown')
        df_final['Công ty'] = df_final['Công ty'].fillna('Other')
        df_final['BU'] = df_final['BU'].fillna('Other')

        # Tạo cột phân loại "Nội thành/Tỉnh" (Ví dụ logic đơn giản dựa trên lộ trình)
        def phan_loai_chuyen(lo_trinh):
            if pd.isna(lo_trinh): return "Khác"
            if "Tỉnh" in str(lo_trinh) or "TP." in str(lo_trinh) and "Hồ Chí Minh" not in str(lo_trinh):
                return "Đi Tỉnh"
            return "Nội Thành"
        
        # Nếu chưa có cột phân loại, tạo tạm để vẽ biểu đồ tròn
        if 'Phạm Vi' not in df_final.columns:
            df_final['Phạm Vi'] = df_final['Lộ trình'].apply(phan_loai_chuyen)

        return df_final

    except Exception as e:
        st.error(f"Lỗi xử lý dữ liệu chi tiết: {e}")
        return pd.DataFrame()

# Load data
df = load_data()

if not df.empty:
    # --- 3. BỘ LỌC PHÂN CẤP (SIDEBAR) ---
    st.sidebar.header("🔍 Bộ Lọc Phân Cấp (Drill-down)")

    # Level 1: Location
    all_locations = sorted(df['Location'].unique())
    selected_location = st.sidebar.multiselect("1. Chọn Khu Vực", all_locations, default=all_locations)
    df_lvl1 = df[df['Location'].isin(selected_location)]

    # Level 2: Công ty
    available_companies = sorted(df_lvl1['Công ty'].unique())
    selected_company = st.sidebar.multiselect("2. Chọn Công Ty", available_companies, default=available_companies)
    df_lvl2 = df_lvl1[df_lvl1['Công ty'].isin(selected_company)]

    # Level 3: BU
    available_bus = sorted(df_lvl2['BU'].unique())
    selected_bu = st.sidebar.multiselect("3. Chọn Bộ Phận (BU)", available_bus, default=available_bus)
    df_filtered = df_lvl2[df_lvl2['BU'].isin(selected_bu)]
    
    # --- 4. KPI SUMMARY (Giống Power BI Cards) ---
    col1, col2, col3, col4 = st.columns(4)
    
    total_trips = len(df_filtered)
    top_user = df_filtered['Người sử dụng xe'].mode()[0] if total_trips > 0 else "N/A"
    active_cars = df_filtered['Biển số xe'].nunique()
    # Giả lập tính tổng giờ (nếu có cột duration), ở đây đếm số chuyến đi tỉnh
    trips_province = len(df_filtered[df_filtered['Phạm Vi'] == 'Đi Tỉnh'])

    with col1: st.markdown(f"<div class='kpi-card'><div class='kpi-value'>{total_trips}</div><div class='kpi-label'>Tổng Chuyến Đi</div></div>", unsafe_allow_html=True)
    with col2: st.markdown(f"<div class='kpi-card'><div class='kpi-value' style='font-size:20px'>{top_user}</div><div class='kpi-label'>Top User</div></div>", unsafe_allow_html=True)
    with col3: st.markdown(f"<div class='kpi-card'><div class='kpi-value'>{active_cars}</div><div class='kpi-label'>Số Xe Hoạt Động</div></div>", unsafe_allow_html=True)
    with col4: st.markdown(f"<div class='kpi-card'><div class='kpi-value'>{trips_province}</div><div class='kpi-label'>Chuyến Đi Tỉnh</div></div>", unsafe_allow_html=True)

    st.markdown("---")

    # --- 5. TABS CHỨC NĂNG ---
    tab_drill, tab_overview, tab_data = st.tabs(["📊 Drill-down Phân Cấp", "📈 Biểu Đồ Tổng Quan", "dữ liệu chi tiết"])

    # TAB 1: DRILL-DOWN (Mới)
    with tab_drill:
        col_sun, col_tree = st.columns(2)
        
        with col_sun:
            st.subheader("Cấu Trúc: Vùng → Công ty → BU")
            if not df_filtered.empty:
                fig_sun = px.sunburst(
                    df_filtered, 
                    path=['Location', 'Công ty', 'BU'], 
                    title="Tỷ trọng Chuyến đi theo Cấu trúc (Click để zoom)",
                    height=500
                )
                st.plotly_chart(fig_sun, use_container_width=True)
            else:
                st.warning("Không có dữ liệu cho bộ lọc này")

        with col_tree:
            st.subheader("Treemap: Phân bổ theo Công ty")
            if not df_filtered.empty:
                # Group data cho Treemap
                df_tree = df_filtered.groupby(['Location', 'Công ty', 'BU']).size().reset_index(name='Số chuyến')
                fig_tree = px.treemap(
                    df_tree,
                    path=['Location', 'Công ty', 'BU'],
                    values='Số chuyến',
                    color='Số chuyến',
                    color_continuous_scale='RdBu',
                    title="Diện tích thể hiện số lượng chuyến đi",
                    height=500
                )
                st.plotly_chart(fig_tree, use_container_width=True)

    # TAB 2: TỔNG QUAN (Các biểu đồ cũ + Biểu đồ xu hướng)
    with tab_overview:
        col_trend, col_pie = st.columns([2, 1])
        
        with col_trend:
            st.subheader("Xu hướng đặt xe theo thời gian")
            if 'Tháng' in df_filtered.columns and not df_filtered.empty:
                df_trend = df_filtered.groupby('Tháng').size().reset_index(name='Số chuyến')
                fig_line = px.area(df_trend, x='Tháng', y='Số chuyến', markers=True, 
                                   title="Số lượng chuyến đi theo tháng", color_discrete_sequence=['#3498db'])
                st.plotly_chart(fig_line, use_container_width=True)
        
        with col_pie:
            st.subheader("Tỷ lệ Nội thành vs Đi Tỉnh")
            if 'Phạm Vi' in df_filtered.columns and not df_filtered.empty:
                df_pie = df_filtered['Phạm Vi'].value_counts().reset_index()
                df_pie.columns = ['Phạm Vi', 'Số lượng']
                fig_donut = px.pie(df_pie, values='Số lượng', names='Phạm Vi', hole=0.5, 
                                   title="Cơ cấu lộ trình", color_discrete_sequence=px.colors.qualitative.Pastel)
                st.plotly_chart(fig_donut, use_container_width=True)

        st.subheader("🏆 Top 10 Xe & Tài xế hoạt động tích cực")
        if not df_filtered.empty:
            top_drivers = df_filtered.groupby(['Biển số xe', 'Tên tài xế']).size().reset_index(name='Số chuyến')
            top_drivers = top_drivers.sort_values('Số chuyến', ascending=False).head(10)
            
            fig_bar = px.bar(top_drivers, x='Số chuyến', y='Tên tài xế', orientation='h', 
                             text='Số chuyến', color='Số chuyến', title="Top Tài xế (theo số chuyến)",
                             hover_data=['Biển số xe'])
            fig_bar.update_layout(yaxis={'categoryorder':'total ascending'})
            st.plotly_chart(fig_bar, use_container_width=True)

    # TAB 3: DATA
    with tab_data:
        st.dataframe(df_filtered)

else:
    st.error("Không thể tải dữ liệu. Vui lòng kiểm tra file Excel (Sheet tên có đúng không?).")