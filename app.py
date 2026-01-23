import streamlit as st
import pandas as pd
import plotly.express as px
import numpy as np

# --- 1. CẤU HÌNH TRANG ---
st.set_page_config(page_title="Fleet Management Dashboard", page_icon="🚗", layout="wide")

# CSS giao diện chuẩn Power BI
st.markdown("""
<style>
    .main-header {font-size: 24px; font-weight: bold; color: #2c3e50; margin-bottom: 20px;}
    .kpi-card {background-color: white; padding: 15px; border-radius: 8px; border-left: 5px solid #007bff; box-shadow: 0 2px 4px rgba(0,0,0,0.1);}
    .kpi-value {font-size: 24px; font-weight: bold; color: #007bff;}
    .kpi-label {font-size: 14px; color: #6c757d;}
    [data-testid="stSidebar"] {background-color: #f8f9fa;}
</style>
""", unsafe_allow_html=True)

st.markdown("<div class='main-header'>🚗 Fleet Management Intelligence (Drill-down Edition)</div>", unsafe_allow_html=True)

# --- 2. HÀM LOAD & XỬ LÝ DỮ LIỆU ---
@st.cache_data
def load_data():
    try:
        # A. ĐỌC DỮ LIỆU
        # Tự động tìm header cho Driver
        df_driver_raw = pd.read_csv("Booking car.xlsx - Driver.csv", header=None)
        # Tìm dòng chứa chữ "Biển số xe" để làm header
        header_idx = df_driver_raw[df_driver_raw.eq("Biển số xe").any(axis=1)].index[0]
        df_driver = pd.read_csv("Booking car.xlsx - Driver.csv", header=header_idx)
        
        df_cbnv = pd.read_csv("Booking car.xlsx - CBNV.csv", header=1)
        df_booking = pd.read_csv("Booking car.xlsx - Booking car.csv")

        # B. LÀM SẠCH (FIX LỖI DUPLICATE LABEL)
        # 1. Driver
        cols_driver = ['Biển số xe', 'Loại nhiên liệu', 'Cost \ncenter', 'Tên tài xế']
        cols_driver = [c for c in cols_driver if c in df_driver.columns]
        df_driver = df_driver[cols_driver].dropna(subset=['Biển số xe']).drop_duplicates(subset=['Biển số xe'], keep='last')
        if 'Cost \ncenter' in df_driver.columns:
            df_driver.rename(columns={'Cost \ncenter': 'Cost Center'}, inplace=True)

        # 2. CBNV
        cols_cbnv = ['Full Name', 'Location', 'Công ty', 'BU', 'Position EN']
        cols_cbnv = [c for c in cols_cbnv if c in df_cbnv.columns]
        df_cbnv = df_cbnv[cols_cbnv].dropna(subset=['Full Name']).drop_duplicates(subset=['Full Name'], keep='first')

        # C. MERGE DATA
        df_final = df_booking.merge(df_driver, on='Biển số xe', how='left')
        df_final = df_final.merge(df_cbnv, left_on='Người sử dụng xe', right_on='Full Name', how='left')

        # D. TÍNH TOÁN CỘT MỚI (PHỤC HỒI TÍNH NĂNG CŨ)
        # 1. Xử lý ngày tháng
        df_final['Ngày khởi hành'] = pd.to_datetime(df_final['Ngày khởi hành'], errors='coerce')
        df_final['Tháng'] = df_final['Ngày khởi hành'].dt.strftime('%Y-%m')
        
        # 2. Tính thời gian chạy (Duration)
        # Giả sử format là HH:MM:SS, cần convert sang timedelta
        for col in ['Giờ khởi hành', 'Giờ kết thúc']:
            df_final[col] = pd.to_datetime(df_final[col], format='%H:%M:%S', errors='coerce').dt.time
            
        # Hàm tính giờ đơn giản (nếu lỗi thì trả về 0)
        def calc_hours(row):
            try:
                t1 = pd.to_timedelta(str(row['Giờ khởi hành']))
                t2 = pd.to_timedelta(str(row['Giờ kết thúc']))
                return (t2 - t1).total_seconds() / 3600
            except:
                return 0
        
        df_final['Số giờ'] = df_final.apply(calc_hours, axis=1)
        df_final['Số giờ'] = df_final['Số giờ'].apply(lambda x: x if x > 0 else 0) # Lọc số âm

        # 3. Phân loại Lộ trình (Tạo cột 'Phạm Vi' cho biểu đồ Donut)
        # Logic: Nếu lộ trình chứa tên tỉnh khác -> Đi tỉnh, ngược lại -> Nội thành
        def classify_route(route):
            route = str(route).lower()
            if 'tỉnh' in route or 'tp.' in route and ('hcm' not in route and 'hà nội' not in route):
                return 'Đi Tỉnh'
            return 'Nội Thành'
        
        df_final['Phạm Vi'] = df_final['Lộ trình'].apply(classify_route)

        # Điền dữ liệu trống để vẽ Sunburst không lỗi
        df_final['Location'] = df_final['Location'].fillna('Unknown')
        df_final['Công ty'] = df_final['Công ty'].fillna('Other')
        df_final['BU'] = df_final['BU'].fillna('Other')

        return df_final

    except Exception as e:
        st.error(f"Có lỗi khi xử lý dữ liệu: {e}")
        return pd.DataFrame()

df = load_data()

if not df.empty:
    # --- 3. BỘ LỌC PHÂN CẤP (SIDEBAR) ---
    st.sidebar.header("🔍 Bộ Lọc Drill-down")
    
    # Level 1
    locs = sorted(df['Location'].unique())
    sel_loc = st.sidebar.multiselect("1. Khu Vực", locs, default=locs)
    df_1 = df[df['Location'].isin(sel_loc)]
    
    # Level 2
    comps = sorted(df_1['Công ty'].unique())
    sel_comp = st.sidebar.multiselect("2. Công Ty", comps, default=comps)
    df_2 = df_1[df_1['Công ty'].isin(sel_comp)]
    
    # Level 3
    bus = sorted(df_2['BU'].unique())
    sel_bu = st.sidebar.multiselect("3. Bộ Phận (BU)", bus, default=bus)
    df_filtered = df_2[df_2['BU'].isin(sel_bu)]

    # --- 4. KPI SUMMARY ---
    col1, col2, col3, col4 = st.columns(4)
    with col1:
        st.markdown(f"<div class='kpi-card'><div class='kpi-value'>{len(df_filtered):,}</div><div class='kpi-label'>Tổng Chuyến Đi</div></div>", unsafe_allow_html=True)
    with col2:
        total_hours = df_filtered['Số giờ'].sum()
        st.markdown(f"<div class='kpi-card'><div class='kpi-value'>{total_hours:,.0f}h</div><div class='kpi-label'>Tổng Giờ Vận Hành</div></div>", unsafe_allow_html=True)
    with col3:
        top_driver = df_filtered['Tên tài xế'].mode()[0] if not df_filtered.empty else "-"
        st.markdown(f"<div class='kpi-card'><div class='kpi-value' style='font-size:18px'>{top_driver}</div><div class='kpi-label'>Tài Xế Chạy Nhiều Nhất</div></div>", unsafe_allow_html=True)
    with col4:
        avg_trip = len(df_filtered) / df_filtered['Biển số xe'].nunique() if not df_filtered.empty else 0
        st.markdown(f"<div class='kpi-card'><div class='kpi-value'>{avg_trip:.1f}</div><div class='kpi-label'>Trung bình chuyến/xe</div></div>", unsafe_allow_html=True)

    st.markdown("---")

    # --- 5. VISUALIZATION TABS ---
    tab1, tab2 = st.tabs(["📊 Cấu Trúc Tổ Chức (Drill-down)", "📈 Hiệu Suất & Xu Hướng"])

    # TAB 1: SUNBURST & TREEMAP (YÊU CẦU CỦA BẠN)
    with tab1:
        c1, c2 = st.columns(2)
        with c1:
            st.subheader("Phân rã chi phí theo cấu trúc")
            if not df_filtered.empty:
                fig_sun = px.sunburst(
                    df_filtered, 
                    path=['Location', 'Công ty', 'BU'], 
                    title="Cấu trúc: Vùng -> Công ty -> BU",
                    height=500
                )
                st.plotly_chart(fig_sun, use_container_width=True)
        
        with c2:
            st.subheader("Tỷ trọng theo Bộ phận")
            if not df_filtered.empty:
                df_tree = df_filtered.groupby(['Location', 'Công ty', 'BU']).size().reset_index(name='Count')
                fig_tree = px.treemap(
                    df_tree, 
                    path=['Location', 'Công ty', 'BU'], 
                    values='Count',
                    color='Count',
                    color_continuous_scale='RdBu',
                    title="Diện tích thể hiện số lượng chuyến đi"
                )
                st.plotly_chart(fig_tree, use_container_width=True)

    # TAB 2: CÁC BIỂU ĐỒ CŨ (KHÔI PHỤC)
    with tab2:
        c3, c4 = st.columns([2, 1])
        with c3:
            st.subheader("Xu hướng sử dụng xe theo tháng")
            if not df_filtered.empty:
                # Group by Month và tính tổng số giờ hoặc số chuyến
                df_trend = df_filtered.groupby('Tháng').agg({'SPid': 'count', 'Số giờ': 'sum'}).reset_index()
                # Vẽ 2 đường: Số chuyến và Số giờ
                fig_line = px.line(df_trend, x='Tháng', y='SPid', markers=True, title="Số lượng chuyến đi")
                fig_line.add_bar(x=df_trend['Tháng'], y=df_trend['Số giờ'], name="Tổng giờ", opacity=0.3)
                st.plotly_chart(fig_line, use_container_width=True)
        
        with c4:
            st.subheader("Tỷ lệ Nội thành vs Đi Tỉnh")
            if 'Phạm Vi' in df_filtered.columns and not df_filtered.empty:
                df_pie = df_filtered['Phạm Vi'].value_counts().reset_index()
                df_pie.columns = ['Loại', 'Số lượng']
                fig_donut = px.pie(df_pie, values='Số lượng', names='Loại', hole=0.5, color_discrete_sequence=px.colors.sequential.RdBu)
                st.plotly_chart(fig_donut, use_container_width=True)

        st.subheader("Top 10 Xe hoạt động hiệu quả nhất")
        if not df_filtered.empty:
            top_cars = df_filtered.groupby('Biển số xe').agg({'Số giờ': 'sum', 'SPid': 'count'}).reset_index()
            top_cars = top_cars.sort_values(by='Số giờ', ascending=False).head(10)
            fig_bar = px.bar(top_cars, x='Số giờ', y='Biển số xe', orientation='h', 
                             text='Số giờ', color='SPid', labels={'SPid': 'Số chuyến'},
                             title="Xếp hạng theo tổng giờ vận hành")
            fig_bar.update_layout(yaxis={'categoryorder':'total ascending'})
            st.plotly_chart(fig_bar, use_container_width=True)

else:
    st.info("Đang chờ dữ liệu... Vui lòng kiểm tra file Excel.")