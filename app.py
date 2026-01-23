import streamlit as st
import pandas as pd
import plotly.express as px

# --- 1. CẤU HÌNH TRANG ---
st.set_page_config(page_title="Fleet Management Dashboard Pro", page_icon="🚗", layout="wide")

# CSS cho giao diện
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
</style>
""", unsafe_allow_html=True)

st.markdown("<div class='main-header'>🚗 Fleet Management Intelligence (Upload Edition)</div>", unsafe_allow_html=True)

# --- 2. SIDEBAR & UPLOAD ---
st.sidebar.header("📂 Dữ Liệu Đầu Vào")
uploaded_file = st.sidebar.file_uploader("Tải lên file 'Booking car.xlsx'", type=["xlsx"])

# --- 3. HÀM XỬ LÝ DỮ LIỆU ---
@st.cache_data
def process_data(file):
    try:
        xls = pd.ExcelFile(file)
        
        # A. ĐỌC DỮ LIỆU TỪ CÁC SHEET
        # 1. Driver (Tìm header 'Biển số xe')
        # Đọc thử sheet Driver
        df_driver_raw = pd.read_excel(xls, sheet_name='Driver', header=None)
        # Tìm dòng chứa header thật
        try:
            header_idx = df_driver_raw[df_driver_raw.eq("Biển số xe").any(axis=1)].index[0]
        except:
            header_idx = 2 # Mặc định
        df_driver = pd.read_excel(xls, sheet_name='Driver', header=header_idx)
        
        # 2. CBNV & Booking (Header cố định)
        df_cbnv = pd.read_excel(xls, sheet_name='CBNV', header=1)
        df_booking = pd.read_excel(xls, sheet_name='Booking car', header=0)

        # B. LÀM SẠCH (Fix lỗi Duplicate Labels)
        
        # --- Driver ---
        df_driver.columns = df_driver.columns.str.replace('\n', ' ').str.strip()
        if 'Cost center' in df_driver.columns: 
            df_driver.rename(columns={'Cost center': 'Cost Center Driver'}, inplace=True)
        # Loại bỏ xe trùng, giữ dòng cuối
        if 'Biển số xe' in df_driver.columns:
            df_driver = df_driver.drop_duplicates(subset=['Biển số xe'], keep='last')
        
        # --- CBNV ---
        # Loại bỏ NV trùng tên
        if 'Full Name' in df_cbnv.columns:
            df_cbnv = df_cbnv.drop_duplicates(subset=['Full Name'], keep='first')

        # C. MERGE DỮ LIỆU
        # Merge Booking - Driver
        df_final = df_booking.merge(df_driver, on='Biển số xe', how='left', suffixes=('', '_Driver'))
        
        # Merge Booking - CBNV
        df_final = df_final.merge(df_cbnv, left_on='Người sử dụng xe', right_on='Full Name', how='left')

        # D. XỬ LÝ THÊM
        df_final['Ngày khởi hành'] = pd.to_datetime(df_final['Ngày khởi hành'], errors='coerce')
        df_final['Tháng'] = df_final['Ngày khởi hành'].dt.strftime('%Y-%m')
        
        # Điền dữ liệu thiếu cho biểu đồ Sunburst
        cols_fill = {'Location': 'Unknown', 'Công ty': 'Other', 'BU': 'Other'}
        for col, val in cols_fill.items():
            if col in df_final.columns:
                df_final[col] = df_final[col].fillna(val)
        
        # Tạo cột phân loại "Nội thành/Tỉnh"
        def phan_loai(route):
            s = str(route).lower()
            if 'tỉnh' in s or ('tp.' in s and 'hồ chí minh' not in s): return 'Đi Tỉnh'
            return 'Nội Thành'
            
        if 'Lộ trình' in df_final.columns:
            df_final['Phạm Vi'] = df_final['Lộ trình'].apply(phan_loai)
        else:
            df_final['Phạm Vi'] = 'N/A'

        return df_final

    except Exception as e:
        st.error(f"Lỗi khi đọc file Excel: {e}")
        return pd.DataFrame()

# --- 4. LOGIC CHÍNH ---
if uploaded_file is not None:
    df = process_data(uploaded_file)
    
    if not df.empty:
        # --- BỘ LỌC DRILL-DOWN ---
        st.sidebar.markdown("---")
        st.sidebar.header("🔍 Bộ Lọc Drill-down")
        
        # Level 1
        locs = sorted(df['Location'].unique())
        sel_loc = st.sidebar.multiselect("1. Khu Vực", locs, default=locs)
        df_l1 = df[df['Location'].isin(sel_loc)]
        
        # Level 2
        comps = sorted(df_l1['Công ty'].unique())
        sel_comp = st.sidebar.multiselect("2. Công Ty", comps, default=comps)
        df_l2 = df_l1[df_l1['Công ty'].isin(sel_comp)]
        
        # Level 3
        bus = sorted(df_l2['BU'].unique())
        sel_bu = st.sidebar.multiselect("3. Bộ Phận (BU)", bus, default=bus)
        df_filtered = df_l2[df_l2['BU'].isin(sel_bu)]
        
        # --- KPI CARDS ---
        col1, col2, col3, col4 = st.columns(4)
        with col1: 
            st.markdown(f"<div class='kpi-card'><div class='kpi-value'>{len(df_filtered)}</div><div class='kpi-label'>Tổng Chuyến</div></div>", unsafe_allow_html=True)
        with col2: 
            top_user = df_filtered['Người sử dụng xe'].mode()[0] if not df_filtered.empty else "-"
            st.markdown(f"<div class='kpi-card'><div class='kpi-value' style='font-size:20px'>{top_user}</div><div class='kpi-label'>Top User</div></div>", unsafe_allow_html=True)
        with col3: 
            st.markdown(f"<div class='kpi-card'><div class='kpi-value'>{df_filtered['Biển số xe'].nunique()}</div><div class='kpi-label'>Xe Hoạt Động</div></div>", unsafe_allow_html=True)
        with col4: 
            tinh_count = len(df_filtered[df_filtered['Phạm Vi']=='Đi Tỉnh'])
            st.markdown(f"<div class='kpi-card'><div class='kpi-value'>{tinh_count}</div><div class='kpi-label'>Chuyến Đi Tỉnh</div></div>", unsafe_allow_html=True)
        
        st.markdown("---")
        
        # --- TABS & CHARTS ---
        tab1, tab2, tab3 = st.tabs(["📊 Phân Cấp (Drill-down)", "📈 Xu Hướng & Top", "📋 Dữ Liệu"])
        
        with tab1:
            c1, c2 = st.columns(2)
            with c1:
                st.subheader("Sunburst: Cấu trúc Tổ chức")
                if not df_filtered.empty:
                    fig = px.sunburst(df_filtered, path=['Location', 'Công ty', 'BU'], height=500, title="Tương tác để xem chi tiết")
                    st.plotly_chart(fig, use_container_width=True)
            with c2:
                st.subheader("Treemap: Phân bổ Số chuyến")
                if not df_filtered.empty:
                    df_tree = df_filtered.groupby(['Location', 'Công ty', 'BU']).size().reset_index(name='Count')
                    fig = px.treemap(df_tree, path=['Location', 'Công ty', 'BU'], values='Count', color='Count', height=500)
                    st.plotly_chart(fig, use_container_width=True)
                    
        with tab2:
            c1, c2 = st.columns([2,1])
            with c1:
                st.subheader("Xu hướng theo Tháng")
                if 'Tháng' in df_filtered.columns:
                    df_trend = df_filtered.groupby('Tháng').size().reset_index(name='Count')
                    fig = px.area(df_trend, x='Tháng', y='Count', markers=True)
                    st.plotly_chart(fig, use_container_width=True)
            with c2:
                st.subheader("Tỷ lệ Lộ trình")
                df_pie = df_filtered['Phạm Vi'].value_counts().reset_index()
                df_pie.columns = ['Phạm Vi', 'Count']
                fig = px.pie(df_pie, values='Count', names='Phạm Vi', hole=0.5)
                st.plotly_chart(fig, use_container_width=True)
                
        with tab3:
            st.dataframe(df_filtered)
            
    else:
        st.warning("File Excel không chứa dữ liệu hợp lệ hoặc lỗi đọc file.")
else:
    # Màn hình chờ khi chưa upload file
    st.info("👋 Vui lòng tải file 'Booking car.xlsx' lên để bắt đầu phân tích!")