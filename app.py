import streamlit as st
import pandas as pd
import plotly.express as px

# --- 1. CẤU HÌNH TRANG (Full Width) ---
st.set_page_config(page_title="Executive Fleet Dashboard", page_icon="🚘", layout="wide")

# --- 2. CSS TÙY CHỈNH (Làm đẹp giống Power BI) ---
st.markdown("""
<style>
    /* Tổng thể */
    .main {background-color: #f5f7f9;}
    
    /* Header */
    .header-title {font-size: 28px; font-weight: 700; color: #1e3a8a; margin-bottom: 0px;}
    .header-subtitle {font-size: 14px; color: #64748b; margin-bottom: 20px;}
    
    /* Khung Bộ lọc (Filter Container) */
    .filter-container {
        background-color: #ffffff;
        padding: 20px;
        border-radius: 10px;
        box-shadow: 0 4px 6px -1px rgba(0, 0, 0, 0.1);
        margin-bottom: 20px;
        border-top: 4px solid #3b82f6;
    }
    
    /* KPI Card Style */
    .kpi-card {
        background-color: white;
        padding: 20px;
        border-radius: 10px;
        box-shadow: 0 2px 4px rgba(0,0,0,0.05);
        text-align: center;
        border: 1px solid #e2e8f0;
    }
    .kpi-value {font-size: 32px; font-weight: 800; color: #2563eb;}
    .kpi-label {font-size: 13px; font-weight: 600; color: #64748b; text-transform: uppercase; letter-spacing: 1px;}
    
    /* Chart Container */
    .chart-box {
        background-color: white;
        padding: 15px;
        border-radius: 10px;
        box-shadow: 0 2px 4px rgba(0,0,0,0.05);
        margin-bottom: 20px;
    }
</style>
""", unsafe_allow_html=True)

# --- 3. HÀM XỬ LÝ DỮ LIỆU (Giữ nguyên logic sửa lỗi) ---
@st.cache_data
def process_data(file):
    try:
        xls = pd.ExcelFile(file)
        # Đọc dữ liệu
        df_driver_raw = pd.read_excel(xls, sheet_name='Driver', header=None)
        try:
            header_idx = df_driver_raw[df_driver_raw.eq("Biển số xe").any(axis=1)].index[0]
        except:
            header_idx = 2
        df_driver = pd.read_excel(xls, sheet_name='Driver', header=header_idx)
        df_cbnv = pd.read_excel(xls, sheet_name='CBNV', header=1)
        df_booking = pd.read_excel(xls, sheet_name='Booking car', header=0)

        # Làm sạch
        df_driver.columns = df_driver.columns.str.replace('\n', ' ').str.strip()
        if 'Biển số xe' in df_driver.columns:
            df_driver = df_driver.drop_duplicates(subset=['Biển số xe'], keep='last')
        if 'Full Name' in df_cbnv.columns:
            df_cbnv = df_cbnv.drop_duplicates(subset=['Full Name'], keep='first')

        # Merge
        df_final = df_booking.merge(df_driver, on='Biển số xe', how='left', suffixes=('', '_Driver'))
        df_final = df_final.merge(df_cbnv, left_on='Người sử dụng xe', right_on='Full Name', how='left')

        # Xử lý cột
        df_final['Ngày khởi hành'] = pd.to_datetime(df_final['Ngày khởi hành'], errors='coerce')
        df_final['Tháng'] = df_final['Ngày khởi hành'].dt.strftime('%Y-%m')
        
        # Ép kiểu string để tránh lỗi
        cols_fill = {'Location': 'Unknown', 'Công ty': 'Other', 'BU': 'Other'}
        for col, val in cols_fill.items():
            if col in df_final.columns:
                df_final[col] = df_final[col].fillna(val).astype(str)
        
        # Phân loại
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
        return pd.DataFrame()

# --- 4. GIAO DIỆN CHÍNH ---

# Header Section
c1, c2 = st.columns([3, 1])
with c1:
    st.markdown('<div class="header-title">🚘 FLEET MANAGEMENT DASHBOARD</div>', unsafe_allow_html=True)
    st.markdown('<div class="header-subtitle">Báo cáo quản trị vận hành xe & chi phí</div>', unsafe_allow_html=True)
with c2:
    uploaded_file = st.file_uploader("📂 Upload File Excel", type=["xlsx"])

if uploaded_file is not None:
    df = process_data(uploaded_file)
    
    if not df.empty:
        # --- SECTION: BỘ LỌC THÔNG MINH (SLICER) ---
        # Đóng khung bộ lọc lại cho gọn
        st.markdown('<div class="filter-container">', unsafe_allow_html=True)
        st.write("**📌 Bộ Lọc Dữ Liệu (Drill-Down Logic)**")
        
        f1, f2, f3, f4 = st.columns(4)
        
        with f1:
            # Lọc Khu vực
            locs = sorted(df['Location'].unique())
            sel_loc = st.multiselect("1. Chọn Khu Vực", locs, default=locs)
            df_l1 = df[df['Location'].isin(sel_loc)]
            
        with f2:
            # Lọc Công ty (Chỉ hiện cty thuộc Khu vực đã chọn)
            comps = sorted(df_l1['Công ty'].unique())
            sel_comp = st.multiselect("2. Chọn Công Ty", comps, default=comps)
            df_l2 = df_l1[df_l1['Công ty'].isin(sel_comp)]
            
        with f3:
            # Lọc BU
            bus = sorted(df_l2['BU'].unique())
            sel_bu = st.multiselect("3. Chọn Bộ Phận (BU)", bus, default=bus)
            df_filtered = df_l2[df_l2['BU'].isin(sel_bu)]
            
        with f4:
            # Lọc Tháng (Thêm cái này cho tiện)
            months = sorted(df['Tháng'].dropna().unique())
            sel_month = st.multiselect("4. Chọn Tháng", months, default=months)
            if sel_month:
                df_filtered = df_filtered[df_filtered['Tháng'].isin(sel_month)]

        st.markdown('</div>', unsafe_allow_html=True)

        # --- SECTION: KPI CARDS ---
        k1, k2, k3, k4 = st.columns(4)
        
        total_trips = len(df_filtered)
        active_cars = df_filtered['Biển số xe'].nunique()
        top_user = df_filtered['Người sử dụng xe'].mode()[0] if not df_filtered.empty else "-"
        # Giả sử 1 chuyến đi tỉnh = 1
        province_trips = len(df_filtered[df_filtered['Phạm Vi'] == 'Đi Tỉnh'])

        with k1:
            st.markdown(f"""<div class="kpi-card">
                            <div class="kpi-label">Tổng Số Chuyến</div>
                            <div class="kpi-value">{total_trips}</div>
                        </div>""", unsafe_allow_html=True)
        with k2:
            st.markdown(f"""<div class="kpi-card">
                            <div class="kpi-label">Số Xe Vận Hành</div>
                            <div class="kpi-value">{active_cars}</div>
                        </div>""", unsafe_allow_html=True)
        with k3:
            st.markdown(f"""<div class="kpi-card">
                            <div class="kpi-label">Chuyến Đi Tỉnh</div>
                            <div class="kpi-value">{province_trips}</div>
                        </div>""", unsafe_allow_html=True)
        with k4:
             st.markdown(f"""<div class="kpi-card">
                            <div class="kpi-label">Nhân sự đi nhiều nhất</div>
                            <div class="kpi-value" style="font-size:18px; margin-top:10px">{top_user}</div>
                        </div>""", unsafe_allow_html=True)

        st.markdown("<br>", unsafe_allow_html=True) # Spacer

        # --- SECTION: CHARTS GRID ---
        
        # Hàng 1: Tổng quan cấu trúc & Xu hướng (2 biểu đồ lớn)
        row1_1, row1_2 = st.columns([1, 1])
        
        with row1_1:
            st.markdown('<div class="chart-box">', unsafe_allow_html=True)
            st.subheader("📊 Cấu trúc Vận hành (Sunburst)")
            if not df_filtered.empty:
                fig_sun = px.sunburst(
                    df_filtered, 
                    path=['Location', 'Công ty', 'BU'], 
                    color_discrete_sequence=px.colors.qualitative.Prism,
                    height=400
                )
                fig_sun.update_layout(margin=dict(t=0, l=0, r=0, b=0))
                st.plotly_chart(fig_sun, use_container_width=True)
            st.markdown('</div>', unsafe_allow_html=True)
            
        with row1_2:
            st.markdown('<div class="chart-box">', unsafe_allow_html=True)
            st.subheader("📈 Xu hướng Đặt xe (Trend)")
            if 'Tháng' in df_filtered.columns and not df_filtered.empty:
                df_trend = df_filtered.groupby('Tháng').size().reset_index(name='Số chuyến')
                fig_line = px.area(df_trend, x='Tháng', y='Số chuyến', 
                                   line_shape='spline',
                                   color_discrete_sequence=['#3b82f6'])
                fig_line.update_layout(xaxis_title=None, yaxis_title=None, height=400, margin=dict(t=20, l=0, r=0, b=0))
                st.plotly_chart(fig_line, use_container_width=True)
            st.markdown('</div>', unsafe_allow_html=True)

        # Hàng 2: Chi tiết (Treemap & Top List)
        row2_1, row2_2 = st.columns([2, 1])
        
        with row2_1:
            st.markdown('<div class="chart-box">', unsafe_allow_html=True)
            st.subheader("🏢 Phân bổ Chuyến đi theo Phòng ban (Treemap)")
            if not df_filtered.empty:
                df_tree = df_filtered.groupby(['Công ty', 'BU']).size().reset_index(name='Count')
                fig_tree = px.treemap(
                    df_tree, 
                    path=['Công ty', 'BU'], 
                    values='Count',
                    color='Count',
                    color_continuous_scale='Blues',
                    height=400
                )
                fig_tree.update_layout(margin=dict(t=0, l=0, r=0, b=0))
                st.plotly_chart(fig_tree, use_container_width=True)
            st.markdown('</div>', unsafe_allow_html=True)
            
        with row2_2:
            st.markdown('<div class="chart-box">', unsafe_allow_html=True)
            st.subheader("🏆 Top Tài Xế")
            if not df_filtered.empty:
                top_driver = df_filtered['Tên tài xế'].value_counts().head(7).reset_index()
                top_driver.columns = ['Tài xế', 'Số chuyến']
                fig_bar = px.bar(top_driver, x='Số chuyến', y='Tài xế', orientation='h', text='Số chuyến', color_discrete_sequence=['#1e40af'])
                fig_bar.update_layout(yaxis={'categoryorder':'total ascending'}, height=400, margin=dict(t=0, l=0, r=0, b=0))
                st.plotly_chart(fig_bar, use_container_width=True)
            st.markdown('</div>', unsafe_allow_html=True)

        # --- SECTION: DATA TABLE (Ẩn trong Expander cho gọn) ---
        with st.expander("📂 Xem dữ liệu chi tiết (Excel View)"):
            st.dataframe(df_filtered, use_container_width=True)

    else:
        st.error("File không hợp lệ. Vui lòng kiểm tra lại.")
else:
    # Màn hình chờ đẹp
    st.info("👋 Xin chào! Vui lòng tải file **Booking car.xlsx** lên để hiển thị Dashboard.")