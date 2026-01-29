import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import numpy as np

# --- 1. CẤU HÌNH TRANG & CSS PRO ---
st.set_page_config(
    page_title="Fleet Management Pro Dashboard",
    page_icon="🚘",
    layout="wide",
    initial_sidebar_state="expanded"
)

# CSS Styling nâng cao cho Card và Layout
st.markdown("""
<style>
    /* Tổng thể */
    .main { background-color: #f8f9fa; }
    
    /* KPI Cards */
    .kpi-container {
        background-color: white; padding: 20px; border-radius: 12px;
        box-shadow: 0 4px 6px rgba(0,0,0,0.05); border-left: 5px solid #3498db;
        text-align: center; margin-bottom: 10px;
    }
    .kpi-title { font-size: 14px; color: #7f8c8d; font-weight: 600; text-transform: uppercase; letter-spacing: 0.5px; }
    .kpi-value { font-size: 28px; font-weight: 800; color: #2c3e50; margin: 10px 0; }
    .kpi-delta { font-size: 12px; color: #27ae60; font-weight: 600; }
    
    /* Tabs */
    .stTabs [data-baseweb="tab-list"] { gap: 10px; }
    .stTabs [data-baseweb="tab"] {
        height: 50px; white-space: pre-wrap; background-color: white;
        border-radius: 8px 8px 0 0; padding-top: 10px; padding-bottom: 10px;
        box-shadow: 0 -2px 5px rgba(0,0,0,0.02);
    }
    .stTabs [aria-selected="true"] { background-color: #e8f4f8; color: #007bff; font-weight: bold; border-top: 3px solid #007bff; }
</style>
""", unsafe_allow_html=True)

# --- 2. HÀM XỬ LÝ DỮ LIỆU NÂNG CAO ---
@st.cache_data
def load_and_process_data(file):
    try:
        # A. ĐỌC FILE THÔNG MINH
        if file.name.endswith('.csv'):
            df = pd.read_csv(file, header=3) # Header dòng 4
        else:
            xl = pd.ExcelFile(file)
            # Tìm sheet chứa "Booking" và "Car"
            target_sheet = next((s for s in xl.sheet_names if "booking" in s.lower() and "car" in s.lower()), xl.sheet_names[0])
            df = pd.read_excel(file, sheet_name=target_sheet, header=3)

        # B. CHUẨN HÓA CỘT
        df.columns = [str(c).strip().replace('\n', ' ') for c in df.columns]
        
        # Mapping cột (Thêm các cột chi phí chi tiết)
        col_map = {
            'Ngày Tháng Năm': 'Date',
            'Biển số xe': 'Car_Plate',
            'Tên tài xế': 'Driver',
            'Bộ phận': 'Department',
            'Cost center': 'Cost_Center',
            'Km sử dụng': 'Km_Used',
            'Tổng chi phí': 'Total_Cost',
            'Lộ trình': 'Route',
            'Giờ khởi hành': 'Start_Time',
            'Giờ kết thúc': 'End_Time',
            'Ngoài giờ': 'OT_Hours',
            # Các cột thành phần chi phí (dựa trên file mẫu)
            'Chi phí nhiên liệu': 'Cost_Fuel',
            'Phí cầu đường': 'Cost_Toll',
            'VETC': 'Cost_VETC',
            'Sửa chữa': 'Cost_Repair',
            'Bảo dưỡng': 'Cost_Maintenance',
            'Tiền cơm': 'Cost_Meal'
        }
        
        # Chỉ giữ lại các cột có trong map và rename
        cols_present = [c for c in col_map.keys() if c in df.columns]
        df = df[cols_present].rename(columns=col_map)
        
        # Xóa dòng rỗng
        df.dropna(how='all', inplace=True)
        if 'Date' in df.columns:
            df['Date'] = pd.to_datetime(df['Date'], errors='coerce')
            df = df.dropna(subset=['Date'])
            
        # C. FEATURE ENGINEERING (TẠO CỘT MỚI CHO PHÂN TÍCH)
        
        # 1. Thời gian
        df['Month_Str'] = df['Date'].dt.strftime('%m-%Y')
        df['Day_Of_Week'] = df['Date'].dt.day_name() # Monday, Tuesday...
        df['Day_Index'] = df['Date'].dt.dayofweek    # 0, 1, 2... để sort
        
        # 2. Xử lý số liệu
        numeric_cols = ['Km_Used', 'Total_Cost', 'Cost_Fuel', 'Cost_Toll', 'Cost_VETC', 'Cost_Repair', 'Cost_Maintenance', 'Cost_Meal']
        for col in numeric_cols:
            if col in df.columns:
                df[col] = pd.to_numeric(df[col], errors='coerce').fillna(0)
        
        # Gom nhóm chi phí "Khác" (Tổng - các thành phần đã biết)
        known_cost_cols = [c for c in numeric_cols if c in df.columns and c != 'Total_Cost' and c != 'Km_Used']
        df['Cost_Other'] = df['Total_Cost'] - df[known_cost_cols].sum(axis=1)
        df['Cost_Other'] = df['Cost_Other'].apply(lambda x: x if x > 0 else 0) # Tránh số âm do làm tròn

        # 3. Phân loại Lộ trình (Heuristic đơn giản)
        if 'Route' in df.columns:
            df['Route'] = df['Route'].astype(str).fillna("")
            def classify_route(s):
                s = s.lower()
                # Nếu lộ trình chứa các từ khóa địa phương hoặc rất ngắn -> Nội tỉnh
                if len(s) < 5 or any(x in s for x in ['hcm', 'sài gòn', 'q1', 'q7', 'thủ đức', 'bình thạnh', 'nội thành', 'city']):
                    return 'Nội Tỉnh'
                return 'Ngoại Tỉnh'
            df['Route_Type'] = df['Route'].apply(classify_route)
        else:
            df['Route_Type'] = 'Không xác định'

        # 4. Xử lý Cost Center & Bộ phận
        if 'Department' in df.columns:
            df['Department'] = df['Department'].astype(str).str.strip()
        if 'Cost_Center' in df.columns:
            df['Cost_Center'] = df['Cost_Center'].astype(str).str.strip().str.replace('.0', '', regex=False)

        # 5. Xử lý Giờ (Lấy giờ bắt đầu để vẽ Heatmap)
        if 'Start_Time' in df.columns:
            # Cố gắng convert sang string rồi lấy 2 ký tự đầu
            df['Start_Hour'] = df['Start_Time'].astype(str).str.extract(r'(\d{1,2})').astype(float).fillna(0).astype(int)
        else:
            df['Start_Hour'] = 0

        return df
    except Exception as e:
        st.error(f"Lỗi xử lý dữ liệu: {e}")
        return pd.DataFrame()

# --- 3. HÀM VẼ BIỂU ĐỒ (HELPER) ---
def card_metric(title, value, suffix="", delta=""):
    st.markdown(f"""
    <div class="kpi-container">
        <div class="kpi-title">{title}</div>
        <div class="kpi-value">{value} <span style="font-size:16px; color:#999">{suffix}</span></div>
        <div class="kpi-delta">{delta}</div>
    </div>
    """, unsafe_allow_html=True)

# --- 4. GIAO DIỆN CHÍNH ---
st.title("📊 Fleet Operations Center")
st.markdown("Hệ thống phân tích hiệu suất và chi phí đội xe toàn diện")

# Upload File
uploaded_file = st.sidebar.file_uploader("📂 Upload Data (Excel/CSV)", type=['xlsx', 'csv'])

if uploaded_file:
    df = load_and_process_data(uploaded_file)
    
    if not df.empty:
        # --- SIDEBAR FILTERS ---
        st.sidebar.markdown("---")
        st.sidebar.header("🔍 Bộ Lọc Dữ Liệu")
        
        # Filter Tháng
        all_months = sorted(df['Month_Str'].unique())
        sel_month = st.sidebar.multiselect("Tháng", all_months, default=all_months)
        
        # Filter Department
        all_depts = sorted(df['Department'].unique())
        sel_dept = st.sidebar.multiselect("Bộ phận / BU", all_depts, default=all_depts)
        
        # Filter Cost Center
        if 'Cost_Center' in df.columns:
            all_cc = sorted(df['Cost_Center'].unique())
            sel_cc = st.sidebar.multiselect("Cost Center", all_cc, default=[])
        
        # Filter Logic
        mask = df['Month_Str'].isin(sel_month) & df['Department'].isin(sel_dept)
        if 'Cost_Center' in df.columns and sel_cc:
            mask = mask & df['Cost_Center'].isin(sel_cc)
            
        df_sub = df[mask]
        
        if df_sub.empty:
            st.warning("Không có dữ liệu phù hợp bộ lọc.")
            st.stop()

        # --- KPI OVERVIEW ROW ---
        tot_cost = df_sub['Total_Cost'].sum()
        tot_km = df_sub['Km_Used'].sum()
        tot_trips = len(df_sub)
        avg_cost_km = tot_cost / tot_km if tot_km > 0 else 0
        
        col1, col2, col3, col4 = st.columns(4)
        with col1: card_metric("Tổng Chi Phí", f"{tot_cost:,.0f}", "VNĐ")
        with col2: card_metric("Tổng Km Vận Hành", f"{tot_km:,.0f}", "Km")
        with col3: card_metric("Số Chuyến Xe", f"{tot_trips:,}", "Trip")
        with col4: card_metric("Chi Phí / Km", f"{avg_cost_km:,.0f}", "VNĐ/Km")

        # --- TABS LAYOUT ---
        tab1, tab2, tab3, tab4, tab5 = st.tabs([
            "💰 Tài Chính & Ngân Sách", 
            "🚗 Đội Xe & Hiệu Suất", 
            "🗺️ Lộ Trình & Vận Hành", 
            "👥 Nhân Sự & Sử Dụng",
            "📄 Dữ Liệu Chi Tiết"
        ])

        # === TAB 1: TÀI CHÍNH ===
        with tab1:
            c1, c2 = st.columns([2, 1])
            with c1:
                st.subheader("Cấu Trúc Chi Phí Vận Hành")
                # Chuẩn bị dữ liệu cho Stacked Bar hoặc Pie
                cost_cols = {'Cost_Fuel': 'Nhiên liệu', 'Cost_Toll': 'Cầu đường', 'Cost_VETC': 'VETC', 
                             'Cost_Repair': 'Sửa chữa', 'Cost_Maintenance': 'Bảo dưỡng', 'Cost_Meal': 'Tiền cơm', 'Cost_Other': 'Khác'}
                # Chỉ lấy cột có trong df
                valid_cost_cols = {k:v for k,v in cost_cols.items() if k in df_sub.columns}
                
                cost_sum = df_sub[list(valid_cost_cols.keys())].sum().rename(index=valid_cost_cols).reset_index()
                cost_sum.columns = ['Loại Chi Phí', 'Giá Trị']
                
                # --- SỬA LỖI: Dùng px.pie với hole thay vì px.donut ---
                fig_struct = px.pie(cost_sum, values='Giá Trị', names='Loại Chi Phí', hole=0.4, 
                                    color_discrete_sequence=px.colors.qualitative.Pastel)
                st.plotly_chart(fig_struct, use_container_width=True)
                
            with c2:
                st.subheader("Top Cost Center")
                if 'Cost_Center' in df_sub.columns:
                    cc_stat = df_sub.groupby('Cost_Center')['Total_Cost'].sum().nlargest(10).reset_index()
                    fig_cc = px.bar(cc_stat, x='Total_Cost', y='Cost_Center', orientation='h', 
                                    text_auto='.2s', color='Total_Cost', color_continuous_scale='Blues')
                    st.plotly_chart(fig_cc, use_container_width=True)
            
            st.subheader("Xu Hướng Chi Phí & Km Theo Thời Gian")
            trend_df = df_sub.groupby('Date')[['Total_Cost', 'Km_Used']].sum().reset_index()
            fig_trend = go.Figure()
            fig_trend.add_trace(go.Bar(x=trend_df['Date'], y=trend_df['Total_Cost'], name='Chi Phí', marker_color='#3498db'))
            fig_trend.add_trace(go.Scatter(x=trend_df['Date'], y=trend_df['Km_Used'], name='Km', yaxis='y2', line=dict(color='#e74c3c', width=3)))
            fig_trend.update_layout(yaxis2=dict(overlaying='y', side='right'), hovermode='x unified')
            st.plotly_chart(fig_trend, use_container_width=True)

        # === TAB 2: ĐỘI XE ===
        with tab2:
            c1, c2 = st.columns(2)
            with c1:
                st.subheader("Top Xe Hoạt Động (Km)")
                car_km = df_sub.groupby('Car_Plate')['Km_Used'].sum().nlargest(10).reset_index()
                fig_car = px.bar(car_km, x='Car_Plate', y='Km_Used', color='Km_Used', title="Top 10 Xe Chạy Nhiều Nhất", color_continuous_scale='Viridis')
                st.plotly_chart(fig_car, use_container_width=True)
            
            with c2:
                st.subheader("Hiệu Quả Chi Phí (Cost/Km) Từng Xe")
                car_eff = df_sub.groupby('Car_Plate')[['Total_Cost', 'Km_Used']].sum().reset_index()
                car_eff = car_eff[car_eff['Km_Used'] > 0] # Tránh chia 0
                car_eff['Cost_Per_Km'] = car_eff['Total_Cost'] / car_eff['Km_Used']
                
                fig_eff = px.scatter(car_eff, x='Km_Used', y='Total_Cost', size='Cost_Per_Km', color='Car_Plate',
                                     hover_data=['Cost_Per_Km'], title="Tương quan Chi phí vs Km (Bóng to = Tốn kém/km)")
                st.plotly_chart(fig_eff, use_container_width=True)

        # === TAB 3: LỘ TRÌNH ===
        with tab3:
            c1, c2 = st.columns(2)
            with c1:
                st.subheader("Tỷ Lệ Nội Tỉnh vs Ngoại Tỉnh")
                route_type = df_sub['Route_Type'].value_counts().reset_index()
                route_type.columns = ['Loại', 'Số chuyến']
                
                # --- SỬA LỖI: Dùng px.pie với hole thay vì px.donut ---
                fig_route = px.pie(route_type, values='Số chuyến', names='Loại', hole=0.5, 
                                   color_discrete_sequence=['#2ecc71', '#e67e22'])
                st.plotly_chart(fig_route, use_container_width=True)
            
            with c2:
                st.subheader("Mật Độ Sử Dụng (Heatmap)")
                # Heatmap Thứ vs Giờ
                heatmap_data = df_sub.groupby(['Day_Of_Week', 'Start_Hour']).size().reset_index(name='Count')
                # Sắp xếp thứ
                days_order = ['Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Saturday', 'Sunday']
                
                fig_heat = px.density_heatmap(heatmap_data, x='Start_Hour', y='Day_Of_Week', z='Count', 
                                              category_orders={'Day_Of_Week': days_order},
                                              color_continuous_scale='RdBu_r', title="Mật độ đặt xe theo Giờ & Thứ")
                st.plotly_chart(fig_heat, use_container_width=True)

        # === TAB 4: NHÂN SỰ ===
        with tab4:
            c1, c2 = st.columns(2)
            with c1:
                st.subheader("Top Bộ Phận Đặt Xe")
                dept_stat = df_sub.groupby('Department')['Total_Cost'].sum().nlargest(10).reset_index().sort_values('Total_Cost')
                fig_dept = px.bar(dept_stat, x='Total_Cost', y='Department', orientation='h', text_auto='.2s')
                st.plotly_chart(fig_dept, use_container_width=True)
            
            with c2:
                st.subheader("Top Tài Xế (Theo Km)")
                driver_stat = df_sub.groupby('Driver')['Km_Used'].sum().nlargest(10).reset_index().sort_values('Km_Used')
                fig_driver = px.bar(driver_stat, x='Km_Used', y='Driver', orientation='h', color='Km_Used')
                st.plotly_chart(fig_driver, use_container_width=True)

        # === TAB 5: DATA ===
        with tab5:
            st.dataframe(df_sub.style.format({
                "Total_Cost": "{:,.0f}", 
                "Km_Used": "{:,.0f}",
                "Cost_Fuel": "{:,.0f}"
            }), height=600)

else:
    st.info("👋 Xin chào! Vui lòng tải lên file Excel (Data-SuDungXe) để bắt đầu phân tích.")