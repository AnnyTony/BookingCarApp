import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
from plotly.subplots import make_subplots
from datetime import datetime

# --- 1. CẤU HÌNH TRANG ---
st.set_page_config(
    page_title="Dashboard Đội Xe Toàn Diện",
    page_icon="🚘",
    layout="wide",
    initial_sidebar_state="expanded"
)

# --- CSS CUSTOM: 3D UI & GLASSMORPHISM ---
st.markdown("""
<style>
    .stApp { background-color: #f0f4f8; }
    
    /* 3D Sidebar */
    [data-testid="stSidebar"] {
        background: linear-gradient(180deg, #ffffff 0%, #eef2f6 100%);
        box-shadow: 4px 0 15px rgba(0,0,0,0.05);
        border-right: 1px solid #dae1e7;
    }
    
    /* Card Design */
    .dashboard-card {
        background: white; padding: 20px; border-radius: 16px;
        box-shadow: 0 4px 20px rgba(0,0,0,0.06); margin-bottom: 20px;
        border: 1px solid #ffffff;
        transition: transform 0.3s ease;
    }
    .dashboard-card:hover { transform: translateY(-5px); box-shadow: 0 8px 25px rgba(0,0,0,0.1); }
    
    /* KPI Box */
    .kpi-box {
        background: white; padding: 20px; border-radius: 14px;
        border-left: 6px solid #3b82f6;
        box-shadow: 0 4px 15px rgba(0,0,0,0.05);
    }
    .kpi-label { font-size: 12px; color: #64748b; font-weight: 700; text-transform: uppercase; letter-spacing: 0.8px; }
    .kpi-val { font-size: 26px; font-weight: 800; color: #1e293b; margin: 8px 0; }
    .kpi-sub { font-size: 12px; color: #10b981; font-weight: 600; }
    
    /* Filter Box */
    .filter-box {
        background: white; padding: 20px; border-radius: 12px;
        box-shadow: inset 0 2px 4px rgba(0,0,0,0.03); border: 1px solid #e2e8f0;
    }
    
    /* Tabs */
    .stTabs [data-baseweb="tab-list"] { background: white; padding: 8px; border-radius: 30px; gap: 5px; box-shadow: 0 2px 10px rgba(0,0,0,0.03); }
    .stTabs [aria-selected="true"] { background-color: #e0f2fe; color: #0284c7; border-radius: 25px; font-weight: bold; border-bottom: none; }
</style>
""", unsafe_allow_html=True)

# --- 2. HÀM XỬ LÝ DỮ LIỆU NÂNG CAO ---
@st.cache_data
def load_data(file):
    try:
        # Đọc file
        if file.name.endswith('.csv'):
            df = pd.read_csv(file, header=3)
        else:
            xl = pd.ExcelFile(file)
            target = next((s for s in xl.sheet_names if "booking" in s.lower()), xl.sheet_names[0])
            df = pd.read_excel(file, sheet_name=target, header=3)

        df.columns = [str(c).strip().replace('\n', ' ') for c in df.columns]
        
        # Mapping cột
        col_map = {
            'Ngày Tháng Năm': 'Date', 'Biển số xe': 'Car', 'Tên tài xế': 'Driver',
            'Bộ phận': 'Dept', 'Cost center': 'CostCenter', 'Km sử dụng': 'Km',
            'Tổng chi phí': 'Cost', 'Lộ trình': 'Route', 'Người sử dụng xe': 'User',
            'Chi phí nhiên liệu': 'Fuel', 'Phí cầu đường': 'Toll', 'Sửa chữa': 'Repair',
            'Giờ khởi hành': 'Start_Time', 'Giờ kết thúc': 'End_Time', 'Công Ty': 'Company'
        }
        cols = [c for c in col_map.keys() if c in df.columns]
        df = df[cols].rename(columns=col_map)
        
        df.dropna(how='all', inplace=True)
        if 'Date' in df.columns:
            df['Date'] = pd.to_datetime(df['Date'], errors='coerce')
            df = df.dropna(subset=['Date'])
            df['Tháng'] = df['Date'].dt.strftime('%m-%Y')
            df['SortMonth'] = df['Date'].dt.to_period('M')

        # Chuyển số
        for c in ['Km', 'Cost', 'Fuel', 'Toll', 'Repair']:
            if c in df.columns: df[c] = pd.to_numeric(df[c], errors='coerce').fillna(0)
            
        # Xử lý Lộ Trình
        if 'Route' in df.columns:
            df['Route'] = df['Route'].astype(str).fillna("")
            df['Route_Type'] = df['Route'].apply(lambda s: 'Nội Tỉnh' if len(str(s)) < 5 or any(k in str(s).lower() for k in ['hcm', 'sài gòn', 'q1', 'q7', 'city']) else 'Ngoại Tỉnh')
        
        # --- TÍNH TOÁN THỜI GIAN (DURATION) ---
        # Logic: Giả sử cùng ngày. Nếu Start/End lỗi -> Duration = 0
        if 'Start_Time' in df.columns and 'End_Time' in df.columns:
            def calc_duration(row):
                try:
                    # Chuyển đổi sang datetime object (chỉ lấy giờ)
                    s = pd.to_datetime(str(row['Start_Time']), format='%H:%M:%S', errors='coerce')
                    e = pd.to_datetime(str(row['End_Time']), format='%H:%M:%S', errors='coerce')
                    if pd.notnull(s) and pd.notnull(e):
                        diff = (e - s).total_seconds() / 3600 # Ra số giờ
                        return diff if diff > 0 else 0
                    return 0
                except: return 0
            df['Duration_Hours'] = df.apply(calc_duration, axis=1)
        else:
            df['Duration_Hours'] = 0

        # Làm sạch Text
        for c in ['Dept', 'Driver', 'Car', 'Company']:
            if c in df.columns: df[c] = df[c].astype(str).str.strip()
            
        return df
    except Exception as e:
        return pd.DataFrame()

# --- 3. HELPER FUNCTIONS ---
def draw_kpi(title, val, unit, color, sub_text=""):
    st.markdown(f"""
    <div class="kpi-box" style="border-left-color: {color}">
        <div class="kpi-label">{title}</div>
        <div class="kpi-val">{val}</div>
        <div class="kpi-sub">{unit} {sub_text}</div>
    </div>
    """, unsafe_allow_html=True)

# --- 4. MAIN APP ---
st.title("🚀 Fleet Commander Dashboard")
st.caption("Hệ thống quản trị & phân tích hiệu suất đội xe chuyên sâu")

# --- SIDEBAR ---
with st.sidebar:
    st.image("https://cdn-icons-png.flaticon.com/512/3097/3097144.png", width=50)
    st.markdown("### 🎛️ Bảng Điều Khiển")
    
    uploaded_file = st.file_uploader("Upload Data (Booking Car)", type=['xlsx', 'csv'])
    df = pd.DataFrame()
    if uploaded_file: df = load_data(uploaded_file)

    if not df.empty:
        st.write("")
        st.markdown('<div class="filter-box">', unsafe_allow_html=True)
        st.markdown("**🔍 Bộ Lọc Dữ Liệu**")
        
        # Sort months
        if 'SortMonth' in df.columns:
            months = sorted(df['Tháng'].unique(), key=lambda x: df[df['Tháng']==x]['SortMonth'].iloc[0])
        else: months = sorted(df['Tháng'].unique())
            
        sel_month = st.multiselect("Tháng", months, default=months)
        sel_dept = st.multiselect("Bộ Phận / BU", sorted(df['Dept'].unique()), default=sorted(df['Dept'].unique()))
        
        st.markdown('</div>', unsafe_allow_html=True)
        
        # Filter Logic
        mask = pd.Series(True, index=df.index)
        if sel_month: mask &= df['Tháng'].isin(sel_month)
        if sel_dept: mask &= df['Dept'].isin(sel_dept)
        df_sub = df[mask]
    else: df_sub = pd.DataFrame()

if not df_sub.empty:
    # --- GLOBAL KPIs ---
    c1, c2, c3, c4 = st.columns(4)
    total_cost = df_sub['Cost'].sum()
    total_km = df_sub['Km'].sum()
    total_hours = df_sub['Duration_Hours'].sum()
    total_trips = len(df_sub)
    
    # Tính occupancy đơn giản (Số giờ chạy / (Số xe * 9h * 26 ngày)) - Ước lượng
    unique_cars = df_sub['Car'].nunique()
    est_capacity_hours = unique_cars * 9 * 26 * len(sel_month) if len(sel_month) > 0 else 1
    occupancy_rate = (total_hours / est_capacity_hours) * 100 if est_capacity_hours > 0 else 0

    with c1: draw_kpi("Tổng Chi Phí", f"{total_cost:,.0f}", "VNĐ", "#ef4444")
    with c2: draw_kpi("Tổng Km", f"{total_km:,.0f}", "Km", "#3b82f6")
    with c3: draw_kpi("Tổng Giờ Vận Hành", f"{total_hours:,.0f}", "Giờ", "#f59e0b")
    with c4: draw_kpi("Số Chuyến Xe", f"{total_trips:,}", "Trips", "#10b981")

    st.write("")
    
    # --- TABS ---
    tab_overview, tab_perf, tab_rank, tab_explore = st.tabs([
        "📊 Tổng Quan", 
        "⚡ Hiệu Suất & Công Suất (New)", 
        "🏆 Bảng Xếp Hạng", 
        "🛠️ Tự Do Phân Tích"
    ])

    # === TAB 1: TỔNG QUAN ===
    with tab_overview:
        col_L, col_R = st.columns([2, 1])
        
        with col_L:
            st.markdown('<div class="dashboard-card">', unsafe_allow_html=True)
            st.markdown("### 🌊 Biểu Đồ Sóng: Chi Phí vs Quãng Đường")
            daily = df_sub.groupby('Date')[['Cost', 'Km']].sum().reset_index()
            
            fig_trend = make_subplots(specs=[[{"secondary_y": True}]])
            fig_trend.add_trace(go.Scatter(x=daily['Date'], y=daily['Cost'], name="Chi Phí", fill='tozeroy', line=dict(color='#ef4444')), secondary_y=False)
            fig_trend.add_trace(go.Scatter(x=daily['Date'], y=daily['Km'], name="Km", line=dict(color='#3b82f6', width=3)), secondary_y=True)
            
            fig_trend.update_layout(height=400, hovermode='x unified', margin=dict(t=10, b=10, l=10, r=10))
            fig_trend.update_yaxes(title_text="VNĐ", secondary_y=False)
            fig_trend.update_yaxes(title_text="Km", secondary_y=True)
            st.plotly_chart(fig_trend, use_container_width=True)
            st.markdown('</div>', unsafe_allow_html=True)

        with col_R:
            st.markdown('<div class="dashboard-card">', unsafe_allow_html=True)
            st.markdown("### 🏢 Phân Bổ Theo Công Ty")
            if 'Company' in df_sub.columns:
                comp_stats = df_sub['Company'].value_counts().reset_index()
                comp_stats.columns = ['Công Ty', 'Số Chuyến']
                fig_comp = px.pie(comp_stats, values='Số Chuyến', names='Công Ty', hole=0.6, color_discrete_sequence=px.colors.qualitative.Prism)
                fig_comp.update_layout(height=400, margin=dict(t=10, b=10))
                st.plotly_chart(fig_comp, use_container_width=True)
            else: st.warning("Không có cột 'Công Ty'")
            st.markdown('</div>', unsafe_allow_html=True)

    # === TAB 2: HIỆU SUẤT & CÔNG SUẤT (NEW FEATURE) ===
    with tab_perf:
        st.markdown("### ⚡ Phân Tích Sâu Về Hiệu Quả Sử Dụng Đội Xe")
        
        c1, c2 = st.columns(2)
        
        # 1. BIỂU ĐỒ CÔNG SUẤT (Utilization Rate)
        with c1:
            st.markdown('<div class="dashboard-card">', unsafe_allow_html=True)
            st.subheader("📈 Tỷ Lệ Xe Hoạt Động Theo Ngày (% Fleet Utilization)")
            st.caption("Có bao nhiêu % tổng số xe được sử dụng mỗi ngày?")
            
            # Tính tổng số xe duy nhất trong dữ liệu (Active Fleet)
            total_active_cars = df['Car'].nunique() 
            
            # Tính số xe hoạt động theo ngày
            daily_active = df_sub.groupby('Date')['Car'].nunique().reset_index()
            daily_active['Utilization'] = (daily_active['Car'] / total_active_cars) * 100
            
            fig_util = px.line(daily_active, x='Date', y='Utilization', markers=True, 
                               labels={'Utilization': '% Xe hoạt động'}, color_discrete_sequence=['#8b5cf6'])
            fig_util.add_hline(y=100, line_dash="dot", annotation_text="Max Capacity")
            fig_util.update_layout(yaxis_range=[0, 110], height=350)
            st.plotly_chart(fig_util, use_container_width=True)
            st.markdown('</div>', unsafe_allow_html=True)
            
        # 2. BIỂU ĐỒ TỶ LỆ LẤP ĐẦY (Occupancy Rate)
        with c2:
            st.markdown('<div class="dashboard-card">', unsafe_allow_html=True)
            st.subheader("⏳ Tỷ Lệ Lấp Đầy Thời Gian (Occupancy Rate)")
            st.caption("Xe nào chạy nhiều giờ nhất? (Giả định Full công suất = 200 giờ/tháng)")
            
            car_hours = df_sub.groupby('Car')['Duration_Hours'].sum().reset_index()
            # Giả định: 1 xe "chăm chỉ" chạy 200h/tháng.
            car_hours['Occupancy_Score'] = (car_hours['Duration_Hours'] / 200) * 100 
            car_hours = car_hours.sort_values('Duration_Hours', ascending=False).head(10)
            
            fig_occ = px.bar(car_hours, x='Occupancy_Score', y='Car', orientation='h', 
                             color='Occupancy_Score', color_continuous_scale='Viridis',
                             text_auto='.1f', labels={'Occupancy_Score': 'Điểm Lấp Đầy (Index)'})
            st.plotly_chart(fig_occ, use_container_width=True)
            st.markdown('</div>', unsafe_allow_html=True)

        # 3. SCATTER PLOT HIỆU SUẤT (Khôi phục)
        st.markdown('<div class="dashboard-card">', unsafe_allow_html=True)
        st.subheader("🎯 Ma Trận Hiệu Suất: Chi Phí vs Quãng Đường")
        car_perf = df_sub.groupby('Car')[['Cost', 'Km']].sum().reset_index()
        car_perf = car_perf[car_perf['Km'] > 0]
        
        if not car_perf.empty:
            fig_sc = px.scatter(car_perf, x='Km', y='Cost', size='Km', color='Car',
                                title="Xe nằm góc TRÊN TRÁI là kém hiệu quả (Tốn tiền - Đi ít)",
                                hover_name='Car')
            st.plotly_chart(fig_sc, use_container_width=True)
        st.markdown('</div>', unsafe_allow_html=True)

    # === TAB 3: BẢNG XẾP HẠNG (RANKINGS) ===
    with tab_rank:
        st.markdown("### 🏆 Hall of Fame")
        
        c1, c2 = st.columns(2)
        with c1:
            st.markdown('<div class="dashboard-card">', unsafe_allow_html=True)
            st.subheader("👮 Top Tài Xế (Km)")
            top_driver = df_sub.groupby('Driver')['Km'].sum().nlargest(10).reset_index().sort_values('Km')
            fig_drv = px.bar(top_driver, x='Km', y='Driver', orientation='h', text_auto='.2s', color='Km', color_continuous_scale='Teal')
            st.plotly_chart(fig_drv, use_container_width=True)
            st.markdown('</div>', unsafe_allow_html=True)
            
        with c2:
            st.markdown('<div class="dashboard-card">', unsafe_allow_html=True)
            st.subheader("👥 Top Người Dùng (Chi phí)")
            top_user = df_sub.groupby('User')['Cost'].sum().nlargest(10).reset_index().sort_values('Cost')
            fig_user = px.bar(top_user, x='Cost', y='User', orientation='h', text_auto='.2s', color='Cost', color_continuous_scale='Purples')
            st.plotly_chart(fig_user, use_container_width=True)
            st.markdown('</div>', unsafe_allow_html=True)
            
        c3, c4 = st.columns(2)
        with c3:
            st.markdown('<div class="dashboard-card">', unsafe_allow_html=True)
            st.subheader("⛽ Top Xe Tốn Xăng Nhất")
            if 'Fuel' in df_sub.columns:
                top_fuel = df_sub.groupby('Car')['Fuel'].sum().nlargest(10).reset_index().sort_values('Fuel')
                fig_fuel = px.bar(top_fuel, x='Fuel', y='Car', orientation='h', text_auto='.2s', color='Fuel', color_continuous_scale='Reds')
                st.plotly_chart(fig_fuel, use_container_width=True)
            st.markdown('</div>', unsafe_allow_html=True)
            
        with c4:
            st.markdown('<div class="dashboard-card">', unsafe_allow_html=True)
            st.subheader("🏢 Top Bộ Phận Sử Dụng")
            top_dept = df_sub.groupby('Dept')['Cost'].sum().nlargest(10).reset_index().sort_values('Cost')
            fig_dept = px.bar(top_dept, x='Cost', y='Dept', orientation='h', text_auto='.2s', color='Cost', color_continuous_scale='Blues')
            st.plotly_chart(fig_dept, use_container_width=True)
            st.markdown('</div>', unsafe_allow_html=True)

    # === TAB 4: SELF-SERVICE ===
    with tab_explore:
        st.markdown('<div class="filter-box">💡 <strong>Chế độ Chuyên gia:</strong> Tự do phân tích dữ liệu theo ý muốn.</div>', unsafe_allow_html=True)
        c1, c2, c3, c4 = st.columns(4)
        with c1: chart_type = st.selectbox("Loại Biểu Đồ", ["Bar", "Line", "Pie", "Scatter", "H-Bar"])
        with c2: 
            dim_map = {'Dept': 'Bộ Phận', 'Driver': 'Tài Xế', 'Car': 'Xe', 'Tháng': 'Tháng', 'User': 'User', 'Route_Type': 'Lộ Trình', 'Company': 'Công Ty'}
            valid_dims = [k for k in dim_map.keys() if k in df_sub.columns]
            x_axis = st.selectbox("Trục X", valid_dims, format_func=lambda x: dim_map[x])
        with c3: 
            met_map = {'Cost': 'Chi Phí', 'Km': 'Km', 'Fuel': 'Xăng', 'Duration_Hours': 'Giờ Chạy'}
            y_axis = st.selectbox("Trục Y", [k for k in met_map.keys() if k in df_sub.columns], format_func=lambda x: met_map[x])
        with c4: color_by = st.selectbox("Màu Sắc", ["None"] + [k for k in valid_dims if k != x_axis])

        grp = [x_axis]
        if color_by != "None": grp.append(color_by)
        df_chart = df_sub.groupby(grp, as_index=False)[y_axis].sum()
        
        if chart_type == "Bar": fig = px.bar(df_chart, x=x_axis, y=y_axis, color=color_by if color_by!="None" else None, title=f"{met_map[y_axis]} theo {dim_map[x_axis]}")
        elif chart_type == "H-Bar": fig = px.bar(df_chart, x=y_axis, y=x_axis, orientation='h', color=color_by if color_by!="None" else None, title=f"{met_map[y_axis]} theo {dim_map[x_axis]}")
        elif chart_type == "Pie": fig = px.pie(df_chart, values=y_axis, names=x_axis, title=f"Tỷ lệ {met_map[y_axis]}")
        elif chart_type == "Line": fig = px.line(df_chart, x=x_axis, y=y_axis, markers=True, color=color_by if color_by!="None" else None)
        elif chart_type == "Scatter": fig = px.scatter(df_chart, x=x_axis, y=y_axis, size=y_axis, color=color_by if color_by!="None" else None)
        
        st.plotly_chart(fig, use_container_width=True)

else:
    st.info("👋 Hãy tải file Excel lên để bắt đầu.")