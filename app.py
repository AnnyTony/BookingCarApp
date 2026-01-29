import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
from plotly.subplots import make_subplots

# --- 1. CẤU HÌNH TRANG ---
st.set_page_config(
    page_title="Dashboard Đội Xe Toàn Diện",
    page_icon="🚘",
    layout="wide",
    initial_sidebar_state="expanded"
)

# CSS Custom: Giao diện thẻ nổi (Card UI)
st.markdown("""
<style>
    .stApp { background-color: #f0f2f5; }
    
    /* Card Styles */
    .dashboard-card {
        background-color: white; padding: 20px; border-radius: 12px;
        box-shadow: 0 2px 5px rgba(0,0,0,0.05); margin-bottom: 20px;
    }
    
    /* KPI Metric */
    .kpi-box {
        background: white; padding: 15px; border-radius: 10px;
        border-left: 4px solid #3b82f6;
        box-shadow: 0 2px 4px rgba(0,0,0,0.05);
    }
    .kpi-label { font-size: 13px; color: #64748b; font-weight: 600; text-transform: uppercase; }
    .kpi-val { font-size: 24px; font-weight: 800; color: #1e293b; margin-top: 5px; }
    .kpi-sub { font-size: 12px; color: #94a3b8; }

    /* Tabs */
    .stTabs [data-baseweb="tab-list"] { background: white; padding: 10px; border-radius: 10px; gap: 10px; }
    .stTabs [aria-selected="true"] { color: #2563eb; border-bottom: 2px solid #2563eb; font-weight: bold; }
</style>
""", unsafe_allow_html=True)

# --- 2. HÀM XỬ LÝ DỮ LIỆU ---
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

        # Chuẩn hóa cột
        df.columns = [str(c).strip().replace('\n', ' ') for c in df.columns]
        
        # Mapping cột
        col_map = {
            'Ngày Tháng Năm': 'Date', 'Biển số xe': 'Car', 'Tên tài xế': 'Driver',
            'Bộ phận': 'Dept', 'Cost center': 'CostCenter', 'Km sử dụng': 'Km',
            'Tổng chi phí': 'Cost', 'Lộ trình': 'Route', 'Người sử dụng xe': 'User',
            'Chi phí nhiên liệu': 'Fuel', 'Phí cầu đường': 'Toll', 'Sửa chữa': 'Repair'
        }
        cols = [c for c in col_map.keys() if c in df.columns]
        df = df[cols].rename(columns=col_map)
        
        # Xử lý dữ liệu
        df.dropna(how='all', inplace=True)
        if 'Date' in df.columns:
            df['Date'] = pd.to_datetime(df['Date'], errors='coerce')
            df = df.dropna(subset=['Date'])
            df['Tháng'] = df['Date'].dt.strftime('%m-%Y')
            df['SortMonth'] = df['Date'].dt.to_period('M')
        
        # Chuyển số
        for c in ['Km', 'Cost', 'Fuel', 'Toll', 'Repair']:
            if c in df.columns: df[c] = pd.to_numeric(df[c], errors='coerce').fillna(0)
            
        # Phân loại Lộ Trình (Logic: chứa từ khóa HCM/SG -> Nội tỉnh)
        if 'Route' in df.columns:
            df['Route'] = df['Route'].astype(str).fillna("")
            def get_route_type(s):
                s = s.lower()
                if len(s) < 5 or any(k in s for k in ['hcm', 'sài gòn', 'q1', 'q7', 'thủ đức', 'city']):
                    return 'Nội Tỉnh'
                return 'Ngoại Tỉnh'
            df['Route_Type'] = df['Route'].apply(get_route_type)
        else:
            df['Route_Type'] = 'Khác'
            
        return df
    except: return pd.DataFrame()

# --- 3. UI COMPONENTS ---
def draw_kpi(title, val, unit, color):
    st.markdown(f"""
    <div class="kpi-box" style="border-left-color: {color}">
        <div class="kpi-label">{title}</div>
        <div class="kpi-val">{val}</div>
        <div class="kpi-sub">{unit}</div>
    </div>
    """, unsafe_allow_html=True)

# --- 4. MAIN APP ---
st.title("🚀 Dashboard Quản Trị Đội Xe")

with st.sidebar:
    st.header("📂 Dữ liệu & Bộ lọc")
    uploaded_file = st.file_uploader("Upload Excel/CSV", type=['xlsx', 'csv'])
    
    df = pd.DataFrame()
    if uploaded_file:
        df = load_data(uploaded_file)

    if not df.empty:
        st.divider()
        months = sorted(df['Tháng'].unique())
        sel_month = st.multiselect("Tháng", months, default=months)
        
        depts = sorted(df['Dept'].dropna().unique())
        sel_dept = st.multiselect("Bộ Phận", depts, default=depts)
        
        mask = df['Tháng'].isin(sel_month) & df['Dept'].isin(sel_dept)
        df_sub = df[mask]
    else:
        df_sub = pd.DataFrame()

if not df_sub.empty:
    # --- KPI ROW ---
    c1, c2, c3, c4 = st.columns(4)
    cost = df_sub['Cost'].sum()
    km = df_sub['Km'].sum()
    fuel = df_sub['Fuel'].sum() if 'Fuel' in df_sub.columns else 0
    with c1: draw_kpi("Tổng Chi Phí", f"{cost:,.0f}", "VNĐ", "#ef4444")
    with c2: draw_kpi("Tổng Km", f"{km:,.0f}", "Km", "#3b82f6")
    with c3: draw_kpi("Chi Phí Nhiên Liệu", f"{fuel:,.0f}", "VNĐ", "#f59e0b")
    avg = cost/km if km>0 else 0
    with c4: draw_kpi("Chi Phí / Km", f"{avg:,.0f}", "VNĐ/Km", "#10b981")

    st.write("")

    # --- MAIN TABS ---
    tab1, tab2, tab3 = st.tabs(["📊 Tổng Quan & Xu Hướng", "🏆 Bảng Xếp Hạng (Top)", "📄 Dữ Liệu Chi Tiết"])

    # === TAB 1: TỔNG QUAN ===
    with tab1:
        # ROW 1: XU HƯỚNG (Split Charts)
        st.markdown("### 📈 Xu Hướng Hoạt Động (Tách Biệt)")
        st.markdown('<div class="dashboard-card">', unsafe_allow_html=True)
        
        daily = df_sub.groupby('Date')[['Cost', 'Km']].sum().reset_index()
        
        # Vẽ biểu đồ 2 dòng (Subplots)
        fig_trend = make_subplots(rows=2, cols=1, shared_xaxes=True, 
                                  vertical_spacing=0.1,
                                  subplot_titles=("Xu Hướng Chi Phí (VNĐ)", "Xu Hướng Quãng Đường (Km)"))
        
        # Chart 1: Chi phí (Area)
        fig_trend.add_trace(go.Scatter(x=daily['Date'], y=daily['Cost'], fill='tozeroy', 
                                       name='Chi Phí', line=dict(color='#ef4444', width=2)), row=1, col=1)
        
        # Chart 2: Km (Area)
        fig_trend.add_trace(go.Scatter(x=daily['Date'], y=daily['Km'], fill='tozeroy', 
                                       name='Km', line=dict(color='#3b82f6', width=2)), row=2, col=1)
        
        fig_trend.update_layout(height=500, showlegend=False, margin=dict(t=30, b=10, l=10, r=10))
        st.plotly_chart(fig_trend, use_container_width=True)
        st.markdown('</div>', unsafe_allow_html=True)
        
        # ROW 2: HIỆU SUẤT & LỘ TRÌNH
        c_left, c_right = st.columns(2)
        
        with c_left:
            st.markdown('<div class="dashboard-card">', unsafe_allow_html=True)
            st.markdown("#### 🍩 Tỷ Lệ Lộ Trình (Nội vs Ngoại Tỉnh)")
            route_stats = df_sub['Route_Type'].value_counts().reset_index()
            route_stats.columns = ['Loại', 'Số chuyến']
            fig_route = px.pie(route_stats, names='Loại', values='Số chuyến', hole=0.6, 
                               color_discrete_sequence=['#10b981', '#f59e0b'])
            st.plotly_chart(fig_route, use_container_width=True)
            st.markdown('</div>', unsafe_allow_html=True)
            
        with c_right:
            st.markdown('<div class="dashboard-card">', unsafe_allow_html=True)
            st.markdown("#### 🎯 Phân Tích Hiệu Suất Xe (Scatter)")
            # Gom theo xe
            car_perf = df_sub.groupby('Car')[['Cost', 'Km']].sum().reset_index()
            # Tính Cost/Km
            car_perf['AVG'] = car_perf['Cost'] / car_perf['Km']
            
            fig_scatter = px.scatter(car_perf, x='Km', y='Cost', color='Car', size='Km',
                                     title="Tương quan: Đi nhiều (Phải) vs Tốn tiền (Trên)",
                                     hover_data=['AVG'])
            # Kẻ đường trung bình
            fig_scatter.add_shape(type="line", x0=0, y0=0, x1=car_perf['Km'].max(), y1=car_perf['Cost'].max(),
                                  line=dict(color="Gray", dash="dash"))
            st.plotly_chart(fig_scatter, use_container_width=True)
            st.markdown('</div>', unsafe_allow_html=True)

    # === TAB 2: TOP CHARTS ===
    with tab2:
        st.markdown("### 🏆 Bảng Xếp Hạng Tiêu Biểu")
        
        col_top1, col_top2 = st.columns(2)
        
        # 1. Top Người Dùng
        with col_top1:
            st.markdown('<div class="dashboard-card">', unsafe_allow_html=True)
            st.subheader("👥 Top Người Dùng (Book xe nhiều nhất)")
            if 'User' in df_sub.columns:
                top_user = df_sub.groupby('User')['Cost'].sum().nlargest(10).reset_index().sort_values('Cost')
                fig_user = px.bar(top_user, x='Cost', y='User', orientation='h', text_auto='.2s', 
                                  color='Cost', color_continuous_scale='Purples')
                st.plotly_chart(fig_user, use_container_width=True)
            else:
                st.warning("Không tìm thấy cột 'Người sử dụng xe'")
            st.markdown('</div>', unsafe_allow_html=True)

        # 2. Top Xe Ngốn Xăng
        with col_top2:
            st.markdown('<div class="dashboard-card">', unsafe_allow_html=True)
            st.subheader("⛽ Top Xe Tiêu Thụ Nhiên Liệu (VNĐ)")
            if 'Fuel' in df_sub.columns:
                top_fuel = df_sub.groupby('Car')['Fuel'].sum().nlargest(10).reset_index().sort_values('Fuel')
                fig_fuel = px.bar(top_fuel, x='Fuel', y='Car', orientation='h', text_auto='.2s', 
                                  color='Fuel', color_continuous_scale='Reds')
                st.plotly_chart(fig_fuel, use_container_width=True)
            else:
                st.warning("Không có dữ liệu nhiên liệu")
            st.markdown('</div>', unsafe_allow_html=True)

        col_top3, col_top4 = st.columns(2)
        
        # 3. Top Xe Chạy Nhiều
        with col_top3:
            st.markdown('<div class="dashboard-card">', unsafe_allow_html=True)
            st.subheader("🚗 Top Xe Hoạt Động (Km)")
            top_km = df_sub.groupby('Car')['Km'].sum().nlargest(10).reset_index().sort_values('Km')
            fig_km = px.bar(top_km, x='Km', y='Car', orientation='h', text_auto='.2s', 
                              color='Km', color_continuous_scale='Teal')
            st.plotly_chart(fig_km, use_container_width=True)
            st.markdown('</div>', unsafe_allow_html=True)
            
        # 4. Top Bộ Phận
        with col_top4:
            st.markdown('<div class="dashboard-card">', unsafe_allow_html=True)
            st.subheader("🏢 Top Bộ Phận (Chi phí)")
            top_dept = df_sub.groupby('Dept')['Cost'].sum().nlargest(10).reset_index().sort_values('Cost')
            fig_dept = px.bar(top_dept, x='Cost', y='Dept', orientation='h', text_auto='.2s', 
                              color='Cost', color_continuous_scale='Blues')
            st.plotly_chart(fig_dept, use_container_width=True)
            st.markdown('</div>', unsafe_allow_html=True)

    # === TAB 3: DATA ===
    with tab3:
        st.dataframe(df_sub.style.format({"Cost": "{:,.0f}", "Km": "{:,.0f}", "Fuel": "{:,.0f}"}))

else:
    st.info("👋 Hãy tải file Excel (Data-SuDungXe) để bắt đầu.")