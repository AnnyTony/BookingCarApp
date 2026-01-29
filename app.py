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

# --- CSS CUSTOM: 3D SIDEBAR & CARD UI ---
st.markdown("""
<style>
    /* Tổng thể Background */
    .stApp { background-color: #f4f7f6; }
    
    /* === 3D SIDEBAR STYLE === */
    [data-testid="stSidebar"] {
        background: linear-gradient(180deg, #ffffff 0%, #f0f2f6 100%);
        box-shadow: 2px 0 15px rgba(0,0,0,0.05);
        border-right: 1px solid #e0e0e0;
    }
    
    /* Filter Box Container */
    .filter-container {
        background-color: white;
        padding: 20px;
        border-radius: 15px;
        box-shadow: 0 10px 25px rgba(0,0,0,0.08); /* Hiệu ứng nổi 3D */
        border: 1px solid #ffffff;
        margin-bottom: 20px;
    }
    
    .filter-header {
        font-size: 16px; font-weight: 700; color: #2c3e50;
        margin-bottom: 15px; border-bottom: 2px solid #3498db;
        padding-bottom: 5px; display: inline-block;
    }

    /* === DASHBOARD CARDS === */
    .dashboard-card {
        background-color: white; padding: 20px; border-radius: 15px;
        box-shadow: 0 4px 15px rgba(0,0,0,0.05); margin-bottom: 20px;
        transition: transform 0.2s;
        border: 1px solid #f1f3f4;
    }
    .dashboard-card:hover {
        transform: translateY(-3px); /* Hiệu ứng nhấc lên khi hover */
        box-shadow: 0 8px 20px rgba(0,0,0,0.1);
    }
    
    /* === KPI METRIC BOX === */
    .kpi-box {
        background: white; padding: 20px; border-radius: 15px;
        border-left: 5px solid #3b82f6;
        box-shadow: 0 5px 15px rgba(0,0,0,0.05);
        text-align: left;
    }
    .kpi-label { font-size: 13px; color: #7f8c8d; font-weight: 700; text-transform: uppercase; letter-spacing: 0.5px; }
    .kpi-val { font-size: 28px; font-weight: 800; color: #2c3e50; margin-top: 5px; }
    .kpi-sub { font-size: 12px; color: #95a5a6; font-weight: 500; }

    /* Tabs Styling */
    .stTabs [data-baseweb="tab-list"] { background: white; padding: 10px 20px; border-radius: 50px; gap: 10px; box-shadow: 0 2px 5px rgba(0,0,0,0.02); }
    .stTabs [aria-selected="true"] { color: #3498db !important; border-bottom-color: #3498db !important; font-weight: bold; }
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
            
        # Phân loại Lộ Trình
        if 'Route' in df.columns:
            df['Route'] = df['Route'].astype(str).fillna("")
            def get_route_type(s):
                s = s.lower()
                if len(s) < 5 or any(k in s for k in ['hcm', 'sài gòn', 'q1', 'q7', 'thủ đức', 'city']): return 'Nội Tỉnh'
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

# --- SIDEBAR 3D ---
with st.sidebar:
    st.markdown('<div class="filter-header">📂 DỮ LIỆU ĐẦU VÀO</div>', unsafe_allow_html=True)
    uploaded_file = st.file_uploader("", type=['xlsx', 'csv'])
    
    df = pd.DataFrame()
    if uploaded_file:
        df = load_data(uploaded_file)

    if not df.empty:
        st.write("")
        # Container tạo khối 3D cho bộ lọc
        st.markdown('<div class="filter-container">', unsafe_allow_html=True)
        st.markdown('<div class="filter-header">🔍 BỘ LỌC DỮ LIỆU</div>', unsafe_allow_html=True)
        
        # Sort tháng
        if 'SortMonth' in df.columns:
            months = sorted(df['Tháng'].unique(), key=lambda x: df[df['Tháng']==x]['SortMonth'].iloc[0])
        else:
            months = sorted(df['Tháng'].unique())
            
        sel_month = st.multiselect("Chọn Tháng", months, default=months)
        
        depts = sorted(df['Dept'].dropna().unique())
        sel_dept = st.multiselect("Chọn Bộ Phận / BU", depts, default=depts)
        
        st.markdown('</div>', unsafe_allow_html=True) # End filter container
        
        # Logic lọc
        mask = pd.Series(True, index=df.index)
        if sel_month: mask &= df['Tháng'].isin(sel_month)
        if sel_dept: mask &= df['Dept'].isin(sel_dept)
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
    with c2: draw_kpi("Tổng Km Vận Hành", f"{km:,.0f}", "Km", "#3b82f6")
    with c3: draw_kpi("Chi Phí Nhiên Liệu", f"{fuel:,.0f}", "VNĐ", "#f59e0b")
    avg = cost/km if km > 0 else 0
    with c4: draw_kpi("Hiệu Suất (Cost/Km)", f"{avg:,.0f}", "VNĐ/Km", "#10b981")

    st.write("")

    # --- MAIN TABS ---
    tab1, tab2, tab3, tab4 = st.tabs(["📊 Tổng Quan & Xu Hướng", "🏆 Bảng Xếp Hạng & Hiệu Suất", "🛠️ Tự Do Phân Tích", "📄 Dữ Liệu Chi Tiết"])

    # === TAB 1: TỔNG QUAN ===
    with tab1:
        # ROW 1: XU HƯỚNG
        st.markdown("### 📈 Phân Tích Xu Hướng (Trend)")
        st.markdown('<div class="dashboard-card">', unsafe_allow_html=True)
        daily = df_sub.groupby('Date')[['Cost', 'Km']].sum().reset_index()
        
        fig_trend = make_subplots(rows=2, cols=1, shared_xaxes=True, vertical_spacing=0.1,
                                  subplot_titles=("Dòng Tiền Chi Phí (VNĐ)", "Quãng Đường Vận Hành (Km)"))
        
        fig_trend.add_trace(go.Scatter(x=daily['Date'], y=daily['Cost'], fill='tozeroy', 
                                       name='Chi Phí', line=dict(color='#ef4444', width=2)), row=1, col=1)
        fig_trend.add_trace(go.Scatter(x=daily['Date'], y=daily['Km'], fill='tozeroy', 
                                       name='Km', line=dict(color='#3b82f6', width=2)), row=2, col=1)
        
        fig_trend.update_layout(height=450, showlegend=False, margin=dict(t=30, b=10, l=10, r=10), template="plotly_white")
        st.plotly_chart(fig_trend, use_container_width=True)
        st.markdown('</div>', unsafe_allow_html=True)
        
        # ROW 2: PIE CHARTS
        c_left, c_right = st.columns(2)
        with c_left:
            st.markdown('<div class="dashboard-card">', unsafe_allow_html=True)
            st.markdown("#### 🍩 Tỷ Lệ Lộ Trình (Nội vs Ngoại Tỉnh)")
            if 'Route_Type' in df_sub.columns:
                route_stats = df_sub['Route_Type'].value_counts().reset_index()
                route_stats.columns = ['Loại', 'Số chuyến']
                fig_route = px.pie(route_stats, names='Loại', values='Số chuyến', hole=0.6, 
                                   color_discrete_sequence=['#2ecc71', '#f1c40f'])
                st.plotly_chart(fig_route, use_container_width=True)
            st.markdown('</div>', unsafe_allow_html=True)
            
        with c_right:
            st.markdown('<div class="dashboard-card">', unsafe_allow_html=True)
            st.markdown("#### 🎯 Phân Tích Tương Quan (Scatter)")
            car_perf = df_sub.groupby('Car')[['Cost', 'Km']].sum().reset_index()
            car_perf_clean = car_perf[car_perf['Km'] > 0].copy() # Filter km > 0
            
            if not car_perf_clean.empty:
                car_perf_clean['AVG'] = car_perf_clean['Cost'] / car_perf_clean['Km']
                fig_scatter = px.scatter(car_perf_clean, x='Km', y='Cost', color='Car', size='Km',
                                         hover_data={'AVG': ':.0f', 'Cost': ':.0f', 'Km': ':.0f'})
                st.plotly_chart(fig_scatter, use_container_width=True)
            else:
                st.info("Chưa đủ dữ liệu xe hoạt động > 0km")
            st.markdown('</div>', unsafe_allow_html=True)

    # === TAB 2: TOP & HIỆU SUẤT (MỚI) ===
    with tab2:
        st.markdown("### 🏆 Xếp Hạng & Hiệu Suất")
        
        # MỚI: BIỂU ĐỒ HIỆU SUẤT XE
        st.markdown('<div class="dashboard-card">', unsafe_allow_html=True)
        st.subheader("🏎️ Xếp Hạng Hiệu Suất (Chi Phí / Km)")
        st.caption("Xe có cột càng dài = Chi phí vận hành trên 1km càng cao (Kém hiệu quả hoặc xe đặc chủng)")
        
        car_eff = df_sub.groupby('Car')[['Cost', 'Km']].sum().reset_index()
        car_eff = car_eff[car_eff['Km'] > 100] # Chỉ tính xe chạy trên 100km để tránh sai số
        car_eff['Efficiency'] = car_eff['Cost'] / car_eff['Km']
        car_eff = car_eff.sort_values('Efficiency', ascending=False).head(15) # Top 15 tốn kém nhất
        
        if not car_eff.empty:
            fig_eff = px.bar(car_eff, x='Efficiency', y='Car', orientation='h', text_auto='.0f',
                             color='Efficiency', color_continuous_scale='Redor',
                             labels={'Efficiency': 'VNĐ/Km'})
            fig_eff.update_layout(yaxis={'categoryorder':'total ascending'}, height=400)
            st.plotly_chart(fig_eff, use_container_width=True)
        else:
            st.warning("Chưa đủ dữ liệu (Km > 100) để tính hiệu suất.")
        st.markdown('</div>', unsafe_allow_html=True)

        col_top1, col_top2 = st.columns(2)
        
        # Top User
        with col_top1:
            st.markdown('<div class="dashboard-card">', unsafe_allow_html=True)
            st.subheader("👥 Top Người Dùng (Chi Phí)")
            if 'User' in df_sub.columns:
                top_user = df_sub.groupby('User')['Cost'].sum().nlargest(10).reset_index().sort_values('Cost')
                fig_user = px.bar(top_user, x='Cost', y='User', orientation='h', text_auto='.2s', color='Cost', color_continuous_scale='Purples')
                st.plotly_chart(fig_user, use_container_width=True)
            st.markdown('</div>', unsafe_allow_html=True)

        # Top Km
        with col_top2:
            st.markdown('<div class="dashboard-card">', unsafe_allow_html=True)
            st.subheader("🚗 Top Xe Hoạt Động (Km)")
            top_km = df_sub.groupby('Car')['Km'].sum().nlargest(10).reset_index().sort_values('Km')
            top_km = top_km[top_km['Km']>0]
            fig_km = px.bar(top_km, x='Km', y='Car', orientation='h', text_auto='.2s', color='Km', color_continuous_scale='Teal')
            st.plotly_chart(fig_km, use_container_width=True)
            st.markdown('</div>', unsafe_allow_html=True)

    # === TAB 3: SELF-SERVICE ===
    with tab3:
        st.markdown("""<div style="background-color:#e3f2fd; padding:15px; border-radius:10px; margin-bottom:20px; border:1px solid #90caf9;">
            <strong>💡 Chế độ Chuyên Gia:</strong> Tự do lựa chọn biểu đồ để phân tích sâu.</div>""", unsafe_allow_html=True)
        
        c1, c2, c3, c4 = st.columns(4)
        with c1: chart_type = st.selectbox("1. Loại Biểu Đồ", ["Cột (Bar)", "Đường (Line)", "Vùng (Area)", "Bánh (Pie)", "Phân Tán (Scatter)", "Cột Ngang (H-Bar)"])
        with c2: 
            dim_map = {'Dept': 'Bộ Phận', 'Driver': 'Tài Xế', 'Car': 'Xe', 'Tháng': 'Tháng', 'CostCenter': 'Cost Center', 'Route_Type': 'Lộ Trình', 'User': 'Người Dùng'}
            valid_dims = [k for k in dim_map.keys() if k in df_sub.columns]
            x_axis = st.selectbox("2. Trục X (Phân nhóm)", valid_dims, format_func=lambda x: dim_map[x])
        with c3: 
            metric_map = {'Cost': 'Tổng Chi Phí', 'Km': 'Số Km', 'Fuel': 'Tiền Xăng', 'Toll': 'Phí Cầu Đường'}
            valid_metrics = [k for k in metric_map.keys() if k in df_sub.columns]
            y_axis = st.selectbox("3. Trục Y (Giá trị)", valid_metrics, format_func=lambda x: metric_map[x])
        with c4: 
            color_opts = ["None"] + [k for k in valid_dims if k != x_axis]
            color_by = st.selectbox("4. Phân Màu (Tùy chọn)", color_opts, format_func=lambda x: dim_map.get(x, x))

        st.markdown("---")
        grp_cols = [x_axis]
        if color_by != "None": grp_cols.append(color_by)
        
        df_chart = df_sub.groupby(grp_cols, as_index=False)[y_axis].sum()
        title = f"{metric_map[y_axis]} theo {dim_map[x_axis]}"
        
        if chart_type == "Cột (Bar)": fig = px.bar(df_chart, x=x_axis, y=y_axis, color=color_by if color_by!="None" else None, text_auto='.2s', title=title)
        elif chart_type == "Cột Ngang (H-Bar)": fig = px.bar(df_chart.sort_values(y_axis), x=y_axis, y=x_axis, color=color_by if color_by!="None" else None, orientation='h', text_auto='.2s', title=title)
        elif chart_type == "Đường (Line)": fig = px.line(df_chart, x=x_axis, y=y_axis, color=color_by if color_by!="None" else None, markers=True, title=title)
        elif chart_type == "Vùng (Area)": fig = px.area(df_chart, x=x_axis, y=y_axis, color=color_by if color_by!="None" else None, title=title)
        elif chart_type == "Bánh (Pie)": fig = px.pie(df_chart, names=x_axis, values=y_axis, title=title)
        elif chart_type == "Phân Tán (Scatter)": fig = px.scatter(df_chart, x=x_axis, y=y_axis, color=color_by if color_by!="None" else None, size=y_axis, title=title)

        st.plotly_chart(fig, use_container_width=True)
        with st.expander("Xem số liệu"): st.dataframe(df_chart)

    # === TAB 4: DATA ===
    with tab4:
        st.dataframe(df_sub.style.format({"Cost": "{:,.0f}", "Km": "{:,.0f}", "Fuel": "{:,.0f}"}))

else:
    st.info("👋 Hãy tải file Excel (Data-SuDungXe) để bắt đầu.")