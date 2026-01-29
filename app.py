import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
from plotly.subplots import make_subplots

# --- 1. CẤU HÌNH TRANG ---
st.set_page_config(
    page_title="Báo Cáo Đội Xe",
    page_icon="🚘",
    layout="wide",
    initial_sidebar_state="expanded"
)

# --- CSS: GIAO DIỆN PHẲNG, DỄ NHÌN ---
st.markdown("""
<style>
    /* Nền sáng sủa */
    .stApp { background-color: #f8f9fa; }
    
    /* Sidebar đơn giản */
    [data-testid="stSidebar"] {
        background-color: white;
        border-right: 1px solid #dee2e6;
    }
    
    /* Card (Khung chứa) */
    .simple-card {
        background-color: white;
        padding: 20px;
        border-radius: 8px;
        border: 1px solid #e9ecef;
        box-shadow: 0 2px 4px rgba(0,0,0,0.02);
        margin-bottom: 20px;
    }
    
    /* KPI Box - To rõ */
    .kpi-container {
        background-color: white;
        padding: 15px;
        border-radius: 8px;
        border-left: 5px solid #0d6efd; /* Màu xanh chuẩn */
        box-shadow: 0 2px 4px rgba(0,0,0,0.05);
    }
    .kpi-label { font-size: 14px; color: #6c757d; text-transform: uppercase; font-weight: 600; }
    .kpi-value { font-size: 28px; color: #212529; font-weight: bold; margin: 5px 0; }
    .kpi-note { font-size: 12px; color: #198754; } /* Màu xanh lá */

    /* Tabs */
    .stTabs [data-baseweb="tab-list"] { background: white; padding: 5px; border-radius: 8px; }
    .stTabs [aria-selected="true"] { color: #0d6efd; font-weight: bold; border-bottom: 2px solid #0d6efd; }
</style>
""", unsafe_allow_html=True)

# --- 2. XỬ LÝ DỮ LIỆU (FIX LỖI KM ÂM) ---
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
        
        # Map cột
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
            
        # --- FIX QUAN TRỌNG: Lọc bỏ Km âm hoặc quá lớn (do lỗi nhập liệu) ---
        if 'Km' in df.columns:
            # Chỉ lấy các chuyến có Km > 0 và < 5000 (tránh số ảo 200,000km)
            df = df[(df['Km'] > 0) & (df['Km'] < 5000)]
            
        # Xử lý Lộ Trình
        if 'Route' in df.columns:
            df['Route'] = df['Route'].astype(str).fillna("")
            df['Route_Type'] = df['Route'].apply(lambda s: 'Nội Tỉnh' if len(str(s)) < 5 or any(k in str(s).lower() for k in ['hcm', 'sài gòn', 'q1', 'q7', 'city']) else 'Ngoại Tỉnh')
        
        # Tính thời gian chạy
        if 'Start_Time' in df.columns and 'End_Time' in df.columns:
            def calc_duration(row):
                try:
                    s = pd.to_datetime(str(row['Start_Time']), format='%H:%M:%S', errors='coerce')
                    e = pd.to_datetime(str(row['End_Time']), format='%H:%M:%S', errors='coerce')
                    if pd.notnull(s) and pd.notnull(e):
                        diff = (e - s).total_seconds() / 3600
                        return diff if diff > 0 else 0
                    return 0
                except: return 0
            df['Duration_Hours'] = df.apply(calc_duration, axis=1)
        else:
            df['Duration_Hours'] = 0

        # Làm sạch Text
        for c in ['Dept', 'Driver', 'Car', 'Company', 'User']:
            if c in df.columns: df[c] = df[c].astype(str).str.strip()
            
        return df
    except Exception as e:
        return pd.DataFrame()

# --- 3. UI COMPONENTS ---
def kpi_card(title, val, unit, color="#0d6efd"):
    st.markdown(f"""
    <div class="kpi-container" style="border-left-color: {color}">
        <div class="kpi-label">{title}</div>
        <div class="kpi-value">{val}</div>
        <div class="kpi-note">{unit}</div>
    </div>
    """, unsafe_allow_html=True)

# --- 4. MAIN APP ---
st.title("🚘 Báo Cáo Quản Trị Đội Xe")
st.caption("Dữ liệu được làm sạch và hiển thị tối giản")

# --- SIDEBAR ---
with st.sidebar:
    st.header("📂 Dữ Liệu")
    uploaded_file = st.file_uploader("Tải file Excel/CSV", type=['xlsx', 'csv'])
    
    df = pd.DataFrame()
    if uploaded_file: df = load_data(uploaded_file)

    if not df.empty:
        st.markdown("---")
        st.subheader("🔍 Bộ Lọc")
        
        # Sort months
        if 'SortMonth' in df.columns:
            months = sorted(df['Tháng'].unique(), key=lambda x: df[df['Tháng']==x]['SortMonth'].iloc[0])
        else: months = sorted(df['Tháng'].unique())
            
        sel_month = st.multiselect("Tháng", months, default=months)
        sel_dept = st.multiselect("Bộ Phận", sorted(df['Dept'].unique()), default=sorted(df['Dept'].unique()))
        
        # Filter Logic
        mask = pd.Series(True, index=df.index)
        if sel_month: mask &= df['Tháng'].isin(sel_month)
        if sel_dept: mask &= df['Dept'].isin(sel_dept)
        df_sub = df[mask]
    else: df_sub = pd.DataFrame()

if not df_sub.empty:
    # --- KPI SUMMARY ---
    c1, c2, c3, c4 = st.columns(4)
    total_cost = df_sub['Cost'].sum()
    total_km = df_sub['Km'].sum() # Đã fix lỗi âm
    total_trips = len(df_sub)
    avg_cost = total_cost / total_km if total_km > 0 else 0

    with c1: kpi_card("Tổng Chi Phí", f"{total_cost:,.0f}", "VNĐ", "#dc3545") # Đỏ
    with c2: kpi_card("Tổng Số Km", f"{total_km:,.0f}", "Km", "#0d6efd") # Xanh
    with c3: kpi_card("Tổng Số Chuyến", f"{total_trips:,}", "Chuyến", "#198754") # Lục
    with c4: kpi_card("Trung Bình", f"{avg_cost:,.0f}", "VNĐ/Km", "#ffc107") # Vàng

    st.markdown("<br>", unsafe_allow_html=True)
    
    # --- TABS ---
    tab_overview, tab_perf, tab_rank, tab_explore = st.tabs([
        "📊 Tổng Quan", 
        "⚡ Hiệu Suất", 
        "🏆 Xếp Hạng", 
        "🛠️ Tự Phân Tích"
    ])

    # === TAB 1: TỔNG QUAN ===
    with tab_overview:
        col_L, col_R = st.columns([2, 1])
        
        with col_L:
            st.markdown('<div class="simple-card">', unsafe_allow_html=True)
            st.subheader("📈 Xu Hướng Theo Thời Gian")
            daily = df_sub.groupby('Date')[['Cost', 'Km']].sum().reset_index()
            
            # Combo Chart
            fig_trend = make_subplots(specs=[[{"secondary_y": True}]])
            fig_trend.add_trace(go.Bar(x=daily['Date'], y=daily['Cost'], name="Chi Phí (VNĐ)", 
                                       marker_color='#aacbff', opacity=0.8), secondary_y=False)
            fig_trend.add_trace(go.Scatter(x=daily['Date'], y=daily['Km'], name="Km Vận Hành", 
                                           line=dict(color='#0d6efd', width=3)), secondary_y=True)
            
            fig_trend.update_layout(height=400, hovermode='x unified', showlegend=True, 
                                    template='plotly_white', margin=dict(t=10, b=10))
            st.plotly_chart(fig_trend, use_container_width=True)
            st.markdown('</div>', unsafe_allow_html=True)

        with col_R:
            st.markdown('<div class="simple-card">', unsafe_allow_html=True)
            st.subheader("🏢 Phân Bổ Theo Công Ty")
            if 'Company' in df_sub.columns:
                comp_stats = df_sub['Company'].value_counts().reset_index()
                comp_stats.columns = ['Công Ty', 'Số Chuyến']
                fig_comp = px.pie(comp_stats, values='Số Chuyến', names='Công Ty', 
                                  hole=0.5, color_discrete_sequence=px.colors.qualitative.Pastel)
                fig_comp.update_layout(height=400, margin=dict(t=10, b=10))
                st.plotly_chart(fig_comp, use_container_width=True)
            else: st.info("Không có dữ liệu Công ty")
            st.markdown('</div>', unsafe_allow_html=True)
        
        # --- NEW: Bảng dữ liệu chi tiết ---
        with st.expander("📄 Xem chi tiết dữ liệu (Danh sách chuyến xe)"):
            st.dataframe(df_sub.style.format({"Cost": "{:,.0f}", "Km": "{:,.0f}"}), use_container_width=True)

    # === TAB 2: HIỆU SUẤT ===
    with tab_perf:
        st.info("💡 Hiệu suất giúp bạn biết xe nào hoạt động hiệu quả, xe nào 'ngồi chơi xơi nước'.")
        
        c1, c2 = st.columns(2)
        
        # 1. Công suất
        with c1:
            st.markdown('<div class="simple-card">', unsafe_allow_html=True)
            st.subheader("📊 Tỷ Lệ Xe Hoạt Động (% Ngày)")
            total_cars = df['Car'].nunique()
            daily_active = df_sub.groupby('Date')['Car'].nunique().reset_index()
            daily_active['Pct'] = (daily_active['Car'] / total_cars) * 100
            
            fig_util = px.bar(daily_active, x='Date', y='Pct', labels={'Pct': '% Xe hoạt động'}, 
                              title="Ngày nào xe đi nhiều nhất?", color_discrete_sequence=['#198754'])
            fig_util.update_layout(height=350, template='plotly_white')
            st.plotly_chart(fig_util, use_container_width=True)
            st.markdown('</div>', unsafe_allow_html=True)
            
        # 2. Scatter Plot
        with c2:
            st.markdown('<div class="simple-card">', unsafe_allow_html=True)
            st.subheader("🎯 Tương Quan: Chi Phí vs Quãng Đường")
            car_perf = df_sub.groupby('Car')[['Cost', 'Km']].sum().reset_index()
            car_perf = car_perf[car_perf['Km'] > 0]
            
            fig_sc = px.scatter(car_perf, x='Km', y='Cost', size='Km', color='Car',
                                labels={'Km': 'Quãng đường (Km)', 'Cost': 'Tổng tiền (VNĐ)'},
                                title="Bóng to = Xe chạy nhiều")
            st.plotly_chart(fig_sc, use_container_width=True)
            st.caption("Gợi ý: Các chấm nằm góc trên bên trái là xe tốn tiền nhưng đi ít.")
            st.markdown('</div>', unsafe_allow_html=True)
            
        # --- NEW: Bảng dữ liệu hiệu suất ---
        with st.expander("📄 Xem bảng tổng hợp hiệu suất xe"):
            car_perf['Avg_Cost_Km'] = car_perf['Cost'] / car_perf['Km']
            st.dataframe(car_perf.style.format({
                "Cost": "{:,.0f}", 
                "Km": "{:,.0f}", 
                "Avg_Cost_Km": "{:,.0f}"
            }), use_container_width=True)

    # === TAB 3: XẾP HẠNG ===
    with tab_rank:
        c1, c2 = st.columns(2)
        with c1:
            st.markdown('<div class="simple-card">', unsafe_allow_html=True)
            st.subheader("🏆 Top Tài Xế (Km)")
            top_driver = df_sub.groupby('Driver')['Km'].sum().nlargest(10).reset_index().sort_values('Km')
            fig = px.bar(top_driver, x='Km', y='Driver', orientation='h', text_auto='.0f', color_discrete_sequence=['#0dcaf0'])
            st.plotly_chart(fig, use_container_width=True)
            st.markdown('</div>', unsafe_allow_html=True)
            
        with c2:
            st.markdown('<div class="simple-card">', unsafe_allow_html=True)
            st.subheader("👥 Top Người Dùng (Chi Phí)")
            top_user = df_sub.groupby('User')['Cost'].sum().nlargest(10).reset_index().sort_values('Cost')
            fig = px.bar(top_user, x='Cost', y='User', orientation='h', text_auto='.2s', color_discrete_sequence=['#6f42c1'])
            st.plotly_chart(fig, use_container_width=True)
            st.markdown('</div>', unsafe_allow_html=True)
            
        # --- NEW: Bảng xếp hạng chi tiết ---
        with st.expander("📄 Xem danh sách xếp hạng chi tiết"):
            col_a, col_b = st.columns(2)
            with col_a:
                st.write("**Top Tài Xế**")
                st.dataframe(top_driver.style.format({"Km": "{:,.0f}"}), use_container_width=True)
            with col_b:
                st.write("**Top Người Dùng**")
                st.dataframe(top_user.style.format({"Cost": "{:,.0f}"}), use_container_width=True)

    # === TAB 4: TỰ PHÂN TÍCH ===
    with tab_explore:
        st.markdown('<div class="simple-card">', unsafe_allow_html=True)
        st.subheader("🛠️ Công Cụ Tự Tạo Biểu Đồ")
        st.caption("Chọn thông tin bạn muốn xem, hệ thống sẽ tự vẽ.")
        
        c1, c2, c3, c4 = st.columns(4)
        with c1: chart_type = st.selectbox("1. Kiểu biểu đồ", ["Cột", "Đường", "Bánh", "Cột Ngang"])
        with c2: 
            dim_map = {'Dept': 'Bộ Phận', 'Driver': 'Tài Xế', 'Car': 'Xe', 'Tháng': 'Tháng', 'Company': 'Công Ty'}
            valid_dims = [k for k in dim_map.keys() if k in df_sub.columns]
            x_axis = st.selectbox("2. Nhóm theo", valid_dims, format_func=lambda x: dim_map[x])
        with c3: 
            met_map = {'Cost': 'Chi Phí', 'Km': 'Số Km', 'Fuel': 'Tiền Xăng'}
            y_axis = st.selectbox("3. Số liệu", [k for k in met_map.keys() if k in df_sub.columns], format_func=lambda x: met_map[x])
        with c4: color_by = st.selectbox("4. Màu sắc (Tùy chọn)", ["None"] + [k for k in valid_dims if k != x_axis])

        grp = [x_axis]
        if color_by != "None": grp.append(color_by)
        df_chart = df_sub.groupby(grp, as_index=False)[y_axis].sum()
        
        title = f"{met_map[y_axis]} theo {dim_map[x_axis]}"
        if chart_type == "Cột": fig = px.bar(df_chart, x=x_axis, y=y_axis, color=color_by if color_by!="None" else None, title=title)
        elif chart_type == "Cột Ngang": fig = px.bar(df_chart.sort_values(y_axis), x=y_axis, y=x_axis, orientation='h', title=title)
        elif chart_type == "Bánh": fig = px.pie(df_chart, values=y_axis, names=x_axis, title=title)
        elif chart_type == "Đường": fig = px.line(df_chart, x=x_axis, y=y_axis, markers=True, title=title)
        
        st.plotly_chart(fig, use_container_width=True)
        
        # --- NEW: Bảng dữ liệu tự phân tích ---
        st.write("---")
        st.write("#### 📄 Dữ liệu chi tiết cho biểu đồ trên:")
        st.dataframe(df_chart.style.format({y_axis: "{:,.0f}"}), use_container_width=True)
        
        st.markdown('</div>', unsafe_allow_html=True)

else:
    st.info("👋 Vui lòng tải file Excel lên để bắt đầu.")