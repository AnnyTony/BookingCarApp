import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go

# --- 1. CẤU HÌNH TRANG ---
st.set_page_config(
    page_title="Báo Cáo Đội Xe",
    page_icon="🚘",
    layout="wide",
    initial_sidebar_state="expanded"
)

# --- CSS TÙY CHỈNH (GIAO DIỆN ĐƠN GIẢN, SẠCH SẼ) ---
st.markdown("""
<style>
    /* Nền trang sáng sủa */
    .stApp { background-color: #f8f9fa; }
    
    /* Card KPI đơn giản */
    .kpi-card {
        background-color: white; border-radius: 10px; padding: 15px;
        border-top: 4px solid #007bff; /* Màu xanh cơ bản */
        box-shadow: 0 2px 4px rgba(0,0,0,0.1); text-align: center;
    }
    .kpi-title { font-size: 14px; color: #6c757d; font-weight: 600; text-transform: uppercase; }
    .kpi-value { font-size: 26px; font-weight: 800; color: #343a40; margin-top: 5px; }
    .kpi-note { font-size: 12px; color: #28a745; font-weight: 500; }

    /* Container cho biểu đồ */
    .chart-container {
        background: white; padding: 20px; border-radius: 10px;
        box-shadow: 0 2px 4px rgba(0,0,0,0.05); margin-bottom: 20px;
    }
    
    /* Tabs */
    .stTabs [data-baseweb="tab-list"] { background: white; padding: 10px; border-radius: 10px; }
    .stTabs [aria-selected="true"] { color: #007bff; border-bottom: 2px solid #007bff; font-weight: bold; }
</style>
""", unsafe_allow_html=True)

# --- 2. HÀM XỬ LÝ DỮ LIỆU (ĐÃ FIX LỖI SỐ ÂM) ---
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

        # Chuẩn hóa tên cột
        df.columns = [str(c).strip().replace('\n', ' ') for c in df.columns]
        
        col_map = {
            'Ngày Tháng Năm': 'Date', 'Biển số xe': 'Car', 'Tên tài xế': 'Driver',
            'Bộ phận': 'Dept', 'Cost center': 'CostCenter', 'Km sử dụng': 'Km',
            'Tổng chi phí': 'Cost', 'Lộ trình': 'Route', 'Người sử dụng xe': 'User',
            'Chi phí nhiên liệu': 'Fuel', 'Phí cầu đường': 'Toll', 
            'Giờ khởi hành': 'Start_Time', 'Giờ kết thúc': 'End_Time', 'Công Ty': 'Company'
        }
        cols = [c for c in col_map.keys() if c in df.columns]
        df = df[cols].rename(columns=col_map)
        
        # Xử lý Ngày Tháng
        df.dropna(how='all', inplace=True)
        if 'Date' in df.columns:
            df['Date'] = pd.to_datetime(df['Date'], errors='coerce')
            df = df.dropna(subset=['Date'])
            df['Tháng'] = df['Date'].dt.strftime('%m-%Y')
            df['SortMonth'] = df['Date'].dt.to_period('M') # Để sắp xếp tháng

        # Chuyển số liệu & LÀM SẠCH (Quan trọng)
        for c in ['Km', 'Cost', 'Fuel', 'Toll']:
            if c in df.columns: df[c] = pd.to_numeric(df[c], errors='coerce').fillna(0)
            
        # --- FIX LỖI SỐ ÂM: Chỉ lấy dòng có Km > 0 và Cost >= 0 ---
        df = df[(df['Km'] > 0) & (df['Cost'] >= 0)]

        # Tính thời gian chạy (Duration) cho biểu đồ Hiệu suất
        if 'Start_Time' in df.columns and 'End_Time' in df.columns:
            def calc_hours(row):
                try:
                    s = pd.to_datetime(str(row['Start_Time']), format='%H:%M:%S', errors='coerce')
                    e = pd.to_datetime(str(row['End_Time']), format='%H:%M:%S', errors='coerce')
                    diff = (e - s).total_seconds() / 3600
                    return diff if diff > 0 else 0
                except: return 0
            df['Hours'] = df.apply(calc_hours, axis=1)
        else:
            df['Hours'] = 0

        # Phân loại Lộ Trình đơn giản
        if 'Route' in df.columns:
            df['Route'] = df['Route'].astype(str).fillna("")
            df['Route_Type'] = df['Route'].apply(lambda s: 'Nội Tỉnh' if any(k in str(s).lower() for k in ['hcm', 'sài gòn', 'q1', 'city']) else 'Ngoại Tỉnh')

        return df
    except Exception as e:
        return pd.DataFrame()

# --- 3. GIAO DIỆN CHÍNH ---
st.title("🚘 Báo Cáo Hoạt Động Đội Xe")
st.markdown("---")

# Sidebar: Đơn giản hóa
with st.sidebar:
    st.header("📂 Dữ Liệu")
    uploaded_file = st.file_uploader("Chọn file Excel", type=['xlsx', 'csv'])
    
    df = pd.DataFrame()
    if uploaded_file: df = load_data(uploaded_file)

    if not df.empty:
        st.write("---")
        st.header("🔍 Bộ Lọc")
        
        # Sắp xếp tháng đúng thứ tự
        if 'SortMonth' in df.columns:
            months = sorted(df['Tháng'].unique(), key=lambda x: df[df['Tháng']==x]['SortMonth'].iloc[0])
        else: months = sorted(df['Tháng'].unique())
            
        sel_month = st.multiselect("Chọn Tháng", months, default=months)
        sel_dept = st.multiselect("Chọn Bộ Phận", sorted(df['Dept'].astype(str).unique()), default=sorted(df['Dept'].astype(str).unique()))
        
        # Áp dụng lọc
        mask = df['Tháng'].isin(sel_month) & df['Dept'].isin(sel_dept)
        df_sub = df[mask]
    else: df_sub = pd.DataFrame()

if not df_sub.empty:
    # --- PHẦN 1: KPI (CON SỐ QUAN TRỌNG NHẤT) ---
    c1, c2, c3, c4 = st.columns(4)
    
    total_cost = df_sub['Cost'].sum()
    total_km = df_sub['Km'].sum()
    total_trips = len(df_sub)
    cost_per_km = total_cost / total_km if total_km > 0 else 0
    
    with c1: st.markdown(f'<div class="kpi-card"><div class="kpi-title">Tổng Chi Phí</div><div class="kpi-value">{total_cost:,.0f}</div><div class="kpi-note">VNĐ</div></div>', unsafe_allow_html=True)
    with c2: st.markdown(f'<div class="kpi-card"><div class="kpi-title">Tổng Km Đã Chạy</div><div class="kpi-value">{total_km:,.0f}</div><div class="kpi-note">Km (Đã lọc số âm)</div></div>', unsafe_allow_html=True)
    with c3: st.markdown(f'<div class="kpi-card"><div class="kpi-title">Số Chuyến Xe</div><div class="kpi-value">{total_trips:,}</div><div class="kpi-note">Chuyến</div></div>', unsafe_allow_html=True)
    with c4: st.markdown(f'<div class="kpi-card"><div class="kpi-title">Trung Bình / Km</div><div class="kpi-value">{cost_per_km:,.0f}</div><div class="kpi-note">VNĐ / Km</div></div>', unsafe_allow_html=True)

    st.write("")

    # --- PHẦN 2: NỘI DUNG CHÍNH (TABS) ---
    tab_overview, tab_rank, tab_perf, tab_data = st.tabs(["📊 Tổng Quan", "🏆 Top Xếp Hạng", "⚡ Hiệu Suất Xe", "📄 Dữ Liệu Chi Tiết"])

    # === TAB 1: TỔNG QUAN ===
    with tab_overview:
        c_left, c_right = st.columns([2, 1])
        
        with c_left:
            st.markdown('<div class="chart-container">', unsafe_allow_html=True)
            st.subheader("📈 Xu Hướng: Chi Phí và Km (Theo Ngày)")
            # Dùng biểu đồ Combo đơn giản: Cột là Tiền, Dây là Km
            daily = df_sub.groupby('Date')[['Cost', 'Km']].sum().reset_index()
            
            fig_combo = go.Figure()
            fig_combo.add_trace(go.Bar(x=daily['Date'], y=daily['Cost'], name='Chi Phí (VNĐ)', marker_color='#6c757d', opacity=0.6))
            fig_combo.add_trace(go.Scatter(x=daily['Date'], y=daily['Km'], name='Số Km', yaxis='y2', line=dict(color='#007bff', width=3)))
            
            fig_combo.update_layout(
                yaxis=dict(title="VNĐ"),
                yaxis2=dict(title="Km", overlaying='y', side='right'),
                legend=dict(orientation="h", y=1.1),
                height=400, margin=dict(l=20, r=20, t=40, b=20)
            )
            st.plotly_chart(fig_combo, use_container_width=True)
            st.markdown('</div>', unsafe_allow_html=True)
            
        with c_right:
            st.markdown('<div class="chart-container">', unsafe_allow_html=True)
            st.subheader("🍩 Cơ Cấu Chi Phí")
            # Gom nhóm chi phí
            cost_data = {
                'Xăng': df_sub['Fuel'].sum(),
                'Cầu Đường': df_sub['Toll'].sum(),
                'Khác': df_sub['Cost'].sum() - df_sub['Fuel'].sum() - df_sub['Toll'].sum()
            }
            cost_df = pd.DataFrame(list(cost_data.items()), columns=['Loại', 'Tiền'])
            cost_df = cost_df[cost_df['Tiền'] > 0] # Chỉ hiện cái nào có tiền
            
            fig_pie = px.pie(cost_df, values='Tiền', names='Loại', hole=0.5, color_discrete_sequence=px.colors.qualitative.Pastel)
            fig_pie.update_traces(textposition='inside', textinfo='percent+label')
            fig_pie.update_layout(height=400, margin=dict(t=20, b=20, l=20, r=20), showlegend=False)
            st.plotly_chart(fig_pie, use_container_width=True)
            st.markdown('</div>', unsafe_allow_html=True)

    # === TAB 2: TOP XẾP HẠNG ===
    with tab_rank:
        st.info("💡 Đây là các bảng xếp hạng giúp bạn biết ai/xe nào hoạt động nhiều nhất.")
        c1, c2 = st.columns(2)
        
        with c1:
            st.markdown('<div class="chart-container">', unsafe_allow_html=True)
            st.subheader("👮 Top 10 Tài Xế (Km)")
            top_drv = df_sub.groupby('Driver')['Km'].sum().nlargest(10).reset_index().sort_values('Km')
            fig_drv = px.bar(top_drv, x='Km', y='Driver', orientation='h', text_auto='.2s', title="", color='Km', color_continuous_scale='Blues')
            st.plotly_chart(fig_drv, use_container_width=True)
            st.markdown('</div>', unsafe_allow_html=True)
            
        with c2:
            st.markdown('<div class="chart-container">', unsafe_allow_html=True)
            st.subheader("🏢 Top 10 Bộ Phận (Chi Phí)")
            top_dept = df_sub.groupby('Dept')['Cost'].sum().nlargest(10).reset_index().sort_values('Cost')
            fig_dept = px.bar(top_dept, x='Cost', y='Dept', orientation='h', text_auto='.2s', title="", color='Cost', color_continuous_scale='Reds')
            st.plotly_chart(fig_dept, use_container_width=True)
            st.markdown('</div>', unsafe_allow_html=True)

    # === TAB 3: HIỆU SUẤT (ĐƠN GIẢN HÓA) ===
    with tab_perf:
        c1, c2 = st.columns(2)
        
        with c1:
            st.markdown('<div class="chart-container">', unsafe_allow_html=True)
            st.subheader("⏳ Top Xe Bận Rộn Nhất (Giờ hoạt động)")
            st.caption("Xe nào chạy nhiều giờ nhất trong tháng?")
            
            top_busy = df_sub.groupby('Car')['Hours'].sum().nlargest(10).reset_index().sort_values('Hours')
            fig_busy = px.bar(top_busy, x='Hours', y='Car', orientation='h', text_auto='.0f', color='Hours', color_continuous_scale='Greens')
            fig_busy.update_layout(xaxis_title="Tổng Giờ Chạy", yaxis_title="")
            st.plotly_chart(fig_busy, use_container_width=True)
            st.markdown('</div>', unsafe_allow_html=True)
            
        with c2:
            st.markdown('<div class="chart-container">', unsafe_allow_html=True)
            st.subheader("📉 Công Suất Đội Xe Theo Ngày")
            st.caption("Mỗi ngày có bao nhiêu xe lăn bánh?")
            
            daily_active = df_sub.groupby('Date')['Car'].nunique().reset_index()
            fig_line = px.line(daily_active, x='Date', y='Car', markers=True, title="")
            fig_line.update_traces(line_color='#28a745', line_width=3)
            fig_line.update_layout(yaxis_title="Số lượng xe")
            st.plotly_chart(fig_line, use_container_width=True)
            st.markdown('</div>', unsafe_allow_html=True)

    # === TAB 4: DỮ LIỆU ===
    with tab_data:
        st.dataframe(df_sub.style.format({"Cost": "{:,.0f}", "Km": "{:,.0f}"}))

else:
    st.info("👋 Vui lòng tải file Excel lên để bắt đầu.")