import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import numpy as np

# --- 1. CẤU HÌNH TRANG ---
st.set_page_config(
    page_title="Hệ Thống Quản Trị Đội Xe",
    page_icon="🚘",
    layout="wide",
    initial_sidebar_state="expanded"
)

# CSS Styling: Giao diện Clean & Professional
st.markdown("""
<style>
    /* Tổng thể */
    .stApp { background-color: #f4f6f8; }
    
    /* KPI Cards */
    .metric-card {
        background: white; border-radius: 12px; padding: 20px;
        box-shadow: 0 4px 6px rgba(0,0,0,0.05); border-left: 5px solid #3B82F6;
        transition: transform 0.2s;
    }
    .metric-card:hover { transform: translateY(-5px); }
    .metric-title { font-size: 13px; color: #64748b; font-weight: 700; text-transform: uppercase; letter-spacing: 0.5px; }
    .metric-value { font-size: 28px; font-weight: 800; color: #0f172a; margin: 8px 0; }
    .metric-unit { font-size: 12px; color: #94a3b8; font-weight: 500; }
    
    /* Tabs & Charts */
    .stTabs [data-baseweb="tab-list"] { background: white; padding: 10px 20px; border-radius: 10px; gap: 20px; }
    .stTabs [aria-selected="true"] { color: #2563eb; border-bottom: 2px solid #2563eb; font-weight: bold; }
    .chart-container { background: white; padding: 20px; border-radius: 12px; box-shadow: 0 2px 4px rgba(0,0,0,0.02); margin-bottom: 20px; }
</style>
""", unsafe_allow_html=True)

# --- 2. XỬ LÝ DỮ LIỆU ---
@st.cache_data
def load_and_process_data(file):
    try:
        # A. Đọc file
        if file.name.endswith('.csv'):
            df = pd.read_csv(file, header=3)
        else:
            xl = pd.ExcelFile(file)
            target = next((s for s in xl.sheet_names if "booking" in s.lower()), xl.sheet_names[0])
            df = pd.read_excel(file, sheet_name=target, header=3)

        # B. Chuẩn hóa tên cột
        df.columns = [str(c).strip().replace('\n', ' ') for c in df.columns]
        
        # C. Mapping cột (Từ Tiếng Việt file gốc -> Tiếng Anh Code)
        col_map = {
            'Ngày Tháng Năm': 'Date', 'Biển số xe': 'Car', 'Tên tài xế': 'Driver',
            'Bộ phận': 'Dept', 'Cost center': 'CostCenter', 'Km sử dụng': 'Km',
            'Tổng chi phí': 'Cost', 'Lộ trình': 'Route', 'Giờ khởi hành': 'Start_Time',
            # Chi phí thành phần
            'Chi phí nhiên liệu': 'Cost_Fuel', 'Phí cầu đường': 'Cost_Toll',
            'VETC': 'Cost_VETC', 'Sửa chữa': 'Cost_Repair', 'Tiền cơm': 'Cost_Meal'
        }
        
        # Chỉ lấy cột có trong file
        cols = [c for c in col_map.keys() if c in df.columns]
        df = df[cols].rename(columns=col_map)
        
        # D. Xử lý dữ liệu
        df.dropna(how='all', inplace=True)
        if 'Date' in df.columns:
            df['Date'] = pd.to_datetime(df['Date'], errors='coerce')
            df = df.dropna(subset=['Date'])
            df['Tháng'] = df['Date'].dt.strftime('%m-%Y')
            df['YearMonth'] = df['Date'].dt.to_period('M') # Để sort
        
        # Chuyển số
        num_cols = ['Km', 'Cost', 'Cost_Fuel', 'Cost_Toll', 'Cost_VETC', 'Cost_Repair', 'Cost_Meal']
        for c in num_cols:
            if c in df.columns: df[c] = pd.to_numeric(df[c], errors='coerce').fillna(0)
            
        # Tính chi phí Khác (Other)
        known_cols = [c for c in num_cols if c in df.columns and c not in ['Cost', 'Km']]
        if known_cols:
            df['Cost_Other'] = df['Cost'] - df[known_cols].sum(axis=1)
            df['Cost_Other'] = df['Cost_Other'].apply(lambda x: x if x > 0 else 0)

        # Phân loại Lộ trình (Nội tỉnh / Ngoại tỉnh)
        if 'Route' in df.columns:
            df['Route'] = df['Route'].astype(str).fillna("")
            df['Route_Type'] = df['Route'].apply(lambda x: 'Nội Tỉnh' if len(str(x)) < 5 or any(k in str(x).lower() for k in ['hcm', 'sài gòn', 'q1', 'q7', 'city']) else 'Ngoại Tỉnh')

        # Làm sạch chuỗi
        for c in ['Dept', 'Driver', 'Car', 'CostCenter']:
            if c in df.columns: df[c] = df[c].astype(str).str.strip()
            
        return df
    except Exception as e:
        return pd.DataFrame()

def kpi_card(title, value, unit, color="#3B82F6"):
    st.markdown(f"""
    <div class="metric-card" style="border-left-color: {color}">
        <div class="metric-title">{title}</div>
        <div class="metric-value">{value}</div>
        <div class="metric-unit">{unit}</div>
    </div>
    """, unsafe_allow_html=True)

# --- 3. GIAO DIỆN CHÍNH ---
st.title("🚘 Dashboard Quản Trị Đội Xe")

# Sidebar Upload & Filter
with st.sidebar:
    st.header("📂 Dữ Liệu & Bộ Lọc")
    uploaded_file = st.file_uploader("Upload Excel/CSV", type=['xlsx', 'csv'])
    
    df = pd.DataFrame()
    if uploaded_file:
        df = load_and_process_data(uploaded_file)
        
    if not df.empty:
        st.divider()
        st.subheader("Bộ Lọc")
        # Sort tháng đúng thứ tự thời gian
        if 'YearMonth' in df.columns:
            sorted_months = df.sort_values('YearMonth')['Tháng'].unique()
        else:
            sorted_months = sorted(df['Tháng'].unique())
            
        sel_month = st.multiselect("Tháng", sorted_months, default=sorted_months)
        sel_dept = st.multiselect("Bộ Phận", sorted(df['Dept'].unique()), default=sorted(df['Dept'].unique()))
        
        # Filter Logic
        mask = df['Tháng'].isin(sel_month) & df['Dept'].isin(sel_dept)
        df_sub = df[mask]
    else:
        df_sub = pd.DataFrame()

# Main Content
if not df_sub.empty:
    # --- KPI Overview ---
    c1, c2, c3, c4 = st.columns(4)
    cost = df_sub['Cost'].sum()
    km = df_sub['Km'].sum()
    with c1: kpi_card("Tổng Chi Phí", f"{cost:,.0f}", "VNĐ", "#ef4444")
    with c2: kpi_card("Tổng Km", f"{km:,.0f}", "Km", "#3b82f6")
    with c3: kpi_card("Số Chuyến", f"{len(df_sub):,}", "Chuyến", "#10b981")
    avg = cost/km if km > 0 else 0
    with c4: kpi_card("Chi Phí / Km", f"{avg:,.0f}", "VNĐ/Km", "#f59e0b")
    
    st.write("") # Spacer

    # --- TABS ---
    tab_overview, tab_explore, tab_data = st.tabs(["📊 Tổng Quan (Dashboard)", "🛠️ Tự Phân Tích (Explorer)", "📄 Dữ Liệu Chi Tiết"])

    # === TAB 1: DASHBOARD TỔNG QUAN (FIXED CHARTS) ===
    with tab_overview:
        # Row 1: Xu hướng & Cơ cấu
        c1, c2 = st.columns([2, 1])
        with c1:
            st.markdown('<div class="chart-container">', unsafe_allow_html=True)
            st.subheader("📈 Xu Hướng Chi Phí & Km Theo Thời Gian")
            trend = df_sub.groupby('Date')[['Cost', 'Km']].sum().reset_index()
            fig_trend = go.Figure()
            fig_trend.add_trace(go.Bar(x=trend['Date'], y=trend['Cost'], name='Chi Phí', marker_color='#60A5FA'))
            fig_trend.add_trace(go.Scatter(x=trend['Date'], y=trend['Km'], name='Km', yaxis='y2', line=dict(color='#F87171', width=3)))
            fig_trend.update_layout(yaxis2=dict(overlaying='y', side='right', showgrid=False), margin=dict(t=10, b=10, l=10, r=10), height=350)
            st.plotly_chart(fig_trend, use_container_width=True)
            st.markdown('</div>', unsafe_allow_html=True)
        
        with c2:
            st.markdown('<div class="chart-container">', unsafe_allow_html=True)
            st.subheader("🍩 Cơ Cấu Chi Phí")
            cost_cols = [c for c in ['Cost_Fuel', 'Cost_Toll', 'Cost_VETC', 'Cost_Repair', 'Cost_Other'] if c in df_sub.columns]
            cost_sum = df_sub[cost_cols].sum().reset_index()
            cost_sum.columns = ['Loại', 'Giá Trị']
            fig_pie = px.pie(cost_sum, values='Giá Trị', names='Loại', hole=0.5, color_discrete_sequence=px.colors.qualitative.Pastel)
            fig_pie.update_layout(margin=dict(t=10, b=10, l=10, r=10), height=350)
            st.plotly_chart(fig_pie, use_container_width=True)
            st.markdown('</div>', unsafe_allow_html=True)

        # Row 2: Top Xếp hạng
        c3, c4 = st.columns(2)
        with c3:
            st.markdown('<div class="chart-container">', unsafe_allow_html=True)
            st.subheader("🏢 Top Bộ Phận Sử Dụng (Chi Phí)")
            top_dept = df_sub.groupby('Dept')['Cost'].sum().nlargest(10).reset_index().sort_values('Cost')
            fig_dept = px.bar(top_dept, x='Cost', y='Dept', orientation='h', text_auto='.2s', color='Cost', color_continuous_scale='Blues')
            fig_dept.update_layout(height=400, margin=dict(t=10, b=10))
            st.plotly_chart(fig_dept, use_container_width=True)
            st.markdown('</div>', unsafe_allow_html=True)
            
        with c4:
            st.markdown('<div class="chart-container">', unsafe_allow_html=True)
            st.subheader("👮 Top Tài Xế (Km Vận Hành)")
            top_driver = df_sub.groupby('Driver')['Km'].sum().nlargest(10).reset_index().sort_values('Km')
            fig_driver = px.bar(top_driver, x='Km', y='Driver', orientation='h', text_auto='.2s', color='Km', color_continuous_scale='Greens')
            fig_driver.update_layout(height=400, margin=dict(t=10, b=10))
            st.plotly_chart(fig_driver, use_container_width=True)
            st.markdown('</div>', unsafe_allow_html=True)

    # === TAB 2: TỰ DO PHÂN TÍCH (FLEXIBLE) ===
    with tab_explore:
        st.info("💡 Hướng dẫn: Chọn các tiêu chí bên dưới để tự tạo biểu đồ theo ý muốn.")
        
        # Controls Row
        col_type, col_x, col_y, col_color = st.columns(4)
        
        with col_type:
            chart_type = st.selectbox("1. Loại Biểu Đồ", 
                                    ["Cột (Bar)", "Đường (Line)", "Vùng (Area)", "Bánh (Pie)", "Phân Tán (Scatter)", "Cột Ngang (H-Bar)"])
        
        with col_x:
            # Map tên cột hiển thị cho đẹp
            dim_map = {'Dept': 'Bộ Phận', 'Driver': 'Tài Xế', 'Car': 'Xe', 'Tháng': 'Tháng', 
                       'CostCenter': 'Cost Center', 'Route_Type': 'Loại Lộ Trình'}
            valid_dims = [k for k in dim_map.keys() if k in df_sub.columns]
            x_axis = st.selectbox("2. Trục X (Phân nhóm)", valid_dims, format_func=lambda x: dim_map[x])
            
        with col_y:
            metric_map = {'Cost': 'Tổng Chi Phí', 'Km': 'Số Km', 'Cost_Fuel': 'Tiền Xăng', 'Cost_Repair': 'Sửa Chữa'}
            valid_metrics = [k for k in metric_map.keys() if k in df_sub.columns]
            y_axis = st.selectbox("3. Trục Y (Giá trị)", valid_metrics, format_func=lambda x: metric_map[x])
            
        with col_color:
            color_opts = ["None"] + [k for k in valid_dims if k != x_axis] # Tránh trùng
            color_by = st.selectbox("4. Phân Màu (Tùy chọn)", color_opts, format_func=lambda x: dim_map.get(x, x))

        # Xử lý dữ liệu vẽ
        st.markdown("---")
        
        grp_cols = [x_axis]
        if color_by != "None": grp_cols.append(color_by)
        
        # Group & Sum
        df_chart = df_sub.groupby(grp_cols, as_index=False)[y_axis].sum()
        
        # Vẽ biểu đồ
        title = f"Biểu đồ {metric_map[y_axis]} theo {dim_map[x_axis]}"
        
        if chart_type == "Cột (Bar)":
            fig = px.bar(df_chart, x=x_axis, y=y_axis, color=color_by if color_by!="None" else None, 
                         text_auto='.2s', title=title)
        elif chart_type == "Cột Ngang (H-Bar)":
             # Sort để nhìn đẹp hơn
            df_chart = df_chart.sort_values(y_axis, ascending=True)
            fig = px.bar(df_chart, x=y_axis, y=x_axis, color=color_by if color_by!="None" else None, 
                         orientation='h', text_auto='.2s', title=title)
        elif chart_type == "Đường (Line)":
            fig = px.line(df_chart, x=x_axis, y=y_axis, color=color_by if color_by!="None" else None, 
                          markers=True, title=title)
        elif chart_type == "Vùng (Area)":
            fig = px.area(df_chart, x=x_axis, y=y_axis, color=color_by if color_by!="None" else None, title=title)
        elif chart_type == "Bánh (Pie)":
            fig = px.pie(df_chart, names=x_axis, values=y_axis, title=title)
        elif chart_type == "Phân Tán (Scatter)":
            fig = px.scatter(df_chart, x=x_axis, y=y_axis, color=color_by if color_by!="None" else None, 
                             size=y_axis, title=title)

        st.plotly_chart(fig, use_container_width=True)
        
        with st.expander("Xem bảng số liệu"):
            st.dataframe(df_chart)

    # === TAB 3: DATA ===
    with tab_data:
        st.dataframe(df_sub.style.format({"Cost": "{:,.0f}", "Km": "{:,.0f}"}))

else:
    st.info("👋 Vui lòng tải file Excel lên để bắt đầu.")