import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import numpy as np

# --- 1. CẤU HÌNH TRANG & THEME ---
st.set_page_config(
    page_title="Pro Fleet Analytics",
    page_icon="🚘",
    layout="wide",
    initial_sidebar_state="expanded"
)

# CSS Pro: Giao diện sạch, hiện đại, Cards nổi
st.markdown("""
<style>
    /* Global Background */
    .stApp { background-color: #f4f6f9; }
    
    /* KPI Cards Design */
    .metric-card {
        background: white;
        border-radius: 12px;
        padding: 20px;
        box-shadow: 0 4px 10px rgba(0,0,0,0.05);
        border-left: 5px solid #3B82F6;
        transition: transform 0.2s;
    }
    .metric-card:hover { transform: translateY(-5px); }
    .metric-title { font-size: 14px; color: #64748b; font-weight: 600; text-transform: uppercase; }
    .metric-value { font-size: 32px; font-weight: 800; color: #1e293b; margin: 10px 0; }
    .metric-sub { font-size: 12px; color: #94a3b8; }
    
    /* Custom Tabs */
    .stTabs [data-baseweb="tab-list"] {
        background-color: white;
        padding: 10px;
        border-radius: 10px;
        box-shadow: 0 2px 5px rgba(0,0,0,0.05);
        gap: 20px;
    }
    .stTabs [data-baseweb="tab"] { border: none; font-weight: 600; }
    .stTabs [aria-selected="true"] { color: #3B82F6; border-bottom: 2px solid #3B82F6; }

    /* Chart Container */
    .chart-box {
        background: white; padding: 20px; border-radius: 12px;
        box-shadow: 0 2px 5px rgba(0,0,0,0.05); margin-bottom: 20px;
    }
</style>
""", unsafe_allow_html=True)

# --- 2. HÀM XỬ LÝ DỮ LIỆU ROBUST ---
@st.cache_data
def load_and_process_data(file):
    try:
        # A. Đọc File Thông Minh
        if file.name.endswith('.csv'):
            df = pd.read_csv(file, header=3)
        else:
            xl = pd.ExcelFile(file)
            # Tìm sheet Booking Car, nếu không thấy thì lấy sheet đầu
            target_sheet = next((s for s in xl.sheet_names if "booking" in s.lower() and "car" in s.lower()), xl.sheet_names[0])
            df = pd.read_excel(file, sheet_name=target_sheet, header=3)

        # B. Chuẩn hóa & Map Cột
        df.columns = [str(c).strip().replace('\n', ' ') for c in df.columns]
        
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
            # Các chi phí thành phần
            'Chi phí nhiên liệu': 'Cost_Fuel',
            'Phí cầu đường': 'Cost_Toll',
            'VETC': 'Cost_VETC',
            'Sửa chữa': 'Cost_Repair',
            'Bảo dưỡng': 'Cost_Maintenance',
            'Tiền cơm': 'Cost_Meal'
        }
        
        cols_present = [c for c in col_map.keys() if c in df.columns]
        df = df[cols_present].rename(columns=col_map)
        df.dropna(how='all', inplace=True)

        if 'Date' in df.columns:
            df['Date'] = pd.to_datetime(df['Date'], errors='coerce')
            df = df.dropna(subset=['Date'])
            df['Month_Str'] = df['Date'].dt.strftime('%m-%Y')
            df['Day_Name'] = df['Date'].dt.day_name()

        # C. Xử lý số liệu & Text
        num_cols = ['Km_Used', 'Total_Cost', 'Cost_Fuel', 'Cost_Toll', 'Cost_VETC', 'Cost_Repair', 'Cost_Maintenance', 'Cost_Meal']
        for col in num_cols:
            if col in df.columns:
                df[col] = pd.to_numeric(df[col], errors='coerce').fillna(0)

        # Tạo cột Cost_Other (Các chi phí khác chưa định danh)
        known_cols = [c for c in num_cols if c in df.columns and c not in ['Total_Cost', 'Km_Used']]
        if known_cols:
            df['Cost_Other'] = df['Total_Cost'] - df[known_cols].sum(axis=1)
            df['Cost_Other'] = df['Cost_Other'].clip(lower=0)
        
        # Phân loại lộ trình (Nội tỉnh vs Ngoại tỉnh)
        if 'Route' in df.columns:
            df['Route'] = df['Route'].astype(str).fillna("")
            df['Route_Type'] = df['Route'].apply(lambda x: 'Nội Tỉnh' if len(str(x)) < 5 or any(k in str(x).lower() for k in ['hcm', 'sài gòn', 'q1', 'q7', 'city']) else 'Ngoại Tỉnh')
        
        # Làm sạch chuỗi
        for col in ['Department', 'Cost_Center', 'Car_Plate', 'Driver']:
            if col in df.columns:
                df[col] = df[col].astype(str).str.strip()

        return df
    except Exception as e:
        st.error(f"Lỗi xử lý file: {e}")
        return pd.DataFrame()

# --- 3. UI COMPONENTS ---
def draw_kpi(title, value, suffix, color="#3B82F6"):
    st.markdown(f"""
    <div class="metric-card" style="border-left-color: {color};">
        <div class="metric-title">{title}</div>
        <div class="metric-value">{value}</div>
        <div class="metric-sub">{suffix}</div>
    </div>
    """, unsafe_allow_html=True)

# --- 4. MAIN APP ---
st.title("🚀 Fleet Analytics Pro")
st.markdown("Hệ thống quản trị dữ liệu vận hành & chi phí xe (Interactive)")

# Sidebar
with st.sidebar:
    st.header("📂 Data & Filters")
    uploaded_file = st.file_uploader("Upload Excel/CSV", type=['xlsx', 'csv'])
    
    df = pd.DataFrame()
    if uploaded_file:
        df = load_and_process_data(uploaded_file)
    
    if not df.empty:
        st.markdown("---")
        st.subheader("Bộ Lọc Thông Minh")
        
        # Dynamic Filters
        months = sorted(df['Month_Str'].unique())
        sel_months = st.multiselect("Tháng", months, default=months)
        
        depts = sorted(df['Department'].unique())
        sel_depts = st.multiselect("Bộ Phận / BU", depts, default=depts)
        
        route_types = sorted(df.get('Route_Type', pd.Series(['N/A'])).unique())
        sel_routes = st.multiselect("Loại Lộ Trình", route_types, default=route_types)

        # Filter Logic
        mask = df['Month_Str'].isin(sel_months) & df['Department'].isin(sel_depts)
        if 'Route_Type' in df.columns:
            mask = mask & df['Route_Type'].isin(sel_routes)
        
        df_sub = df[mask]
    else:
        df_sub = pd.DataFrame()

# Main Content
if not df_sub.empty:
    # --- KPI ROW ---
    c1, c2, c3, c4 = st.columns(4)
    with c1: draw_kpi("Tổng Chi Phí", f"{df_sub['Total_Cost'].sum():,.0f}", "VNĐ", "#ef4444")
    with c2: draw_kpi("Tổng Km", f"{df_sub['Km_Used'].sum():,.0f}", "Km", "#3b82f6")
    with c3: draw_kpi("Số Chuyến", f"{len(df_sub):,}", "Trips", "#10b981")
    avg_cost = df_sub['Total_Cost'].sum() / df_sub['Km_Used'].sum() if df_sub['Km_Used'].sum() > 0 else 0
    with c4: draw_kpi("Chi Phí / Km", f"{avg_cost:,.0f}", "VNĐ/Km", "#f59e0b")
    
    st.markdown("<br>", unsafe_allow_html=True)

    # --- TABS ---
    tab_overview, tab_explore, tab_detail = st.tabs(["📊 Dashboard Tổng Quan", "🛠️ Tự Do Phân Tích (Self-Service)", "📄 Dữ Liệu Chi Tiết"])

    # === TAB 1: OVERVIEW (Pre-built Best Practices) ===
    with tab_overview:
        c1, c2 = st.columns(2)
        with c1:
            st.markdown('<div class="chart-box">', unsafe_allow_html=True)
            st.subheader("Xu Hướng Chi Phí Theo Thời Gian")
            daily_agg = df_sub.groupby('Date')[['Total_Cost', 'Km_Used']].sum().reset_index()
            fig_trend = go.Figure()
            fig_trend.add_trace(go.Bar(x=daily_agg['Date'], y=daily_agg['Total_Cost'], name='Chi Phí', marker_color='#60A5FA'))
            fig_trend.add_trace(go.Scatter(x=daily_agg['Date'], y=daily_agg['Km_Used'], name='Km', yaxis='y2', line=dict(color='#F87171', width=3)))
            fig_trend.update_layout(yaxis2=dict(overlaying='y', side='right'), hovermode='x unified', margin=dict(t=10, b=10, l=10, r=10))
            st.plotly_chart(fig_trend, use_container_width=True)
            st.markdown('</div>', unsafe_allow_html=True)

        with c2:
            st.markdown('<div class="chart-box">', unsafe_allow_html=True)
            st.subheader("Cơ Cấu Chi Phí (Phân rã)")
            cost_cols = [c for c in ['Cost_Fuel', 'Cost_Toll', 'Cost_VETC', 'Cost_Repair', 'Cost_Other'] if c in df_sub.columns]
            cost_sum = df_sub[cost_cols].sum().reset_index()
            cost_sum.columns = ['Loại', 'Giá Trị']
            # Dùng Pie Chart với hole (Donut)
            fig_pie = px.pie(cost_sum, values='Giá Trị', names='Loại', hole=0.5, color_discrete_sequence=px.colors.qualitative.Pastel)
            fig_pie.update_layout(margin=dict(t=10, b=10, l=10, r=10))
            st.plotly_chart(fig_pie, use_container_width=True)
            st.markdown('</div>', unsafe_allow_html=True)

        # Row 2
        c3, c4 = st.columns(2)
        with c3:
            st.markdown('<div class="chart-box">', unsafe_allow_html=True)
            st.subheader("Top Bộ Phận Sử Dụng (Chi phí)")
            dept_agg = df_sub.groupby('Department')['Total_Cost'].sum().nlargest(10).reset_index().sort_values('Total_Cost')
            fig_dept = px.bar(dept_agg, x='Total_Cost', y='Department', orientation='h', text_auto='.2s', color='Total_Cost', color_continuous_scale='Blues')
            fig_dept.update_layout(margin=dict(t=10, b=10, l=10, r=10))
            st.plotly_chart(fig_dept, use_container_width=True)
            st.markdown('</div>', unsafe_allow_html=True)

        with c4:
            st.markdown('<div class="chart-box">', unsafe_allow_html=True)
            st.subheader("Top Tài Xế Chạy Nhiều Nhất (Km)")
            driver_agg = df_sub.groupby('Driver')['Km_Used'].sum().nlargest(10).reset_index().sort_values('Km_Used')
            fig_driver = px.bar(driver_agg, x='Km_Used', y='Driver', orientation='h', text_auto='.2s', color='Km_Used', color_continuous_scale='Greens')
            fig_driver.update_layout(margin=dict(t=10, b=10, l=10, r=10))
            st.plotly_chart(fig_driver, use_container_width=True)
            st.markdown('</div>', unsafe_allow_html=True)

    # === TAB 2: SELF-SERVICE ANALYTICS (FIXED) ===
    with tab_explore:
        st.markdown("""
        <div style="background-color:#e0f2fe; padding:15px; border-radius:10px; margin-bottom:20px; border:1px solid #bae6fd;">
            <strong>💡 Hướng dẫn:</strong> Tại đây bạn đóng vai trò là chuyên gia phân tích. Hãy chọn các tiêu chí bên dưới để hệ thống tự vẽ biểu đồ theo ý bạn.
        </div>
        """, unsafe_allow_html=True)
        
        col_ctrl1, col_ctrl2, col_ctrl3, col_ctrl4 = st.columns(4)
        
        # 1. Chọn loại biểu đồ
        chart_type = col_ctrl1.selectbox("1. Chọn Loại Biểu Đồ", ["Bar Chart (Cột)", "Line Chart (Đường)", "Pie Chart (Tròn)", "Scatter (Phân tán)", "Area (Vùng)"])
        
        # 2. Chọn trục X (Phân nhóm)
        cat_cols = ['Department', 'Driver', 'Car_Plate', 'Route_Type', 'Cost_Center', 'Month_Str', 'Day_Name']
        # Chỉ lấy cột có thật trong df
        valid_cat_cols = [c for c in cat_cols if c in df_sub.columns]
        x_axis = col_ctrl2.selectbox("2. Chọn Chiều Phân Tích (Trục X)", valid_cat_cols)
        
        # 3. Chọn trục Y (Giá trị)
        val_cols = ['Total_Cost', 'Km_Used', 'Cost_Fuel', 'Cost_Repair', 'Cost_Toll']
        valid_val_cols = [c for c in val_cols if c in df_sub.columns]
        y_axis = col_ctrl3.selectbox("3. Chọn Số Liệu (Trục Y)", valid_val_cols)
        
        # 4. Màu sắc (Optional)
        color_by = col_ctrl4.selectbox("4. Phân màu theo (Tùy chọn)", ["None"] + valid_cat_cols)

        # --- XỬ LÝ VẼ BIỂU ĐỒ ĐỘNG (ĐÃ FIX LỖI TRÙNG CỘT) ---
        st.markdown("---")
        
        # Logic Fix: Nếu người dùng chọn Color trùng với X-axis thì bỏ qua Color
        actual_color = None
        group_cols = [x_axis]
        
        if color_by != "None" and color_by != x_axis:
            actual_color = color_by
            group_cols.append(color_by)
            
        # Groupby & Sum - Dùng as_index=False để tránh lỗi 'cannot insert already exists'
        df_grouped = df_sub.groupby(group_cols, as_index=False)[y_axis].sum()
        
        # Vẽ
        if chart_type == "Bar Chart (Cột)":
            fig_custom = px.bar(df_grouped, x=x_axis, y=y_axis, color=actual_color, 
                                text_auto='.2s', title=f"Biểu đồ {y_axis} theo {x_axis}")
            
        elif chart_type == "Line Chart (Đường)":
            fig_custom = px.line(df_grouped, x=x_axis, y=y_axis, color=actual_color, 
                                 markers=True, title=f"Xu hướng {y_axis} theo {x_axis}")
            
        elif chart_type == "Pie Chart (Tròn)":
            # Pie chỉ nhận 1 chiều phân tích tốt nhất
            fig_custom = px.pie(df_grouped, names=x_axis, values=y_axis, title=f"Tỷ trọng {y_axis} theo {x_axis}")
            
        elif chart_type == "Area (Vùng)":
            fig_custom = px.area(df_grouped, x=x_axis, y=y_axis, color=actual_color,
                                 title=f"Vùng giá trị {y_axis} theo {x_axis}")
                                 
        elif chart_type == "Scatter (Phân tán)":
            fig_custom = px.scatter(df_grouped, x=x_axis, y=y_axis, color=actual_color,
                                    size=y_axis, title=f"Phân tán {y_axis} theo {x_axis}")

        st.plotly_chart(fig_custom, use_container_width=True)
        
        with st.expander("Xem dữ liệu nguồn của biểu đồ này"):
            st.dataframe(df_grouped)

    # === TAB 3: DETAIL ===
    with tab_detail:
        st.subheader("Dữ liệu chi tiết")
        st.dataframe(df_sub.style.format({"Total_Cost": "{:,.0f}", "Km_Used": "{:,.0f}"}), height=600)

else:
    st.info("👋 Hãy upload file Excel 'Data-SuDungXe' để bắt đầu.")