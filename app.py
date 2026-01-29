import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import numpy as np

# --- 1. CẤU HÌNH TRANG ---
st.set_page_config(
    page_title="Báo Cáo Đội Xe",
    page_icon="🚘",
    layout="wide",
    initial_sidebar_state="expanded"
)

# CSS: Giao diện sạch, Font chữ to rõ
st.markdown("""
<style>
    .stApp { background-color: #f8f9fa; }
    .metric-card {
        background: white; border-radius: 10px; padding: 15px;
        box-shadow: 0 2px 5px rgba(0,0,0,0.05); border-top: 4px solid #3B82F6;
        text-align: center;
    }
    .metric-val { font-size: 26px; font-weight: bold; color: #1e293b; margin: 5px 0; }
    .metric-lbl { font-size: 14px; color: #64748b; text-transform: uppercase; }
    /* Tabs đẹp hơn */
    .stTabs [data-baseweb="tab-list"] { background: white; padding: 10px; border-radius: 10px; }
    .stTabs [aria-selected="true"] { color: #2563eb !important; border-bottom-color: #2563eb !important; }
</style>
""", unsafe_allow_html=True)

# --- 2. XỬ LÝ DỮ LIỆU ---
@st.cache_data
def load_data(file):
    try:
        if file.name.endswith('.csv'):
            df = pd.read_csv(file, header=3)
        else:
            xl = pd.ExcelFile(file)
            target = next((s for s in xl.sheet_names if "booking" in s.lower()), xl.sheet_names[0])
            df = pd.read_excel(file, sheet_name=target, header=3)

        # Chuẩn hóa tên cột
        df.columns = [str(c).strip().replace('\n', ' ') for c in df.columns]
        
        # Map cột sang tiếng Anh để code dễ hơn
        col_map = {
            'Ngày Tháng Năm': 'Date', 'Biển số xe': 'Car', 'Tên tài xế': 'Driver',
            'Bộ phận': 'Dept', 'Cost center': 'CostCenter', 'Km sử dụng': 'Km',
            'Tổng chi phí': 'Cost', 'Lộ trình': 'Route', 'Giờ khởi hành': 'Start',
            'Chi phí nhiên liệu': 'Fuel', 'Phí cầu đường': 'Toll', 'Sửa chữa': 'Repair'
        }
        df = df.rename(columns={k:v for k,v in col_map.items() if k in df.columns})
        
        # Xử lý dữ liệu
        df.dropna(how='all', inplace=True)
        if 'Date' in df.columns:
            df['Date'] = pd.to_datetime(df['Date'], errors='coerce')
            df = df.dropna(subset=['Date'])
            df['Tháng'] = df['Date'].dt.strftime('%m-%Y')
        
        # Chuyển số
        for c in ['Km', 'Cost', 'Fuel', 'Toll', 'Repair']:
            if c in df.columns: df[c] = pd.to_numeric(df[c], errors='coerce').fillna(0)
            
        # Làm sạch Text
        for c in ['Dept', 'Driver', 'Car']:
            if c in df.columns: df[c] = df[c].astype(str).str.strip()

        return df
    except: return pd.DataFrame()

# --- 3. GIAO DIỆN CHÍNH ---
st.title("🚘 Báo Cáo Quản Trị Đội Xe")

uploaded_file = st.sidebar.file_uploader("Tải file Excel vào đây", type=['xlsx', 'csv'])

if uploaded_file:
    df = load_data(uploaded_file)
    if not df.empty:
        # --- BỘ LỌC ---
        st.sidebar.markdown("---")
        st.sidebar.header("🔍 Bộ Lọc")
        months = sorted(df['Tháng'].unique())
        sel_month = st.sidebar.multiselect("Chọn Tháng", months, default=months)
        
        depts = sorted(df['Dept'].unique())
        sel_dept = st.sidebar.multiselect("Chọn Bộ Phận", depts, default=depts)
        
        # Áp dụng lọc
        mask = df['Tháng'].isin(sel_month) & df['Dept'].isin(sel_dept)
        df_sub = df[mask]
        
        if df_sub.empty: st.warning("Không có dữ liệu!"); st.stop()

        # --- KPI CARDS (Đơn giản hóa) ---
        c1, c2, c3, c4 = st.columns(4)
        with c1: st.markdown(f'<div class="metric-card"><div class="metric-lbl">Tổng Chi Phí</div><div class="metric-val">{df_sub["Cost"].sum():,.0f}</div></div>', unsafe_allow_html=True)
        with c2: st.markdown(f'<div class="metric-card"><div class="metric-lbl">Tổng Km</div><div class="metric-val">{df_sub["Km"].sum():,.0f}</div></div>', unsafe_allow_html=True)
        with c3: st.markdown(f'<div class="metric-card"><div class="metric-lbl">Số Chuyến</div><div class="metric-val">{len(df_sub):,}</div></div>', unsafe_allow_html=True)
        avg = df_sub["Cost"].sum()/df_sub["Km"].sum() if df_sub["Km"].sum()>0 else 0
        with c4: st.markdown(f'<div class="metric-card"><div class="metric-lbl">Giá / Km</div><div class="metric-val">{avg:,.0f}</div></div>', unsafe_allow_html=True)
        
        st.write("") # Spacer

        # --- TABS ---
        tab1, tab2 = st.tabs(["📊 Báo Cáo Trực Quan (Dễ hiểu)", "📄 Dữ Liệu Chi Tiết"])

        with tab1:
            st.info("💡 Mẹo: Chọn loại biểu đồ và dữ liệu bên dưới để hệ thống tự vẽ.")
            
            # --- MENU CHỌN BIỂU ĐỒ (SIMPLE VERSION) ---
            col_type, col_x, col_y = st.columns(3)
            
            with col_type:
                # Dùng từ ngữ thông dụng
                chart_type = st.selectbox("1. Bạn muốn xem kiểu gì?", 
                                        ["So Sánh (Cột Đứng)", "Xếp Hạng (Cột Ngang)", "Cơ Cấu (Bánh Donut)", "Xu Hướng (Đường)"])
            
            with col_x:
                # Map tên cột sang tiếng Việt cho user dễ hiểu
                dim_map = {'Dept': 'Bộ Phận', 'Driver': 'Tài Xế', 'Car': 'Biển Số Xe', 'Tháng': 'Tháng', 'CostCenter': 'Cost Center'}
                # Chỉ lấy cột có trong df
                valid_dims = [k for k in dim_map.keys() if k in df_sub.columns]
                dim_choice = st.selectbox("2. Phân tích theo nhóm nào?", valid_dims, format_func=lambda x: dim_map[x])
            
            with col_y:
                metric_map = {'Cost': 'Tổng Chi Phí (VNĐ)', 'Km': 'Số Km Đã Chạy', 'Fuel': 'Tiền Xăng', 'Toll': 'Phí Cầu Đường'}
                valid_metrics = [k for k in metric_map.keys() if k in df_sub.columns]
                metric_choice = st.selectbox("3. Xem số liệu gì?", valid_metrics, format_func=lambda x: metric_map[x])

            # --- XỬ LÝ & VẼ BIỂU ĐỒ ---
            st.markdown("---")
            
            # Group by
            df_chart = df_sub.groupby(dim_choice, as_index=False)[metric_choice].sum()
            
            # Auto Sort (Sắp xếp từ cao xuống thấp cho dễ nhìn)
            if chart_type in ["So Sánh (Cột Đứng)", "Xếp Hạng (Cột Ngang)"]:
                df_chart = df_chart.sort_values(metric_choice, ascending=False)
            
            # Title
            chart_title = f"Biểu đồ {metric_map[metric_choice]} theo {dim_map[dim_choice]}"

            # Logic vẽ từng loại (Đơn giản hóa tối đa)
            if chart_type == "So Sánh (Cột Đứng)":
                fig = px.bar(df_chart, x=dim_choice, y=metric_choice, 
                             text_auto='.2s', # Hiện số rút gọn (vd: 1.5M)
                             title=chart_title, color=metric_choice, color_continuous_scale='Blues')
                fig.update_layout(xaxis_title=dim_map[dim_choice], yaxis_title="")
                
            elif chart_type == "Xếp Hạng (Cột Ngang)":
                # Thích hợp cho Top Tài xế, Top Bộ phận
                fig = px.bar(df_chart.head(15), x=metric_choice, y=dim_choice, orientation='h', # Top 15 thôi cho đỡ rối
                             text_auto='.2s', 
                             title=f"Top 15 {dim_map[dim_choice]} cao nhất", 
                             color=metric_choice, color_continuous_scale='Teal')
                fig.update_layout(yaxis={'categoryorder':'total ascending'}, xaxis_title="", yaxis_title="")
                
            elif chart_type == "Cơ Cấu (Bánh Donut)":
                fig = px.pie(df_chart, names=dim_choice, values=metric_choice, hole=0.5,
                             title=chart_title)
                fig.update_traces(textposition='inside', textinfo='percent+label')
                
            elif chart_type == "Xu Hướng (Đường)":
                # Nếu xem xu hướng thì nên sort theo thời gian (nếu chọn Tháng)
                if dim_choice == 'Tháng':
                    df_chart = df_chart.sort_values('Tháng') 
                fig = px.line(df_chart, x=dim_choice, y=metric_choice, markers=True,
                              title=chart_title)
                fig.update_traces(line_color='#e11d48', line_width=3)

            # Tinh chỉnh chung cho đẹp
            fig.update_layout(height=500, font=dict(size=14))
            st.plotly_chart(fig, use_container_width=True)
            
            # Show bảng số liệu nhỏ bên dưới cho ai cần đối chiếu
            with st.expander("Xem bảng số liệu chi tiết"):
                st.dataframe(df_chart.style.format({metric_choice: "{:,.0f}"}))

        with tab2:
            st.dataframe(df_sub)
else:
    st.info("👋 Hãy tải file Excel lên để bắt đầu.")