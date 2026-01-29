import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import numpy as np

# --- 1. CẤU HÌNH TRANG ---
st.set_page_config(
    page_title="Dashboard Booking Car",
    page_icon="🚘",
    layout="wide"
)

# CSS Custom
st.markdown("""
<style>
    .kpi-card {
        background-color: #ffffff; border-radius: 10px; padding: 15px;
        box-shadow: 0 2px 4px rgba(0,0,0,0.1); text-align: center; border: 1px solid #eee;
    }
    .kpi-title { font-size: 13px; color: #666; font-weight: 600; text-transform: uppercase; }
    .kpi-value { font-size: 24px; font-weight: 800; color: #007bff; margin-top: 5px; }
    .kpi-note { font-size: 11px; color: #999; margin-top: 5px; }
</style>
""", unsafe_allow_html=True)

# --- 2. HÀM XỬ LÝ DỮ LIỆU ---
@st.cache_data
def load_data(file):
    try:
        # Bước 1: Xác định Header nằm ở đâu
        # Đọc thử 10 dòng đầu để tìm dòng chứa chữ "Ngày Tháng Năm" hoặc "Biển số xe"
        if file.name.endswith('.csv'):
            df_preview = pd.read_csv(file, nrows=10, header=None)
        else:
            # Dùng openpyxl cho file xlsx
            df_preview = pd.read_excel(file, sheet_name=0, nrows=10, header=None)
            
            # Nếu có sheet tên Booking Car thì ưu tiên đọc
            try:
                xl = pd.ExcelFile(file)
                sheet_names = xl.sheet_names
                target_sheet = next((s for s in sheet_names if "booking" in s.lower() and "car" in s.lower()), sheet_names[0])
                df_preview = pd.read_excel(file, sheet_name=target_sheet, nrows=10, header=None)
            except:
                target_sheet = 0 # Fallback

        # Tìm index dòng tiêu đề (Dòng chứa cột 'Ngày Tháng Năm' hoặc 'Date')
        header_row_idx = 3 # Mặc định theo file bạn gửi là dòng index 3 (dòng thứ 4)
        for idx, row in df_preview.iterrows():
            row_str = row.astype(str).str.lower().tolist()
            if any("ngày" in s for s in row_str) and any("biển số" in s for s in row_str):
                header_row_idx = idx
                break
        
        # Bước 2: Đọc file với header tìm được
        if file.name.endswith('.csv'):
            df = pd.read_csv(file, header=header_row_idx)
        else:
            df = pd.read_excel(file, sheet_name=target_sheet, header=header_row_idx)

        # Bước 3: Chuẩn hóa tên cột
        df.columns = [str(c).strip().replace('\n', ' ') for c in df.columns]
        
        # Mapping cột (Sử dụng tên cột tiếng Việt chính xác trong file)
        col_map = {
            'Ngày Tháng Năm': 'Date',
            'Biển số xe': 'Car_Plate',
            'Tên tài xế': 'Driver',
            'Bộ phận': 'Department',
            'Km sử dụng': 'Km_Used',
            'Tổng chi phí': 'Total_Cost',
            'Giờ khởi hành': 'Start_Time',
            'Giờ kết thúc': 'End_Time'
        }
        
        # Lọc các cột tồn tại
        available_cols = [c for c in col_map.keys() if c in df.columns]
        df = df[available_cols].rename(columns=col_map)
        
        # Bước 4: Làm sạch dữ liệu (QUAN TRỌNG ĐỂ TRÁNH LỖI)
        
        # Xóa các dòng rỗng hoàn toàn
        df.dropna(how='all', inplace=True)
        
        # Xử lý Ngày Tháng: Chuyển đổi và xóa các dòng lỗi (NaT)
        if 'Date' in df.columns:
            df['Date'] = pd.to_datetime(df['Date'], errors='coerce')
            df = df.dropna(subset=['Date']) # Xóa dòng nếu ngày lỗi (NaT)
            
            if not df.empty:
                df['Month_Str'] = df['Date'].dt.strftime('%m-%Y')
                df['Month_Sort'] = df['Date'].dt.to_period('M')

        # Xử lý Bộ Phận: Xóa khoảng trắng thừa
        if 'Department' in df.columns:
            df['Department'] = df['Department'].astype(str).str.strip()
            df = df[df['Department'] != 'nan'] # Bỏ các dòng bộ phận là 'nan'

        # Xử lý Số: Chuyển Km và Tiền về số, lỗi thì = 0
        cols_to_numeric = ['Km_Used', 'Total_Cost']
        for col in cols_to_numeric:
            if col in df.columns:
                df[col] = pd.to_numeric(df[col], errors='coerce').fillna(0)

        return df

    except Exception as e:
        st.error(f"Lỗi đọc file: {str(e)}")
        return pd.DataFrame()

# --- 3. GIAO DIỆN CHÍNH ---
st.title("📊 Dashboard Quản Lý Đội Xe")
st.caption("Dữ liệu phân tích từ Tab Booking Car")

uploaded_file = st.file_uploader("Tải lên file Excel (Data-SuDungXe)", type=['xlsx', 'csv'])

if uploaded_file:
    df = load_data(uploaded_file)
    
    if df is not None and not df.empty:
        # --- SIDEBAR FILTERS ---
        st.sidebar.header("🔍 Bộ Lọc")
        
        # Filter Tháng (Sắp xếp đúng theo thời gian)
        if 'Month_Sort' in df.columns:
            sorted_months = df.sort_values('Month_Sort')['Month_Str'].unique()
            selected_months = st.sidebar.multiselect("Chọn Tháng", sorted_months, default=sorted_months)
        else:
            selected_months = []

        # Filter Bộ Phận
        if 'Department' in df.columns:
            all_depts = sorted(df['Department'].unique())
            selected_depts = st.sidebar.multiselect("Chọn Bộ Phận", all_depts, default=all_depts)
        else:
            selected_depts = []

        # Áp dụng lọc
        mask = pd.Series(True, index=df.index)
        if selected_months:
            mask &= df['Month_Str'].isin(selected_months)
        if selected_depts:
            mask &= df['Department'].isin(selected_depts)
            
        df_filtered = df[mask]

        if df_filtered.empty:
            st.warning("Không có dữ liệu phù hợp với bộ lọc đã chọn.")
        else:
            # --- KPI CARDS ---
            total_km = df_filtered['Km_Used'].sum()
            total_cost = df_filtered['Total_Cost'].sum()
            total_trips = len(df_filtered)
            avg_cost = total_cost / total_km if total_km > 0 else 0

            col1, col2, col3, col4 = st.columns(4)
            with col1:
                st.markdown(f'<div class="kpi-card"><div class="kpi-title">Tổng Chi Phí</div><div class="kpi-value">{total_cost:,.0f}</div><div class="kpi-note">VNĐ</div></div>', unsafe_allow_html=True)
            with col2:
                st.markdown(f'<div class="kpi-card"><div class="kpi-title">Tổng Km</div><div class="kpi-value">{total_km:,.0f}</div><div class="kpi-note">Km</div></div>', unsafe_allow_html=True)
            with col3:
                st.markdown(f'<div class="kpi-card"><div class="kpi-title">Số Chuyến Xe</div><div class="kpi-value">{total_trips}</div><div class="kpi-note">Chuyến</div></div>', unsafe_allow_html=True)
            with col4:
                st.markdown(f'<div class="kpi-card"><div class="kpi-title">Chi Phí / Km</div><div class="kpi-value">{avg_cost:,.0f}</div><div class="kpi-note">VNĐ/Km</div></div>', unsafe_allow_html=True)

            st.markdown("---")

            # --- CHARTS ---
            c1, c2 = st.columns(2)

            # Chart 1: Xu hướng theo ngày
            with c1:
                st.subheader("📅 Xu hướng chi phí theo Ngày")
                if 'Date' in df_filtered.columns:
                    daily_data = df_filtered.groupby('Date')[['Total_Cost', 'Km_Used']].sum().reset_index()
                    
                    fig = go.Figure()
                    fig.add_trace(go.Bar(x=daily_data['Date'], y=daily_data['Total_Cost'], name='Chi Phí', marker_color='#007bff'))
                    fig.add_trace(go.Scatter(x=daily_data['Date'], y=daily_data['Km_Used'], name='Km', yaxis='y2', line=dict(color='#ff5733', width=2)))
                    
                    fig.update_layout(
                        yaxis=dict(title='VNĐ'),
                        yaxis2=dict(title='Km', overlaying='y', side='right'),
                        legend=dict(orientation="h", y=1.1),
                        margin=dict(l=20, r=20, t=40, b=20),
                        height=400
                    )
                    st.plotly_chart(fig, use_container_width=True)

            # Chart 2: Top Bộ Phận
            with c2:
                st.subheader("🏢 Top Bộ Phận Sử Dụng (Chi phí)")
                if 'Department' in df_filtered.columns:
                    dept_data = df_filtered.groupby('Department')['Total_Cost'].sum().reset_index().sort_values('Total_Cost', ascending=True).tail(10)
                    fig2 = px.bar(dept_data, x='Total_Cost', y='Department', orientation='h', text_auto='.2s')
                    fig2.update_traces(textfont_size=12, textangle=0, textposition="outside", cliponaxis=False)
                    fig2.update_layout(height=400)
                    st.plotly_chart(fig2, use_container_width=True)

            # --- DATA TABLE ---
            with st.expander("📄 Xem dữ liệu chi tiết"):
                st.dataframe(df_filtered.style.format({"Total_Cost": "{:,.0f}", "Km_Used": "{:,.0f}"}))
    
    else:
        st.warning("File không chứa dữ liệu hợp lệ hoặc Tab 'Booking Car' không tìm thấy.")
else:
    st.info("👋 Vui lòng tải lên file Excel để bắt đầu.")