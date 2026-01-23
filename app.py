import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go

# --- 1. CẤU HÌNH TRANG ---
st.set_page_config(page_title="Executive Fleet Dashboard", page_icon="🏢", layout="wide")

# CSS: Tối giản, Phẳng (Flat Design), Giấu bớt viền thừa
st.markdown("""
<style>
    .block-container {padding-top: 1rem; padding-bottom: 1rem;}
    
    /* Card KPI Style */
    .kpi-box {
        background: linear-gradient(to right, #f8f9fa, #ffffff);
        border-left: 5px solid #0056b3;
        border-radius: 8px;
        padding: 15px;
        box-shadow: 0 1px 3px rgba(0,0,0,0.1);
        text-align: center;
    }
    .kpi-title {font-size: 14px; color: #6c757d; font-weight: 600; text-transform: uppercase;}
    .kpi-value {font-size: 28px; font-weight: 800; color: #0056b3;}
    
    /* Tiêu đề Section */
    .section-title {
        font-size: 18px; 
        font-weight: 700; 
        color: #343a40; 
        border-bottom: 2px solid #e9ecef; 
        padding-bottom: 5px;
        margin-bottom: 15px;
        margin-top: 20px;
    }
</style>
""", unsafe_allow_html=True)

# --- 2. HÀM XỬ LÝ DỮ LIỆU (Giữ nguyên logic chuẩn) ---
@st.cache_data
def process_data(file):
    try:
        xls = pd.ExcelFile(file)
        # Đọc dữ liệu
        df_driver_raw = pd.read_excel(xls, sheet_name='Driver', header=None)
        try:
            header_idx = df_driver_raw[df_driver_raw.eq("Biển số xe").any(axis=1)].index[0]
        except: header_idx = 2
        df_driver = pd.read_excel(xls, sheet_name='Driver', header=header_idx)
        df_cbnv = pd.read_excel(xls, sheet_name='CBNV', header=1)
        df_booking = pd.read_excel(xls, sheet_name='Booking car', header=0)

        # Clean & Merge
        df_driver.columns = df_driver.columns.str.replace('\n', ' ').str.strip()
        if 'Biển số xe' in df_driver.columns: df_driver = df_driver.drop_duplicates(subset=['Biển số xe'], keep='last')
        if 'Full Name' in df_cbnv.columns: df_cbnv = df_cbnv.drop_duplicates(subset=['Full Name'], keep='first')

        df_final = df_booking.merge(df_driver, on='Biển số xe', how='left', suffixes=('', '_Driver'))
        df_final = df_final.merge(df_cbnv, left_on='Người sử dụng xe', right_on='Full Name', how='left')

        # Format
        df_final['Ngày khởi hành'] = pd.to_datetime(df_final['Ngày khởi hành'], errors='coerce')
        df_final['Tháng'] = df_final['Ngày khởi hành'].dt.strftime('%Y-%m')
        
        cols_fill = {'Location': 'Unknown', 'Công ty': 'Other', 'BU': 'Other'}
        for col, val in cols_fill.items():
            if col in df_final.columns: df_final[col] = df_final[col].fillna(val).astype(str)
            
        return df_final
    except Exception as e: return pd.DataFrame()

# --- 3. GIAO DIỆN CHÍNH ---

# Header gọn gàng
c1, c2 = st.columns([4, 2])
with c1:
    st.markdown("### 🏢 HỆ THỐNG BÁO CÁO VẬN HÀNH (PRO VERSION)")
with c2:
    uploaded_file = st.file_uploader("Upload Excel", type=["xlsx"], label_visibility="collapsed")

if uploaded_file is not None:
    df = process_data(uploaded_file)
    if not df.empty:
        
        # --- A. BỘ LỌC ẨN (FILTER PANEL) - GỌN GÀNG HƠN ---
        with st.expander("🔍 BỘ LỌC DỮ LIỆU (Nhấn để mở/đóng)", expanded=False):
            f1, f2, f3 = st.columns(3)
            with f1:
                locs = sorted(df['Location'].unique())
                sel_loc = st.multiselect("Khu Vực", locs, default=locs)
                df_l1 = df[df['Location'].isin(sel_loc)]
            with f2:
                comps = sorted(df_l1['Công ty'].unique())
                sel_comp = st.multiselect("Công Ty", comps, default=comps)
                df_l2 = df_l1[df_l1['Công ty'].isin(sel_comp)]
            with f3:
                bus = sorted(df_l2['BU'].unique())
                sel_bu = st.multiselect("Phòng Ban (BU)", bus, default=bus)
                df_filtered = df_l2[df_l2['BU'].isin(sel_bu)]
            
            # Nút reset (Giả lập bằng cách clear session hoặc chỉ hiện text hướng dẫn)
            st.caption("💡 *Mẹo: Bấm nút 'x' trên bộ lọc để bỏ chọn nhanh, hoặc xóa hết để chọn lại từ đầu.*")

        # --- B. KPI OVERVIEW ---
        st.markdown("<br>", unsafe_allow_html=True)
        k1, k2, k3, k4 = st.columns(4)
        active_cars = df_filtered['Biển số xe'].nunique()
        top_dept = df_filtered['BU'].mode()[0] if not df_filtered.empty else "-"
        
        with k1: st.markdown(f"<div class='kpi-box'><div class='kpi-title'>Tổng Chuyến</div><div class='kpi-value'>{len(df_filtered)}</div></div>", unsafe_allow_html=True)
        with k2: st.markdown(f"<div class='kpi-box'><div class='kpi-title'>Số Xe Vận Hành</div><div class='kpi-value'>{active_cars}</div></div>", unsafe_allow_html=True)
        with k3: st.markdown(f"<div class='kpi-box'><div class='kpi-title'>Phòng Ban Top 1</div><div class='kpi-value' style='font-size:18px'>{top_dept}</div></div>", unsafe_allow_html=True)
        with k4: st.markdown(f"<div class='kpi-box'><div class='kpi-title'>Tài Xế Top 1</div><div class='kpi-value' style='font-size:18px'>{df_filtered['Tên tài xế'].mode()[0] if not df_filtered.empty else '-'}</div></div>", unsafe_allow_html=True)

        # --- C. PHÂN TÍCH CẤU TRÚC (THÔNG MINH HƠN) ---
        st.markdown("<div class='section-title'>📊 PHÂN TÍCH CẤU TRÚC & PHÂN BỔ</div>", unsafe_allow_html=True)

        # Tùy chọn góc nhìn (View Switcher)
        view_mode = st.radio("Chọn góc nhìn phân tích:", 
                             ["1. Tổng quan Luồng (Sankey)", "2. So sánh theo Công ty", "3. Chi tiết Phòng ban"], 
                             horizontal=True)

        if view_mode == "1. Tổng quan Luồng (Sankey)":
            # --- SANKEY DIAGRAM: Biểu đồ luồng (Cực xịn, không bị rối) ---
            st.info("Biểu đồ luồng hiển thị sự phân bổ từ: Vùng → Công ty → Phòng ban")
            if not df_filtered.empty:
                # Chuẩn bị dữ liệu cho Sankey
                # Gom nhóm Vùng -> Công ty
                df_s1 = df_filtered.groupby(['Location', 'Công ty']).size().reset_index(name='value')
                df_s1.columns = ['source', 'target', 'value']
                # Gom nhóm Công ty -> BU
                df_s2 = df_filtered.groupby(['Công ty', 'BU']).size().reset_index(name='value')
                df_s2.columns = ['source', 'target', 'value']
                
                # Gộp lại
                links = pd.concat([df_s1, df_s2], axis=0)
                
                # Tạo danh sách các node duy nhất
                unique_nodes = list(pd.concat([links['source'], links['target']]).unique())
                node_map = {node: i for i, node in enumerate(unique_nodes)}
                
                # Map dữ liệu về index
                links['source_id'] = links['source'].map(node_map)
                links['target_id'] = links['target'].map(node_map)
                
                # Vẽ Sankey
                fig_sankey = go.Figure(data=[go.Sankey(
                    node=dict(
                        pad=15, thickness=20, line=dict(color="black", width=0.5),
                        label=unique_nodes,
                        color="blue"
                    ),
                    link=dict(
                        source=links['source_id'],
                        target=links['target_id'],
                        value=links['value'],
                        color='rgba(0, 0, 255, 0.2)'
                    )
                )])
                fig_sankey.update_layout(title_text="Luồng phân bổ chuyến đi", font_size=10, height=500)
                st.plotly_chart(fig_sankey, use_container_width=True)

        elif view_mode == "2. So sánh theo Công ty":
            # --- BAR CHART: So sánh đơn giản ---
            col_chart1, col_chart2 = st.columns(2)
            with col_chart1:
                df_comp = df_filtered['Công ty'].value_counts().reset_index()
                df_comp.columns = ['Công ty', 'Số chuyến']
                fig = px.bar(df_comp, x='Số chuyến', y='Công ty', orientation='h', text='Số chuyến', 
                             title="Top Công Ty sử dụng xe", color='Số chuyến', color_continuous_scale='Blues')
                st.plotly_chart(fig, use_container_width=True)
            with col_chart2:
                # Biểu đồ tròn cơ cấu
                fig_pie = px.pie(df_comp, values='Số chuyến', names='Công ty', hole=0.4, title="Tỷ trọng giữa các Công ty")
                st.plotly_chart(fig_pie, use_container_width=True)

        elif view_mode == "3. Chi tiết Phòng ban":
            # --- HEATMAP / MATRIX: Nhìn chi tiết mà không rối ---
            st.write("Bảng nhiệt (Heatmap) thể hiện cường độ sử dụng xe theo từng Công ty & Phòng ban")
            if not df_filtered.empty:
                # Pivot table: Dòng là BU, Cột là Công ty (hoặc ngược lại)
                pivot = df_filtered.groupby(['Công ty', 'BU']).size().reset_index(name='Số chuyến')
                
                # Vẽ Treemap nhưng chỉ tô màu theo Công ty cha -> Đỡ rối mắt
                fig_tree = px.treemap(pivot, path=['Công ty', 'BU'], values='Số chuyến',
                                      color='Công ty', # Màu theo công ty cha cho đồng bộ
                                      title="Chi tiết từng Phòng ban (Diện tích = Số lượng)")
                st.plotly_chart(fig_tree, use_container_width=True)
                
                # Hoặc hiển thị bảng dữ liệu đẹp
                st.dataframe(pivot.sort_values('Số chuyến', ascending=False), use_container_width=True)

        # --- D. XU HƯỚNG & CHI TIẾT ---
        st.markdown("<div class='section-title'>📈 XU HƯỚNG & DỮ LIỆU CHI TIẾT</div>", unsafe_allow_html=True)
        
        t1, t2 = st.columns([2, 1])
        with t1:
            if 'Tháng' in df_filtered.columns:
                df_trend = df_filtered.groupby('Tháng').size().reset_index(name='Số chuyến')
                fig_trend = px.area(df_trend, x='Tháng', y='Số chuyến', title="Biểu đồ xu hướng theo thời gian", markers=True)
                st.plotly_chart(fig_trend, use_container_width=True)
        
        with t2:
            st.write("**Top 5 Người đi nhiều nhất**")
            top_users = df_filtered['Người sử dụng xe'].value_counts().head(5).reset_index()
            top_users.columns = ['Nhân viên', 'Số chuyến']
            st.dataframe(top_users, use_container_width=True, hide_index=True)

    else:
        st.warning("File không hợp lệ.")
else:
    st.info("Vui lòng tải file để bắt đầu.")