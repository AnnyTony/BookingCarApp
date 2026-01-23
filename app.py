import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go

# --- 1. CẤU HÌNH TRANG ---
st.set_page_config(page_title="Executive Fleet Dashboard", page_icon="🚘", layout="wide")

# CSS: Flat Design & KPI Cards (Lấy từ code của bạn + tinh chỉnh)
st.markdown("""
<style>
    .block-container {padding-top: 1rem; padding-bottom: 2rem;}
    
    /* KPI Box đẹp mắt */
    .kpi-box {
        background: white;
        border-radius: 10px;
        padding: 20px;
        box-shadow: 0 4px 6px rgba(0,0,0,0.1);
        text-align: center;
        border-bottom: 4px solid #0056b3;
        transition: transform 0.2s;
    }
    .kpi-box:hover {transform: translateY(-5px);}
    .kpi-title {font-size: 14px; color: #6c757d; font-weight: 600; text-transform: uppercase; letter-spacing: 1px;}
    .kpi-value {font-size: 32px; font-weight: 800; color: #2c3e50; margin-top: 10px;}
    .kpi-sub {font-size: 12px; color: #28a745; font-weight: 500;}
    
    /* Tiêu đề Section */
    .section-header {
        font-size: 20px; font-weight: 700; color: #343a40;
        margin: 25px 0 15px 0; padding-left: 10px;
        border-left: 5px solid #0056b3;
    }
</style>
""", unsafe_allow_html=True)

# --- 2. HÀM XỬ LÝ DỮ LIỆU THÔNG MINH (Kết hợp Logic của mình + Driver của bạn) ---
@st.cache_data
def load_data_ultimate(file):
    try:
        xl = pd.ExcelFile(file, engine='openpyxl')
        
        # 1. Tìm tên sheet linh hoạt (Tránh lỗi nếu user đổi tên sheet)
        sheet_driver = next((s for s in xl.sheet_names if 'driver' in s.lower() or 'tài xế' in s.lower()), None)
        sheet_booking = next((s for s in xl.sheet_names if 'booking' in s.lower()), None)
        sheet_cbnv = next((s for s in xl.sheet_names if 'cbnv' in s.lower() or 'staff' in s.lower()), None)
        
        if not sheet_booking: return "❌ Không tìm thấy sheet 'Booking car'."

        # --- Hàm đọc Header thông minh (Quét 10 dòng đầu) ---
        def smart_read(excel, sheet_name, keywords):
            df_preview = excel.parse(sheet_name, header=None, nrows=10)
            header_idx = 0
            for idx, row in df_preview.iterrows():
                row_str = row.astype(str).str.lower().tolist()
                if any(k in row_str for k in keywords):
                    header_idx = idx
                    break
            return excel.parse(sheet_name, header=header_idx)

        # 2. Đọc & Xử lý Driver (Của bạn)
        if sheet_driver:
            df_driver = smart_read(xl, sheet_driver, ['biển số xe', 'tên tài xế'])
            # Clean cột
            df_driver.columns = df_driver.columns.str.strip().str.replace('\n', ' ')
            df_driver = df_driver.drop_duplicates(subset=['Biển số xe'], keep='last')
        else:
            df_driver = pd.DataFrame()

        # 3. Đọc & Xử lý CBNV (Của bạn + Map cột thông minh)
        if sheet_cbnv:
            df_cbnv = smart_read(xl, sheet_cbnv, ['full name', 'họ tên', 'công ty'])
            # Map tên cột chuẩn
            col_map = {}
            for c in df_cbnv.columns:
                c_low = str(c).lower()
                if 'full name' in c_low: col_map[c] = 'Full Name'
                if 'công ty' in c_low: col_map[c] = 'Công ty'
                if 'bu' in c_low or 'bộ phận' in c_low: col_map[c] = 'BU'
                if 'location' in c_low: col_map[c] = 'Location'
            df_cbnv = df_cbnv.rename(columns=col_map)
            df_cbnv = df_cbnv.drop_duplicates(subset=['Full Name'], keep='first')
        else:
            df_cbnv = pd.DataFrame()

        # 4. Đọc Booking & Merge
        df_bk = smart_read(xl, sheet_booking, ['ngày khởi hành', 'giờ khởi hành'])
        df_bk.columns = df_bk.columns.str.strip()

        # Merge dữ liệu (Driver + CBNV)
        # Merge Driver
        if not df_driver.empty and 'Biển số xe' in df_driver.columns:
            df_final = pd.merge(df_bk, df_driver[['Biển số xe', 'Tên tài xế']], on='Biển số xe', how='left', suffixes=('', '_Driver'))
            # Ưu tiên tên tài xế trong booking, nếu ko có lấy từ bảng Driver
            if 'Tên tài xế_Driver' in df_final.columns:
                df_final['Tên tài xế'] = df_final['Tên tài xế'].fillna(df_final['Tên tài xế_Driver'])
        else:
            df_final = df_bk

        # Merge CBNV
        if not df_cbnv.empty and 'Full Name' in df_cbnv.columns:
            df_final = pd.merge(df_final, df_cbnv[['Full Name', 'Công ty', 'BU', 'Location']], 
                                left_on='Người sử dụng xe', right_on='Full Name', how='left')
            
            # Fillna
            for col in ['Công ty', 'BU', 'Location']:
                df_final[col] = df_final[col].fillna('Unknown')
        else:
            df_final['Công ty'] = 'No Data'
            df_final['BU'] = 'No Data'
            df_final['Location'] = 'Unknown'

        # --- LOGIC TÍNH TOÁN (CỦA MÌNH - QUAN TRỌNG) ---
        # 1. Ngày giờ
        df_final['Start_Datetime'] = pd.to_datetime(df_final['Ngày khởi hành'].astype(str) + ' ' + df_final['Giờ khởi hành'].astype(str), errors='coerce')
        df_final['End_Datetime'] = pd.to_datetime(df_final['Ngày khởi hành'].astype(str) + ' ' + df_final['Giờ kết thúc'].astype(str), errors='coerce')
        mask_overnight = df_final['End_Datetime'] < df_final['Start_Datetime']
        df_final.loc[mask_overnight, 'End_Datetime'] += pd.Timedelta(days=1)
        
        df_final['Duration_Hours'] = (df_final['End_Datetime'] - df_final['Start_Datetime']).dt.total_seconds() / 3600
        df_final['Tháng'] = df_final['Start_Datetime'].dt.strftime('%Y-%m')
        
        # 2. Phân loại
        df_final['Loại Chuyến'] = df_final['Duration_Hours'].apply(lambda x: 'Nửa ngày' if x <= 4 else 'Cả ngày')
        
        def check_scope(route):
            s = str(route).lower()
            return "Đi Tỉnh" if any(x in s for x in ['tỉnh', 'tp.', 'bình dương', 'đồng nai', 'vũng tàu', 'hà nội']) else "Nội thành"
        if 'Lộ trình' in df_final.columns:
            df_final['Phạm Vi'] = df_final['Lộ trình'].apply(check_scope)
        else:
            df_final['Phạm Vi'] = 'Unknown'

        return df_final

    except Exception as e:
        return f"Lỗi xử lý: {str(e)}"

# --- 3. GIAO DIỆN CHÍNH ---
st.markdown("### 🏢 HỆ THỐNG QUẢN TRỊ ĐỘI XE (ULTIMATE VERSION)")
uploaded_file = st.file_uploader("Upload file Excel (Booking, Driver, CBNV)", type=['xlsx'], label_visibility="collapsed")

if uploaded_file:
    df = load_data_ultimate(uploaded_file)
    if isinstance(df, str):
        st.error(df)
        st.stop()
        
    # --- A. BỘ LỌC CASCADING (TRONG EXPANDER) ---
    with st.expander("🔍 BỘ LỌC DỮ LIỆU (Nhấn để mở rộng)", expanded=True):
        f1, f2, f3 = st.columns(3)
        with f1:
            locs = sorted(df['Location'].unique())
            sel_loc = st.multiselect("1. Khu Vực (Location)", locs, default=locs)
            df_l1 = df[df['Location'].isin(sel_loc)]
        with f2:
            comps = sorted(df_l1['Công ty'].unique())
            sel_comp = st.multiselect("2. Công Ty", comps, default=comps)
            df_l2 = df_l1[df_l1['Công ty'].isin(sel_comp)]
        with f3:
            bus = sorted(df_l2['BU'].unique())
            sel_bu = st.multiselect("3. Phòng Ban (BU)", bus, default=bus)
            df_filtered = df_l2[df_l2['BU'].isin(sel_bu)]
            
        st.caption(f"Đang hiển thị: {len(df_filtered)} chuyến đi")

    # --- B. KPI CARDS (LOGIC CỦA MÌNH + UI CỦA BẠN) ---
    # Logic Occupancy (Tính toán thông minh)
    total_cars = 21 # Mặc định
    if len(sel_loc) == 1:
        if 'HCM' in sel_loc[0] or 'NAM' in sel_loc[0].upper(): total_cars = 16
        elif 'HN' in sel_loc[0] or 'BAC' in sel_loc[0].upper(): total_cars = 5
    
    if 'Start_Datetime' in df_filtered.columns and not df_filtered.empty:
        days = (df_filtered['Start_Datetime'].max() - df_filtered['Start_Datetime'].min()).days + 1
        cap_hours = total_cars * max(days, 1) * 9
        used_hours = df_filtered['Duration_Hours'].sum()
        occupancy = (used_hours / cap_hours * 100)
    else: occupancy = 0

    st.markdown("<br>", unsafe_allow_html=True)
    k1, k2, k3, k4 = st.columns(4)
    
    with k1: st.markdown(f"<div class='kpi-box'><div class='kpi-title'>Tổng Chuyến</div><div class='kpi-value'>{len(df_filtered)}</div></div>", unsafe_allow_html=True)
    with k2: st.markdown(f"<div class='kpi-box'><div class='kpi-title'>Tổng Giờ Vận Hành</div><div class='kpi-value'>{used_hours:,.0f}h</div></div>", unsafe_allow_html=True)
    with k3: st.markdown(f"<div class='kpi-box'><div class='kpi-title'>Tỷ Lệ Lấp Đầy</div><div class='kpi-value'>{occupancy:.1f}%</div><div class='kpi-sub'>Trên {total_cars} xe</div></div>", unsafe_allow_html=True)
    with k4: st.markdown(f"<div class='kpi-box'><div class='kpi-title'>Xe Hoạt Động</div><div class='kpi-value'>{df_filtered['Biển số xe'].nunique()}/{total_cars}</div></div>", unsafe_allow_html=True)

    # --- C. PHÂN TÍCH CHUYÊN SÂU ---
    
    # 1. BIỂU ĐỒ SANKEY (Luồng dữ liệu - Của bạn)
    st.markdown("<div class='section-header'>📊 LUỒNG PHÂN BỔ: VÙNG ➔ CÔNG TY ➔ BU</div>", unsafe_allow_html=True)
    if not df_filtered.empty:
        # Tạo dữ liệu Sankey
        sankey_data1 = df_filtered.groupby(['Location', 'Công ty']).size().reset_index(name='val')
        sankey_data1.columns = ['source', 'target', 'val']
        sankey_data2 = df_filtered.groupby(['Công ty', 'BU']).size().reset_index(name='val')
        sankey_data2.columns = ['source', 'target', 'val']
        links = pd.concat([sankey_data1, sankey_data2])
        
        nodes = list(pd.concat([links['source'], links['target']]).unique())
        node_map = {node: i for i, node in enumerate(nodes)}
        
        fig_sankey = go.Figure(data=[go.Sankey(
            node=dict(pad=15, thickness=20, line=dict(color="black", width=0.5), label=nodes, color="rgba(0,86,179,0.8)"),
            link=dict(source=links['source'].map(node_map), target=links['target'].map(node_map), value=links['val'], color='rgba(0,86,179,0.2)')
        )])
        fig_sankey.update_layout(height=400, margin=dict(l=0,r=0,t=0,b=0))
        st.plotly_chart(fig_sankey, use_container_width=True)

    # 2. XU HƯỚNG & CHI TIẾT
    c1, c2 = st.columns([1, 1])
    
    with c1:
        st.markdown("<div class='section-header'>📈 LOẠI CHUYẾN & PHẠM VI</div>", unsafe_allow_html=True)
        # Biểu đồ cột chồng (Logic của mình)
        df_type = df_filtered.groupby(['Công ty', 'Loại Chuyến']).size().reset_index(name='Count')
        fig_bar = px.bar(df_type, x='Công ty', y='Count', color='Loại Chuyến', title="Nửa ngày vs Cả ngày", barmode='group')
        st.plotly_chart(fig_bar, use_container_width=True)

    with c2:
        st.markdown("<div class='section-header'>🏆 TOP TÀI XẾ & NGƯỜI DÙNG</div>", unsafe_allow_html=True)
        tab_driver, tab_user = st.tabs(["Tài Xế (Driver)", "Người Dùng (User)"])
        
        with tab_driver:
            if 'Tên tài xế' in df_filtered.columns:
                top_driver = df_filtered['Tên tài xế'].value_counts().head(5).reset_index()
                top_driver.columns = ['Tài xế', 'Số chuyến']
                st.dataframe(top_driver, use_container_width=True, hide_index=True)
                
        with tab_user:
            top_user = df_filtered['Người sử dụng xe'].value_counts().head(5).reset_index()
            top_user.columns = ['Nhân viên', 'Số chuyến']
            st.dataframe(top_user, use_container_width=True, hide_index=True)

else:
    st.info("👋 Hãy upload file Excel để bắt đầu phân tích.")