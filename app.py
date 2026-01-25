import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
from io import BytesIO
from pptx import Presentation
from pptx.util import Inches, Pt

# --- 1. CẤU HÌNH TRANG ---
st.set_page_config(page_title="Executive Fleet Dashboard", page_icon="🚘", layout="wide")

# CSS: Giao diện chuyên nghiệp
st.markdown("""
<style>
    .block-container {padding-top: 1rem; padding-bottom: 2rem;}
    .kpi-box {
        background: white; border-radius: 10px; padding: 20px;
        box-shadow: 0 4px 6px rgba(0,0,0,0.1); text-align: center;
        border-bottom: 4px solid #0056b3;
    }
    .kpi-title {font-size: 14px; color: #6c757d; font-weight: 600; text-transform: uppercase;}
    .kpi-value {font-size: 28px; font-weight: 800; color: #2c3e50; margin-top: 5px;}
    .kpi-sub {font-size: 12px; color: #28a745; font-weight: 500;}
    .section-header {
        font-size: 18px; font-weight: 700; color: #343a40;
        margin: 20px 0 10px 0; padding-left: 10px; border-left: 4px solid #0056b3;
    }
</style>
""", unsafe_allow_html=True)

# --- 2. HÀM XỬ LÝ DỮ LIỆU ---
@st.cache_data
def load_data_ultimate(file):
    try:
        xl = pd.ExcelFile(file, engine='openpyxl')
        
        # Tìm tên sheet
        sheet_driver = next((s for s in xl.sheet_names if 'driver' in s.lower()), None)
        sheet_booking = next((s for s in xl.sheet_names if 'booking' in s.lower()), None)
        sheet_cbnv = next((s for s in xl.sheet_names if 'cbnv' in s.lower()), None)
        
        if not sheet_booking: return "❌ Không tìm thấy sheet 'Booking car'."

        # Hàm đọc thông minh
        def smart_read(excel, sheet_name, keywords):
            df_preview = excel.parse(sheet_name, header=None, nrows=10)
            header_idx = 0
            for idx, row in df_preview.iterrows():
                row_str = row.astype(str).str.lower().tolist()
                if any(k in row_str for k in keywords):
                    header_idx = idx; break
            return excel.parse(sheet_name, header=header_idx)

        # Đọc dữ liệu
        df_bk = smart_read(xl, sheet_booking, ['ngày khởi hành'])
        df_driver = smart_read(xl, sheet_driver, ['biển số xe']) if sheet_driver else pd.DataFrame()
        df_cbnv = smart_read(xl, sheet_cbnv, ['full name']) if sheet_cbnv else pd.DataFrame()

        # Clean Columns
        df_bk.columns = df_bk.columns.str.strip()
        if not df_driver.empty: 
            df_driver.columns = df_driver.columns.str.strip()
            df_driver = df_driver.drop_duplicates(subset=['Biển số xe'], keep='last')
        if not df_cbnv.empty:
            df_cbnv.columns = df_cbnv.columns.str.strip()
            # Map cột CBNV
            col_map = {}
            for c in df_cbnv.columns:
                c_low = str(c).lower()
                if 'full name' in c_low: col_map[c] = 'Full Name'
                if 'công ty' in c_low: col_map[c] = 'Công ty'
                if 'bu' in c_low: col_map[c] = 'BU'
                if 'location' in c_low: col_map[c] = 'Location'
            df_cbnv = df_cbnv.rename(columns=col_map).drop_duplicates(subset=['Full Name'], keep='first')

        # Merge
        df_final = df_bk
        if not df_driver.empty and 'Biển số xe' in df_driver.columns:
            df_final = df_final.merge(df_driver[['Biển số xe', 'Tên tài xế']], on='Biển số xe', how='left', suffixes=('', '_D'))
            if 'Tên tài xế_D' in df_final.columns:
                df_final['Tên tài xế'] = df_final['Tên tài xế'].fillna(df_final['Tên tài xế_D'])
        
        if not df_cbnv.empty and 'Full Name' in df_cbnv.columns:
            df_final = df_final.merge(df_cbnv[['Full Name', 'Công ty', 'BU', 'Location']], left_on='Người sử dụng xe', right_on='Full Name', how='left')
            for c in ['Công ty', 'BU', 'Location']: df_final[c] = df_final[c].fillna('Unknown').astype(str)
        else:
            df_final['Công ty'] = df_final['BU'] = 'No Data'; df_final['Location'] = 'Unknown'

        # Tính toán
        df_final['Start'] = pd.to_datetime(df_final['Ngày khởi hành'].astype(str) + ' ' + df_final['Giờ khởi hành'].astype(str), errors='coerce')
        df_final['End'] = pd.to_datetime(df_final['Ngày khởi hành'].astype(str) + ' ' + df_final['Giờ kết thúc'].astype(str), errors='coerce')
        df_final.loc[df_final['End'] < df_final['Start'], 'End'] += pd.Timedelta(days=1)
        df_final['Duration'] = (df_final['End'] - df_final['Start']).dt.total_seconds() / 3600
        df_final['Tháng'] = df_final['Start'].dt.strftime('%Y-%m')
        df_final['Loại Chuyến'] = df_final['Duration'].apply(lambda x: 'Nửa ngày' if x <= 4 else 'Cả ngày')
        
        # Phạm vi
        def check_scope(r):
            s = str(r).lower()
            return "Đi Tỉnh" if any(x in s for x in ['tỉnh', 'tp.', 'bình dương', 'đồng nai', 'vũng tàu', 'hà nội']) else "Nội thành"
        df_final['Phạm Vi'] = df_final['Lộ trình'].apply(check_scope) if 'Lộ trình' in df_final.columns else 'Unknown'

        return df_final

    except Exception as e: return str(e)

# --- 3. HÀM XUẤT PPTX ---
def create_pptx(kpi_data, df_status, df_comp):
    prs = Presentation()
    
    # Slide 1: Title
    slide = prs.slides.add_slide(prs.slide_layouts[0])
    slide.shapes.title.text = "Báo Cáo Vận Hành Đội Xe"
    slide.placeholders[1].text = "Tự động tạo từ Hệ thống Quản trị"

    # Slide 2: KPI Tổng quan
    slide = prs.slides.add_slide(prs.slide_layouts[1])
    slide.shapes.title.text = "Tổng Quan Hiệu Suất (KPI)"
    content = slide.placeholders[1]
    
    text = f"""
    - Tổng số chuyến đi: {kpi_data['total_trips']}
    - Tổng giờ vận hành: {kpi_data['total_hours']:,.0f} giờ
    - Tỷ lệ lấp đầy (Occupancy): {kpi_data['occupancy']:.1f}%
      (Công thức: Tổng giờ chạy / (Số xe * Số ngày * 9h))
    - Số xe hoạt động: {kpi_data['active_cars']} / {kpi_data['total_cars']} xe
    """
    content.text = text

    # Slide 3: Tỷ lệ Hủy/Từ chối
    slide = prs.slides.add_slide(prs.slide_layouts[1])
    slide.shapes.title.text = "Chất Lượng Vận Hành"
    
    # Tạo bảng Status
    rows, cols = df_status.shape[0] + 1, df_status.shape[1]
    table = slide.shapes.add_table(rows, cols, Inches(1), Inches(2), Inches(8), Inches(3)).table
    
    # Header
    for i, col_name in enumerate(df_status.columns):
        table.cell(0, i).text = str(col_name)
    
    # Body
    for i, row in enumerate(df_status.itertuples(index=False)):
        for j, val in enumerate(row):
            table.cell(i+1, j).text = str(val)

    # Slide 4: Phân bổ Công ty
    slide = prs.slides.add_slide(prs.slide_layouts[1])
    slide.shapes.title.text = "Phân Bổ Theo Công Ty"
    
    rows, cols = min(df_comp.shape[0], 10) + 1, df_comp.shape[1] # Lấy top 10
    table = slide.shapes.add_table(rows, cols, Inches(1), Inches(2), Inches(8), Inches(4)).table
    
    for i, col_name in enumerate(df_comp.columns):
        table.cell(0, i).text = str(col_name)
        
    for i, row in enumerate(df_comp.head(10).itertuples(index=False)):
        for j, val in enumerate(row):
            table.cell(i+1, j).text = str(val)
            
    buffer = BytesIO()
    prs.save(buffer)
    buffer.seek(0)
    return buffer

# --- 4. GIAO DIỆN CHÍNH ---
st.markdown("### 🏢 HỆ THỐNG QUẢN TRỊ & BÁO CÁO ĐỘI XE")
uploaded_file = st.file_uploader("Upload file Excel", type=['xlsx'], label_visibility="collapsed")

if uploaded_file:
    df = load_data_ultimate(uploaded_file)
    if isinstance(df, str): st.error(df); st.stop()
    
    # --- BỘ LỌC ---
    with st.expander("🔍 BỘ LỌC DỮ LIỆU", expanded=True):
        c1, c2, c3 = st.columns(3)
        locs = sorted(df['Location'].unique()); sel_loc = c1.multiselect("Khu vực", locs, default=locs)
        df_l1 = df[df['Location'].isin(sel_loc)]
        comps = sorted(df_l1['Công ty'].unique()); sel_comp = c2.multiselect("Công ty", comps, default=comps)
        df_l2 = df_l1[df_l1['Công ty'].isin(sel_comp)]
        bus = sorted(df_l2['BU'].unique()); sel_bu = c3.multiselect("Phòng ban", bus, default=bus)
        df_filtered = df_l2[df_l2['BU'].isin(sel_bu)]
        st.caption(f"Dữ liệu: {len(df_filtered)} chuyến")

    # --- TÍNH KPI ---
    total_cars = 21
    if len(sel_loc) == 1:
        if 'HCM' in str(sel_loc[0]) or 'NAM' in str(sel_loc[0]).upper(): total_cars = 16
        elif 'HN' in str(sel_loc[0]) or 'BAC' in str(sel_loc[0]).upper(): total_cars = 5
        
    days = (df_filtered['Start'].max() - df_filtered['Start'].min()).days + 1 if not df_filtered.empty else 1
    cap_hours = total_cars * max(days, 1) * 9
    used_hours = df_filtered['Duration'].sum()
    occupancy = (used_hours / cap_hours * 100) if cap_hours > 0 else 0
    
    # KPI Dict cho PPTX
    kpi_data = {
        'total_trips': len(df_filtered),
        'total_hours': used_hours,
        'occupancy': occupancy,
        'active_cars': df_filtered['Biển số xe'].nunique(),
        'total_cars': total_cars
    }

    # Hiển thị KPI Cards
    st.markdown("<br>", unsafe_allow_html=True)
    c1, c2, c3, c4 = st.columns(4)
    c1.markdown(f"<div class='kpi-box'><div class='kpi-title'>Tổng Chuyến</div><div class='kpi-value'>{len(df_filtered)}</div></div>", unsafe_allow_html=True)
    c2.markdown(f"<div class='kpi-box'><div class='kpi-title'>Tổng Giờ</div><div class='kpi-value'>{used_hours:,.0f}h</div></div>", unsafe_allow_html=True)
    c3.markdown(f"<div class='kpi-box'><div class='kpi-title'>Lấp Đầy (Occupancy)</div><div class='kpi-value'>{occupancy:.1f}%</div><div class='kpi-sub'>Công thức: Giờ chạy / ({total_cars} xe * {days} ngày * 9h)</div></div>", unsafe_allow_html=True)
    c4.markdown(f"<div class='kpi-box'><div class='kpi-title'>Xe Hoạt Động</div><div class='kpi-value'>{df_filtered['Biển số xe'].nunique()}/{total_cars}</div></div>", unsafe_allow_html=True)

    # --- CÁC PHÂN TÍCH ---
    t1, t2 = st.tabs(["📊 Phân Tích & Biểu Đồ", "📉 Chất Lượng & Cancel Rate"])
    
    with t1:
        # Chọn loại biểu đồ (Yêu cầu 3)
        chart_type = st.radio("Chọn kiểu biểu đồ:", ["Bar (Cột)", "Pie (Tròn)", "Donut (Vành khuyên)"], horizontal=True)
        
        c1, c2 = st.columns(2)
        with c1:
            st.markdown("<div class='section-header'>Theo Công Ty</div>", unsafe_allow_html=True)
            df_comp = df_filtered['Công ty'].value_counts().reset_index()
            df_comp.columns = ['Công ty', 'Số chuyến']
            
            if "Bar" in chart_type:
                fig = px.bar(df_comp, x='Số chuyến', y='Công ty', orientation='h', text='Số chuyến', title="Top Công Ty")
            elif "Pie" in chart_type:
                fig = px.pie(df_comp, values='Số chuyến', names='Công ty', title="Tỷ trọng Công ty")
            else:
                fig = px.pie(df_comp, values='Số chuyến', names='Công ty', hole=0.4, title="Tỷ trọng Công ty")
            st.plotly_chart(fig, use_container_width=True)
            
        with c2:
            st.markdown("<div class='section-header'>Nội thành vs Đi Tỉnh</div>", unsafe_allow_html=True)
            df_scope = df_filtered['Phạm Vi'].value_counts().reset_index()
            df_scope.columns = ['Phạm Vi', 'Số chuyến']
            
            if "Bar" in chart_type:
                fig2 = px.bar(df_scope, x='Phạm Vi', y='Số chuyến', text='Số chuyến', color='Phạm Vi')
            else:
                fig2 = px.pie(df_scope, values='Số chuyến', names='Phạm Vi', hole=0.4 if "Donut" in chart_type else 0)
            st.plotly_chart(fig2, use_container_width=True)

    with t2:
        # Tỷ lệ Cancel / Reject (Yêu cầu 2)
        st.markdown("<div class='section-header'>Tỷ Lệ Hủy & Từ Chối</div>", unsafe_allow_html=True)
        
        if 'Tình trạng đơn yêu cầu' in df_filtered.columns:
            # Tính toán
            total = len(df_filtered)
            counts = df_filtered['Tình trạng đơn yêu cầu'].fillna('Unknown').value_counts()
            
            cancel_count = counts.get('CANCELED', 0) + counts.get('CANCELLED', 0) # Bắt lỗi chính tả
            reject_count = counts.get('REJECTED_BY_ADMIN', 0)
            
            cancel_rate = (cancel_count / total * 100) if total > 0 else 0
            reject_rate = (reject_count / total * 100) if total > 0 else 0
            
            # Hiển thị số to
            cc1, cc2, cc3 = st.columns(3)
            cc1.metric("Tỷ lệ Hủy (Cancel)", f"{cancel_rate:.1f}%", f"{cancel_count} chuyến", delta_color="inverse")
            cc2.metric("Tỷ lệ Từ chối (Reject)", f"{reject_rate:.1f}%", f"{reject_count} chuyến", delta_color="inverse")
            cc3.metric("Hoàn thành (Closed)", f"{100 - cancel_rate - reject_rate:.1f}%", delta_color="normal")
            
            # Bảng chi tiết
            df_status = counts.reset_index()
            df_status.columns = ['Trạng thái', 'Số lượng']
            df_status['Tỷ lệ %'] = (df_status['Số lượng'] / total * 100).map('{:.1f}%'.format)
            st.dataframe(df_status, use_container_width=True)
            
            # Chuẩn bị data cho PPTX
            df_status_pptx = df_status
        else:
            st.warning("Không có cột 'Tình trạng đơn yêu cầu'")
            df_status_pptx = pd.DataFrame()

    # --- NÚT TẢI PPTX (Yêu cầu 1) ---
    st.markdown("---")
    st.markdown("### 📥 Xuất Báo Cáo")
    
    # Tạo PPTX
    pptx_file = create_pptx(kpi_data, df_status_pptx, df_filtered['Công ty'].value_counts().reset_index())
    
    c_down1, c_down2 = st.columns([1, 4])
    with c_down1:
        st.download_button(
            label="📄 Tải Báo Cáo PPTX",
            data=pptx_file,
            file_name="Bao_Cao_Doi_Xe.pptx",
            mime="application/vnd.openxmlformats-officedocument.presentationml.presentation",
            type="primary"
        )
    with c_down2:
        st.info("💡 File PPTX sẽ chứa các bảng số liệu đã tính toán. Bạn có thể copy bảng này vào slide của sếp và Insert Chart trong PowerPoint cực nhanh.")

else:
    st.info("👋 Upload file Excel để bắt đầu.")