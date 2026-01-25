import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import matplotlib.pyplot as plt
from io import BytesIO
from pptx import Presentation
from pptx.util import Inches, Pt

# --- 1. CẤU HÌNH TRANG & CSS ---
st.set_page_config(page_title="Fleet Management Pro", page_icon="🚘", layout="wide")

st.markdown("""
<style>
    .block-container {padding-top: 1rem; padding-bottom: 3rem;}
    
    /* KPI Card Style - Power BI */
    .kpi-card {
        background-color: white; border-radius: 8px; padding: 15px;
        box-shadow: 0 2px 4px rgba(0,0,0,0.08); border-left: 5px solid #0078d4;
        margin-bottom: 10px;
    }
    .kpi-title {font-size: 13px; color: #666; font-weight: 600; text-transform: uppercase;}
    .kpi-value {font-size: 28px; font-weight: 700; color: #333; margin: 5px 0;}
    .kpi-sub {font-size: 11px; color: #28a745; font-weight: 500;}
    
    /* Header Chart */
    .chart-header {
        font-size: 16px; font-weight: 700; color: #0078d4; 
        margin-bottom: 10px; border-bottom: 2px solid #f0f2f6; padding-bottom: 5px;
    }
</style>
""", unsafe_allow_html=True)

# --- 2. HÀM XỬ LÝ DỮ LIỆU ---
@st.cache_data
def load_data_final(file):
    try:
        xl = pd.ExcelFile(file, engine='openpyxl')
        
        # Tìm sheet
        sheet_driver = next((s for s in xl.sheet_names if 'driver' in s.lower()), None)
        sheet_booking = next((s for s in xl.sheet_names if 'booking' in s.lower()), None)
        sheet_cbnv = next((s for s in xl.sheet_names if 'cbnv' in s.lower()), None)
        
        if not sheet_booking: return "❌ Không tìm thấy sheet 'Booking car'."

        # Hàm đọc header thông minh
        def smart_read(excel, sheet_name, keywords):
            df_preview = excel.parse(sheet_name, header=None, nrows=10)
            header_idx = 0
            for idx, row in df_preview.iterrows():
                row_str = row.astype(str).str.lower().tolist()
                if any(k in row_str for k in keywords):
                    header_idx = idx; break
            return excel.parse(sheet_name, header=header_idx)

        # Load Data
        df_bk = smart_read(xl, sheet_booking, ['ngày khởi hành'])
        df_driver = smart_read(xl, sheet_driver, ['biển số xe']) if sheet_driver else pd.DataFrame()
        df_cbnv = smart_read(xl, sheet_cbnv, ['full name']) if sheet_cbnv else pd.DataFrame()

        # Clean Headers
        df_bk.columns = df_bk.columns.str.strip()
        
        # Merge Driver
        df_final = df_bk
        if not df_driver.empty:
            df_driver.columns = df_driver.columns.str.strip()
            if 'Biển số xe' in df_driver.columns:
                df_driver = df_driver.drop_duplicates(subset=['Biển số xe'], keep='last')
                df_final = df_final.merge(df_driver[['Biển số xe', 'Tên tài xế']], on='Biển số xe', how='left', suffixes=('', '_D'))
                if 'Tên tài xế_D' in df_final.columns:
                    df_final['Tên tài xế'] = df_final['Tên tài xế'].fillna(df_final['Tên tài xế_D'])

        # Merge CBNV
        if not df_cbnv.empty:
            df_cbnv.columns = df_cbnv.columns.str.strip()
            col_map = {}
            for c in df_cbnv.columns:
                if 'full name' in str(c).lower(): col_map[c] = 'Full Name'
                if 'công ty' in str(c).lower(): col_map[c] = 'Công ty'
                if 'bu' in str(c).lower(): col_map[c] = 'BU'
                if 'location' in str(c).lower(): col_map[c] = 'Location'
            
            # Kiểm tra cột tồn tại trước khi rename
            available_cols = [c for c in col_map.keys() if c in df_cbnv.columns]
            df_cbnv = df_cbnv[available_cols].rename(columns=col_map)
            
            if 'Full Name' in df_cbnv.columns:
                df_cbnv = df_cbnv.drop_duplicates(subset=['Full Name'], keep='first')
                df_final = df_final.merge(df_cbnv, left_on='Người sử dụng xe', right_on='Full Name', how='left')

        # Fillna & Format
        for c in ['Công ty', 'BU', 'Location']:
            if c not in df_final.columns: df_final[c] = 'Unknown'
            else: df_final[c] = df_final[c].fillna('Unknown').astype(str)
            
        df_final['Start'] = pd.to_datetime(df_final['Ngày khởi hành'].astype(str) + ' ' + df_final['Giờ khởi hành'].astype(str), errors='coerce')
        df_final['End'] = pd.to_datetime(df_final['Ngày khởi hành'].astype(str) + ' ' + df_final['Giờ kết thúc'].astype(str), errors='coerce')
        df_final.loc[df_final['End'] < df_final['Start'], 'End'] += pd.Timedelta(days=1)
        df_final['Duration'] = (df_final['End'] - df_final['Start']).dt.total_seconds() / 3600
        df_final['Tháng'] = df_final['Start'].dt.strftime('%Y-%m')
        df_final['Năm'] = df_final['Start'].dt.year
        df_final['Loại Chuyến'] = df_final['Duration'].apply(lambda x: 'Nửa ngày' if x <= 4 else 'Cả ngày')
        
        # Scope
        def check_scope(r):
            s = str(r).lower()
            return "Đi Tỉnh" if any(x in s for x in ['tỉnh', 'tp.', 'bình dương', 'đồng nai', 'vũng tàu', 'hà nội']) else "Nội thành"
        df_final['Phạm Vi'] = df_final['Lộ trình'].apply(check_scope) if 'Lộ trình' in df_final.columns else 'Unknown'

        return df_final
    except Exception as e: return str(e)

# --- 3. HÀM TẠO ẢNH CHO PPTX ---
def get_chart_img(data, x, y, kind='bar', title=''):
    plt.figure(figsize=(6, 4))
    if kind == 'bar':
        plt.barh(data[x], data[y], color='#0078d4')
        plt.xlabel(y)
    elif kind == 'pie':
        plt.pie(data[y], labels=data[x], autopct='%1.1f%%')
    plt.title(title); plt.tight_layout()
    img = BytesIO(); plt.savefig(img, format='png', dpi=100); plt.close(); img.seek(0)
    return img

# --- 4. HÀM XUẤT PPTX ---
def export_pptx(kpi, df_status, df_comp):
    prs = Presentation()
    
    # Slide 1: Title
    slide = prs.slides.add_slide(prs.slide_layouts[0])
    slide.shapes.title.text = "Báo Cáo Vận Hành Đội Xe"
    slide.placeholders[1].text = f"Tổng hợp đến tháng {kpi['last_month']}"
    
    # Slide 2: KPI
    slide = prs.slides.add_slide(prs.slide_layouts[1])
    slide.shapes.title.text = "Tổng Quan Hiệu Suất"
    slide.placeholders[1].text = f"""
    • Tổng chuyến: {kpi['trips']} | Tổng giờ: {kpi['hours']:,.0f}h
    • Tỷ lệ Lấp đầy: {kpi['occupancy']:.1f}%
    • Tỷ lệ Hoàn thành: {kpi['success_rate']:.1f}%
    • Tỷ lệ Hủy/Từ chối: {kpi['cancel_rate'] + kpi['reject_rate']:.1f}%
    """
    
    # Slide 3: Chart
    slide = prs.slides.add_slide(prs.slide_layouts[5])
    slide.shapes.title.text = "Phân Bổ Công Ty & Trạng Thái"
    img1 = get_chart_img(df_comp.head(8), 'Công ty', 'Số chuyến', 'bar', 'Top Công ty')
    slide.shapes.add_picture(img1, Inches(0.5), Inches(2), Inches(4.5), Inches(3.5))
    
    img2 = get_chart_img(df_status, 'Trạng thái', 'Số lượng', 'pie', 'Trạng thái')
    slide.shapes.add_picture(img2, Inches(5.2), Inches(2), Inches(4.5), Inches(3.5))
    
    out = BytesIO(); prs.save(out); out.seek(0)
    return out

# --- 5. GIAO DIỆN CHÍNH ---
st.title("📊 Fleet Management Pro")
uploaded_file = st.file_uploader("Upload Excel", type=['xlsx'], label_visibility="collapsed")

if uploaded_file:
    df = load_data_final(uploaded_file)
    if isinstance(df, str): st.error(df); st.stop()
    
    # --- CÂY THƯ MỤC LỌC (HIERARCHY FILTER) ---
    with st.sidebar:
        st.header("🗂️ Cây Thư Mục Lọc")
        st.info("Chọn lần lượt từ trên xuống để xem chi tiết (Drill-down)")
        
        # Level 1: Location
        locs = ["Tất cả"] + sorted(df['Location'].unique().tolist())
        sel_loc = st.selectbox("1. Khu vực (Region):", locs)
        
        # LỌC CẤP 1
        df_l1 = df if sel_loc == "Tất cả" else df[df['Location'] == sel_loc]
        
        # Level 2: Company (Chỉ hiện Công ty thuộc Region đã chọn)
        comps = ["Tất cả"] + sorted(df_l1['Công ty'].unique().tolist())
        sel_comp = st.selectbox("2. Công ty (Entity):", comps)
        
        # LỌC CẤP 2
        df_l2 = df_l1 if sel_comp == "Tất cả" else df_l1[df_l1['Công ty'] == sel_comp]
        
        # Level 3: BU (Chỉ hiện BU thuộc Công ty đã chọn)
        bus = ["Tất cả"] + sorted(df_l2['BU'].unique().tolist())
        sel_bu = st.selectbox("3. Phòng ban (BU):", bus)
        
        # LỌC CẤP 3
        df_filtered = df_l2 if sel_bu == "Tất cả" else df_l2[df_l2['BU'] == sel_bu]
        
        st.markdown("---")
        st.caption(f"Đang xem: **{len(df_filtered)}** chuyến")

    # --- KPI SECTION (CÓ TỶ LỆ HOÀN THÀNH) ---
    # Tính toán
    total_cars = 21
    if 'HCM' in sel_loc or 'NAM' in sel_loc.upper(): total_cars = 16
    elif 'HN' in sel_loc or 'BAC' in sel_loc.upper(): total_cars = 5
    
    days = (df['Start'].max() - df['Start'].min()).days + 1 if not df.empty else 1
    cap = total_cars * max(days, 1) * 9
    used = df_filtered['Duration'].sum()
    occupancy = (used / cap * 100) if cap > 0 else 0
    
    # Status Rates
    counts = df_filtered['Tình trạng đơn yêu cầu'].fillna('Unknown').value_counts()
    total = len(df_filtered)
    cancel = counts.get('CANCELED', 0) + counts.get('CANCELLED', 0)
    reject = counts.get('REJECTED_BY_ADMIN', 0)
    completed = counts.get('CLOSED', 0) + counts.get('APPROVED', 0) # Coi Approved là sắp hoàn thành
    
    suc_rate = (completed / total * 100) if total > 0 else 0
    can_rate = (cancel / total * 100) if total > 0 else 0
    rej_rate = (reject / total * 100) if total > 0 else 0

    # KPI UI
    k1, k2, k3, k4, k5 = st.columns(5)
    k1.markdown(f"<div class='kpi-card'><div class='kpi-title'>Tổng Chuyến</div><div class='kpi-value'>{total}</div></div>", unsafe_allow_html=True)
    k2.markdown(f"<div class='kpi-card'><div class='kpi-title'>Giờ Vận Hành</div><div class='kpi-value'>{used:,.0f}</div></div>", unsafe_allow_html=True)
    k3.markdown(f"<div class='kpi-card'><div class='kpi-title'>Occupancy</div><div class='kpi-value'>{occupancy:.1f}%</div><div class='kpi-sub'>Trên {total_cars} xe</div></div>", unsafe_allow_html=True)
    k4.markdown(f"<div class='kpi-card' style='border-left: 5px solid #107c10'><div class='kpi-title'>Hoàn Thành</div><div class='kpi-value' style='color:#107c10'>{suc_rate:.1f}%</div></div>", unsafe_allow_html=True)
    k5.markdown(f"<div class='kpi-card' style='border-left: 5px solid #d13438'><div class='kpi-title'>Hủy / Từ Chối</div><div class='kpi-value' style='color:#d13438'>{can_rate + rej_rate:.1f}%</div></div>", unsafe_allow_html=True)

    # --- DASHBOARD TABS ---
    t1, t2, t3 = st.tabs(["📊 Phân Tích Đơn Vị (Drill-down)", "📈 Xu Hướng & Top", "📉 Chất Lượng Vận Hành"])
    
    with t1:
        st.write("#### Phân tích theo Cấu trúc")
        
        # LOGIC BIỂU ĐỒ THÔNG MINH (Drill-down Chart)
        if sel_comp == "Tất cả":
            # Level 1: Chưa chọn Cty -> Vẽ biểu đồ so sánh các Công ty
            st.info(f"Đang hiển thị so sánh các Công ty tại {sel_loc}")
            df_g = df_filtered['Công ty'].value_counts().reset_index()
            df_g.columns = ['Công ty', 'Số chuyến']
            fig = px.bar(df_g, x='Số chuyến', y='Công ty', orientation='h', 
                         text='Số chuyến', title="Số chuyến theo Công ty",
                         color='Số chuyến', color_continuous_scale='Blues')
            fig.update_traces(textposition='outside')
            st.plotly_chart(fig, use_container_width=True)
            
        elif sel_bu == "Tất cả":
            # Level 2: Đã chọn Cty, chưa chọn BU -> Vẽ biểu đồ so sánh các BU
            st.info(f"Đang hiển thị so sánh các Phòng ban của {sel_comp}")
            df_g = df_filtered['BU'].value_counts().reset_index()
            df_g.columns = ['Phòng ban', 'Số chuyến']
            fig = px.bar(df_g, x='Số chuyến', y='Phòng ban', orientation='h', 
                         text='Số chuyến', title=f"Phòng ban thuộc {sel_comp}",
                         color='Số chuyến', color_continuous_scale='Teal')
            fig.update_traces(textposition='outside')
            st.plotly_chart(fig, use_container_width=True)
            
        else:
            # Level 3: Đã chọn cụ thể BU -> Vẽ biểu đồ User trong BU đó
            st.info(f"Đang hiển thị nhân sự của {sel_bu} ({sel_comp})")
            df_g = df_filtered['Người sử dụng xe'].value_counts().head(10).reset_index()
            df_g.columns = ['Nhân viên', 'Số chuyến']
            fig = px.bar(df_g, x='Số chuyến', y='Nhân viên', orientation='h', 
                         text='Số chuyến', title=f"Top nhân viên tại {sel_bu}",
                         color='Số chuyến', color_continuous_scale='Purples')
            fig.update_traces(textposition='outside')
            st.plotly_chart(fig, use_container_width=True)

    with t2:
        c_trend, c_rank = st.columns([2, 1])
        with c_trend:
            st.write("#### Xu hướng theo tháng")
            if 'Tháng' in df_filtered.columns:
                df_trend = df_filtered.groupby('Tháng').size().reset_index(name='Số chuyến')
                fig_line = px.line(df_trend, x='Tháng', y='Số chuyến', markers=True, text='Số chuyến')
                fig_line.update_traces(textposition="top center") # SỐ LIỆU TRÊN LINE
                st.plotly_chart(fig_line, use_container_width=True)
        
        with c_rank:
            st.write("#### 🏆 Bảng Xếp Hạng")
            tab_u, tab_d = st.tabs(["Người dùng", "Tài xế"])
            with tab_u:
                top_u = df_filtered['Người sử dụng xe'].value_counts().head(10).reset_index()
                top_u.columns = ['Tên', 'Chuyến']; st.dataframe(top_u, use_container_width=True, hide_index=True)
            with tab_d:
                if 'Tên tài xế' in df_filtered.columns:
                    top_d = df_filtered['Tên tài xế'].value_counts().head(10).reset_index()
                    top_d.columns = ['Tên', 'Chuyến']; st.dataframe(top_d, use_container_width=True, hide_index=True)

    with t3:
        c1, c2 = st.columns(2)
        with c1:
            st.write("#### Tỷ lệ Trạng thái")
            df_st = counts.reset_index()
            df_st.columns = ['Status', 'Count']
            # BIỂU ĐỒ TRÒN CÓ SỐ LIỆU
            fig_pie = px.pie(df_st, values='Count', names='Status', hole=0.4, 
                             color='Status',
                             color_discrete_map={'CLOSED':'#107c10', 'CANCELED':'#d13438', 'REJECTED_BY_ADMIN':'#a80000'})
            fig_pie.update_traces(textinfo='percent+label') 
            st.plotly_chart(fig_pie, use_container_width=True)
            
        with c2:
            st.write("#### Chi tiết Hủy/Từ chối")
            df_bad = df_filtered[df_filtered['Tình trạng đơn yêu cầu'].isin(['CANCELED', 'CANCELLED', 'REJECTED_BY_ADMIN'])]
            if not df_bad.empty:
                st.dataframe(df_bad[['Ngày khởi hành', 'Người sử dụng xe', 'Công ty', 'Tình trạng đơn yêu cầu', 'Note']], use_container_width=True)
            else:
                st.success("Không có chuyến nào bị Hủy hoặc Từ chối trong bộ lọc này.")

    # --- PPTX BUTTON ---
    st.markdown("---")
    kpi_exp = {'trips': total, 'hours': used, 'occupancy': occupancy, 'success_rate': suc_rate, 'cancel_rate': can_rate, 'reject_rate': rej_rate, 'last_month': df['Tháng'].max()}
    df_comp_exp = df_filtered['Công ty'].value_counts().reset_index(); df_comp_exp.columns=['Công ty', 'Số chuyến']
    df_status_exp = df_st
    
    pptx_data = export_pptx(kpi_exp, df_status_exp, df_comp_exp)
    st.download_button("📥 Tải Báo Cáo PPTX (Kèm Biểu Đồ)", pptx_data, "Bao_Cao_Van_Hanh.pptx", "application/vnd.openxmlformats-officedocument.presentationml.presentation", type="primary")

else:
    st.info("👋 Upload file Excel để bắt đầu.")