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
    
    /* KPI Card Style */
    .kpi-card {
        background-color: white; border-radius: 8px; padding: 15px;
        box-shadow: 0 2px 4px rgba(0,0,0,0.08); border-left: 5px solid #0078d4;
        margin-bottom: 10px;
    }
    .kpi-title {font-size: 13px; color: #666; font-weight: 600; text-transform: uppercase;}
    .kpi-value {font-size: 28px; font-weight: 700; color: #333; margin: 5px 0;}
    .kpi-sub {font-size: 11px; color: #28a745; font-weight: 500;}
    
    /* Breadcrumb Style */
    .breadcrumb {
        font-size: 16px; color: #0078d4; font-weight: 600; 
        background-color: #f0f2f6; padding: 10px; border-radius: 5px;
        margin-bottom: 20px;
    }
</style>
""", unsafe_allow_html=True)

# --- 2. HÀM XỬ LÝ DỮ LIỆU ---
@st.cache_data
def load_data_final(file):
    try:
        xl = pd.ExcelFile(file, engine='openpyxl')
        
        sheet_driver = next((s for s in xl.sheet_names if 'driver' in s.lower()), None)
        sheet_booking = next((s for s in xl.sheet_names if 'booking' in s.lower()), None)
        sheet_cbnv = next((s for s in xl.sheet_names if 'cbnv' in s.lower()), None)
        
        if not sheet_booking: return "❌ Không tìm thấy sheet 'Booking car'."

        def smart_read(excel, sheet_name, keywords):
            df_preview = excel.parse(sheet_name, header=None, nrows=10)
            header_idx = 0
            for idx, row in df_preview.iterrows():
                row_str = row.astype(str).str.lower().tolist()
                if any(k in row_str for k in keywords):
                    header_idx = idx; break
            return excel.parse(sheet_name, header=header_idx)

        df_bk = smart_read(xl, sheet_booking, ['ngày khởi hành'])
        df_driver = smart_read(xl, sheet_driver, ['biển số xe']) if sheet_driver else pd.DataFrame()
        df_cbnv = smart_read(xl, sheet_cbnv, ['full name']) if sheet_cbnv else pd.DataFrame()

        df_bk.columns = df_bk.columns.str.strip()
        
        df_final = df_bk
        if not df_driver.empty:
            df_driver.columns = df_driver.columns.str.strip()
            if 'Biển số xe' in df_driver.columns:
                df_driver = df_driver.drop_duplicates(subset=['Biển số xe'], keep='last')
                df_final = df_final.merge(df_driver[['Biển số xe', 'Tên tài xế']], on='Biển số xe', how='left', suffixes=('', '_D'))
                if 'Tên tài xế_D' in df_final.columns:
                    df_final['Tên tài xế'] = df_final['Tên tài xế'].fillna(df_final['Tên tài xế_D'])

        if not df_cbnv.empty:
            df_cbnv.columns = df_cbnv.columns.str.strip()
            col_map = {}
            for c in df_cbnv.columns:
                if 'full name' in str(c).lower(): col_map[c] = 'Full Name'
                if 'công ty' in str(c).lower(): col_map[c] = 'Công ty'
                if 'bu' in str(c).lower(): col_map[c] = 'BU'
                if 'location' in str(c).lower(): col_map[c] = 'Location'
            
            available_cols = [c for c in col_map.keys() if c in df_cbnv.columns]
            df_cbnv = df_cbnv[available_cols].rename(columns=col_map)
            
            if 'Full Name' in df_cbnv.columns:
                df_cbnv = df_cbnv.drop_duplicates(subset=['Full Name'], keep='first')
                df_final = df_final.merge(df_cbnv, left_on='Người sử dụng xe', right_on='Full Name', how='left')

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
def export_pptx(kpi, df_status, df_breakdown, breakdown_col):
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
    
    # Slide 3: Chart Breakdown (Dynamic)
    slide = prs.slides.add_slide(prs.slide_layouts[5])
    slide.shapes.title.text = f"Phân Bổ Theo {breakdown_col}"
    
    # Vẽ biểu đồ động theo cột breakdown hiện tại
    img1 = get_chart_img(df_breakdown.head(10), breakdown_col, 'Số chuyến', 'bar', f'Top {breakdown_col}')
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
    
    # --- LOGIC CÂY THƯ MỤC CASCADING (DRILL-DOWN) ---
    with st.sidebar:
        st.header("🗂️ Phân Cấp Dữ Liệu")
        st.info("Bộ lọc này hoạt động theo cơ chế Cha -> Con (Drill-down)")
        
        # Level 1: Location (Region)
        loc_opts = ["Tất cả"] + sorted(df['Location'].unique().tolist())
        sel_loc = st.selectbox("1. Khu vực (Region):", loc_opts)
        
        # Filter Level 1
        if sel_loc == "Tất cả":
            df_lv1 = df
            current_breakdown = "Location" # Nếu chọn tất cả vùng, biểu đồ sẽ so sánh các Vùng
            drill_status = "Toàn quốc"
        else:
            df_lv1 = df[df['Location'] == sel_loc]
            current_breakdown = "Công ty" # Nếu chọn 1 vùng, biểu đồ sẽ so sánh các Công ty trong vùng đó
            drill_status = f"{sel_loc}"

        # Level 2: Company (Entity) - Options depend on Level 1
        comp_opts = ["Tất cả"] + sorted(df_lv1['Công ty'].unique().tolist())
        sel_comp = st.selectbox("2. Công ty (Entity):", comp_opts)
        
        # Filter Level 2
        if sel_comp == "Tất cả":
            df_lv2 = df_lv1
            # Giữ nguyên breakdown là Công ty
        else:
            df_lv2 = df_lv1[df_lv1['Công ty'] == sel_comp]
            current_breakdown = "BU" # Nếu chọn 1 Cty, biểu đồ so sánh các BU
            drill_status += f" > {sel_comp}"

        # Level 3: BU (Department) - Options depend on Level 2
        bu_opts = ["Tất cả"] + sorted(df_lv2['BU'].unique().tolist())
        sel_bu = st.selectbox("3. Phòng ban (BU):", bu_opts)
        
        # Filter Level 3
        if sel_bu == "Tất cả":
            df_final = df_lv2
        else:
            df_final = df_lv2[df_lv2['BU'] == sel_bu]
            current_breakdown = "Người sử dụng xe" # Nếu chọn 1 BU, biểu đồ so sánh Nhân viên
            drill_status += f" > {sel_bu}"
        
        st.markdown("---")
        st.caption(f"Đang xem: **{len(df_final)}** chuyến")

    # --- BREADCRUMB & KPI ---
    st.markdown(f"<div class='breadcrumb'>📍 Đang xem: {drill_status}</div>", unsafe_allow_html=True)

    # Tính toán KPI
    total_cars = 21
    # Logic xe thông minh theo vùng đang chọn
    if sel_loc != "Tất cả":
        if 'HCM' in str(sel_loc) or 'NAM' in str(sel_loc).upper(): total_cars = 16
        elif 'HN' in str(sel_loc) or 'BAC' in str(sel_loc).upper(): total_cars = 5
    
    days = (df['Start'].max() - df['Start'].min()).days + 1 if not df.empty else 1
    cap = total_cars * max(days, 1) * 9
    used = df_final['Duration'].sum()
    occupancy = (used / cap * 100) if cap > 0 else 0
    
    # Status Rates
    counts = df_final['Tình trạng đơn yêu cầu'].fillna('Unknown').value_counts()
    total_trips = len(df_final)
    cancel = counts.get('CANCELED', 0) + counts.get('CANCELLED', 0)
    reject = counts.get('REJECTED_BY_ADMIN', 0)
    completed = counts.get('CLOSED', 0) + counts.get('APPROVED', 0)
    
    suc_rate = (completed / total_trips * 100) if total_trips > 0 else 0
    can_rate = (cancel / total_trips * 100) if total_trips > 0 else 0
    rej_rate = (reject / total_trips * 100) if total_trips > 0 else 0

    # KPI UI
    k1, k2, k3, k4, k5 = st.columns(5)
    k1.markdown(f"<div class='kpi-card'><div class='kpi-title'>Tổng Chuyến</div><div class='kpi-value'>{total_trips}</div></div>", unsafe_allow_html=True)
    k2.markdown(f"<div class='kpi-card'><div class='kpi-title'>Giờ Vận Hành</div><div class='kpi-value'>{used:,.0f}</div></div>", unsafe_allow_html=True)
    k3.markdown(f"<div class='kpi-card'><div class='kpi-title'>Occupancy</div><div class='kpi-value'>{occupancy:.1f}%</div><div class='kpi-sub'>Trên {total_cars} xe</div></div>", unsafe_allow_html=True)
    k4.markdown(f"<div class='kpi-card' style='border-left: 5px solid #107c10'><div class='kpi-title'>Hoàn Thành</div><div class='kpi-value' style='color:#107c10'>{suc_rate:.1f}%</div></div>", unsafe_allow_html=True)
    k5.markdown(f"<div class='kpi-card' style='border-left: 5px solid #d13438'><div class='kpi-title'>Hủy / Từ Chối</div><div class='kpi-value' style='color:#d13438'>{can_rate + rej_rate:.1f}%</div></div>", unsafe_allow_html=True)

    # --- DYNAMIC CHART SECTION ---
    st.markdown("---")
    
    c_main, c_trend = st.columns([2, 1])
    
    with c_main:
        # Tự động thay đổi tiêu đề và dữ liệu biểu đồ dựa trên cấp độ Drill-down
        st.markdown(f"<div class='chart-header'>📊 Phân bổ theo {current_breakdown}</div>", unsafe_allow_html=True)
        
        # Prepare Data for Main Chart
        df_agg = df_final[current_breakdown].value_counts().reset_index().head(15) # Top 15 items
        df_agg.columns = [current_breakdown, 'Số chuyến']
        
        # Cho phép user chỉnh loại biểu đồ
        chart_type = st.radio("Loại biểu đồ:", ["Cột (Bar)", "Tròn (Pie)"], horizontal=True, label_visibility="collapsed")
        
        if "Cột" in chart_type:
            fig = px.bar(df_agg, x='Số chuyến', y=current_breakdown, orientation='h', 
                         text='Số chuyến', color='Số chuyến', color_continuous_scale='Blues')
            fig.update_traces(textposition='outside')
            fig.update_layout(yaxis={'categoryorder':'total ascending'})
        else:
            fig = px.pie(df_agg, values='Số chuyến', names=current_breakdown, hole=0.4)
            fig.update_traces(textinfo='percent+label')
            
        st.plotly_chart(fig, use_container_width=True)

    with c_trend:
        st.markdown(f"<div class='chart-header'>📈 Xu hướng (Tại {drill_status})</div>", unsafe_allow_html=True)
        if 'Tháng' in df_final.columns:
            df_trend = df_final.groupby('Tháng').size().reset_index(name='Số chuyến')
            fig_trend = px.area(df_trend, x='Tháng', y='Số chuyến', markers=True)
            fig_trend.update_layout(height=400)
            st.plotly_chart(fig_trend, use_container_width=True)
        else:
            st.info("Chưa có dữ liệu tháng.")

    # --- TOP LISTS ---
    st.markdown("---")
    st.markdown("<div class='chart-header'>🏆 Bảng Xếp Hạng Chi Tiết</div>", unsafe_allow_html=True)
    
    t1, t2, t3 = st.columns(3)
    with t1:
        st.write("**Top Tài xế**")
        if 'Tên tài xế' in df_final.columns:
            top_d = df_final['Tên tài xế'].value_counts().head(10).reset_index(name='Số chuyến')
            st.dataframe(top_d, use_container_width=True, hide_index=True)
            
    with t2:
        st.write("**Top Người dùng**")
        top_u = df_final['Người sử dụng xe'].value_counts().head(10).reset_index(name='Số chuyến')
        st.dataframe(top_u, use_container_width=True, hide_index=True)
        
    with t3:
        st.write("**Chất lượng (Cancel/Reject)**")
        df_st = counts.reset_index(name='Số lượng')
        fig_st = px.pie(df_st, values='Số lượng', names='index', 
                        color='index',
                        color_discrete_map={'CLOSED':'#107c10', 'CANCELED':'#d13438', 'REJECTED_BY_ADMIN':'#a80000'})
        st.plotly_chart(fig_st, use_container_width=True)

    # --- PPTX DOWNLOAD ---
    st.markdown("---")
    # Prepare export data based on current view
    kpi_exp = {'trips': total_trips, 'hours': used, 'occupancy': occupancy, 'success_rate': suc_rate, 'cancel_rate': can_rate, 'reject_rate': rej_rate, 'last_month': df['Tháng'].max()}
    
    # Export Dynamic Chart Data
    df_breakdown_exp = df_final[current_breakdown].value_counts().reset_index()
    df_breakdown_exp.columns = [current_breakdown, 'Số chuyến']
    
    df_status_exp = df_st
    df_status_exp.columns = ['Trạng thái', 'Số lượng'] # Rename for safety
    
    pptx_data = export_pptx(kpi_exp, df_status_exp, df_breakdown_exp, current_breakdown)
    
    st.download_button(f"📥 Tải Báo Cáo PPTX (Góc nhìn: {current_breakdown})", pptx_data, "Bao_Cao_Van_Hanh.pptx", "application/vnd.openxmlformats-officedocument.presentationml.presentation", type="primary")

else:
    st.info("👋 Upload file Excel để bắt đầu.")