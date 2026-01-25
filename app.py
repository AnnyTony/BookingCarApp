import streamlit as st
import pandas as pd
import plotly.express as px
import matplotlib.pyplot as plt
from io import BytesIO
from pptx import Presentation
from pptx.util import Inches, Pt, Cm
from pptx.enum.text import PP_ALIGN
from pptx.dml.color import RGBColor

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
        
        if not sheet_booking: return "❌ Không tìm thấy sheet 'Booking car' (hoặc tên tương tự)."

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
            
            available_cols = [c for c in col_map.keys() if c in df_cbnv.columns]
            df_cbnv = df_cbnv[available_cols].rename(columns=col_map)
            
            if 'Full Name' in df_cbnv.columns:
                df_cbnv = df_cbnv.drop_duplicates(subset=['Full Name'], keep='first')
                df_final = df_final.merge(df_cbnv, left_on='Người sử dụng xe', right_on='Full Name', how='left')

        # Fillna & Format
        for c in ['Công ty', 'BU', 'Location']:
            if c not in df_final.columns: df_final[c] = 'Unknown'
            else: df_final[c] = df_final[c].fillna('Unknown').astype(str)
            
        # Xử lý ngày tháng
        df_final['Start'] = pd.to_datetime(df_final['Ngày khởi hành'].astype(str) + ' ' + df_final['Giờ khởi hành'].astype(str), errors='coerce')
        df_final['End'] = pd.to_datetime(df_final['Ngày khởi hành'].astype(str) + ' ' + df_final['Giờ kết thúc'].astype(str), errors='coerce')
        
        # Xử lý trường hợp qua đêm hoặc lỗi giờ
        df_final.loc[df_final['End'] < df_final['Start'], 'End'] += pd.Timedelta(days=1)
        
        df_final['Duration'] = (df_final['End'] - df_final['Start']).dt.total_seconds() / 3600
        df_final['Tháng'] = df_final['Start'].dt.strftime('%Y-%m')
        df_final['Năm'] = df_final['Start'].dt.year
        
        return df_final
    except Exception as e: return f"Lỗi xử lý dữ liệu: {str(e)}"

# --- 3. HÀM TẠO ẢNH CHO PPTX ---
def get_chart_img(data, x, y, kind='bar', title=''):
    plt.figure(figsize=(6, 4))
    if kind == 'bar':
        plt.barh(data[x], data[y], color='#0078d4')
        plt.xlabel(y)
        plt.gca().invert_yaxis() # Đảo ngược trục Y để cái cao nhất lên đầu
    elif kind == 'pie':
        plt.pie(data[y], labels=data[x], autopct='%1.1f%%', startangle=90)
    plt.title(title)
    plt.tight_layout()
    img = BytesIO(); plt.savefig(img, format='png', dpi=100); plt.close(); img.seek(0)
    return img

# --- 4. HÀM XUẤT PPTX (NÂNG CẤP) ---
def export_pptx(kpi, df_status, df_comp, df_bad_trips):
    prs = Presentation()
    
    # Hàm hỗ trợ tạo slide title nhanh
    def add_title_slide(title, subtitle):
        slide = prs.slides.add_slide(prs.slide_layouts[0])
        slide.shapes.title.text = title
        slide.placeholders[1].text = subtitle
        return slide

    # Slide 1: Title
    add_title_slide("BÁO CÁO VẬN HÀNH ĐỘI XE", f"Cập nhật đến: {kpi['last_month']}")
    
    # Slide 2: KPI Tổng quan
    slide2 = prs.slides.add_slide(prs.slide_layouts[1])
    slide2.shapes.title.text = "TỔNG QUAN HIỆU SUẤT"
    
    content_box = slide2.shapes.placeholders[1]
    tf = content_box.text_frame
    tf.text = f"Tổng số chuyến đi: {kpi['trips']} chuyến"
    p = tf.add_paragraph(); p.text = f"Tổng giờ vận hành: {kpi['hours']:,.0f} giờ"
    p = tf.add_paragraph(); p.text = f"Tỷ lệ lấp đầy (Occupancy): {kpi['occupancy']:.1f}%"
    p = tf.add_paragraph(); p.text = f"Tỷ lệ Hoàn thành: {kpi['success_rate']:.1f}%"
    p = tf.add_paragraph(); p.text = f"Tỷ lệ Hủy/Từ chối: {kpi['cancel_rate'] + kpi['reject_rate']:.1f}%"

    # Slide 3: Charts
    slide3 = prs.slides.add_slide(prs.slide_layouts[5]) # Title only
    slide3.shapes.title.text = "PHÂN BỔ THEO CÔNG TY & TRẠNG THÁI"
    
    # Chèn ảnh biểu đồ
    img1 = get_chart_img(df_comp.head(8), 'Công ty', 'Số chuyến', 'bar', 'Top Công ty sử dụng nhiều nhất')
    slide3.shapes.add_picture(img1, Inches(0.5), Inches(2), Inches(4.5), Inches(3.5))
    
    img2 = get_chart_img(df_status, 'Trạng thái', 'Số lượng', 'pie', 'Tỷ lệ trạng thái đơn')
    slide3.shapes.add_picture(img2, Inches(5.2), Inches(2), Inches(4.5), Inches(3.5))

    # Slide 4: Table Chi tiết Hủy/Từ chối (NEW)
    slide4 = prs.slides.add_slide(prs.slide_layouts[5])
    slide4.shapes.title.text = "CHI TIẾT ĐƠN HỦY / TỪ CHỐI (TOP 10)"
    
    if not df_bad_trips.empty:
        rows, cols = min(len(df_bad_trips)+1, 11), 4
        left, top, width, height = Inches(0.5), Inches(1.5), Inches(9), Inches(0.8)
        table = slide4.shapes.add_table(rows, cols, left, top, width, height).table
        
        # Set column widths
        table.columns[0].width = Inches(1.5) # Ngày
        table.columns[1].width = Inches(2.5) # User
        table.columns[2].width = Inches(2.0) # Status
        table.columns[3].width = Inches(3.0) # Lý do
        
        # Header
        headers = ['Ngày', 'Người dùng', 'Trạng thái', 'Ghi chú']
        for i, h in enumerate(headers):
            cell = table.cell(0, i)
            cell.text = h
            cell.fill.solid()
            cell.fill.fore_color.rgb = RGBColor(0, 120, 212) # Blue header
            cell.text_frame.paragraphs[0].font.color.rgb = RGBColor(255, 255, 255)
            cell.text_frame.paragraphs[0].font.bold = True
            
        # Rows
        for i, row in enumerate(df_bad_trips.head(10).itertuples(), start=1):
            table.cell(i, 0).text = str(row.Start_Str)
            table.cell(i, 1).text = str(row.User)
            table.cell(i, 2).text = str(row.Status)
            table.cell(i, 3).text = str(row.Note) if str(row.Note) != 'nan' else ""
            
    else:
        txBox = slide4.shapes.add_textbox(Inches(1), Inches(2), Inches(5), Inches(1))
        txBox.text_frame.text = "Tuyệt vời! Không có chuyến nào bị hủy hoặc từ chối trong giai đoạn này."

    out = BytesIO(); prs.save(out); out.seek(0)
    return out

# --- 5. GIAO DIỆN CHÍNH ---
st.title("📊 Fleet Management Pro")
uploaded_file = st.file_uploader("Upload Excel", type=['xlsx'], label_visibility="collapsed")

if uploaded_file:
    df = load_data_final(uploaded_file)
    if isinstance(df, str): st.error(df); st.stop()
    
    # --- SIDEBAR FILTERS ---
    with st.sidebar:
        st.header("🗂️ Bộ Lọc Dữ Liệu")
        
        # 1. Date Filter (MỚI)
        min_date = df['Start'].min().date()
        max_date = df['Start'].max().date()
        
        date_range = st.date_input("Khoảng thời gian:", value=(min_date, max_date), min_value=min_date, max_value=max_date)
        
        # Logic lọc ngày
        if isinstance(date_range, tuple) and len(date_range) == 2:
            start_d, end_d = date_range
            df_date_filtered = df[(df['Start'].dt.date >= start_d) & (df['Start'].dt.date <= end_d)]
        else:
            df_date_filtered = df

        st.markdown("---")
        
        # 2. Hierarchy Filter
        st.caption("Lọc theo tổ chức:")
        locs = ["Tất cả"] + sorted(df_date_filtered['Location'].unique().tolist())
        sel_loc = st.selectbox("1. Khu vực (Region):", locs)
        
        df_l1 = df_date_filtered if sel_loc == "Tất cả" else df_date_filtered[df_date_filtered['Location'] == sel_loc]
        
        comps = ["Tất cả"] + sorted(df_l1['Công ty'].unique().tolist())
        sel_comp = st.selectbox("2. Công ty (Entity):", comps)
        
        df_l2 = df_l1 if sel_comp == "Tất cả" else df_l1[df_l1['Công ty'] == sel_comp]
        
        bus = ["Tất cả"] + sorted(df_l2['BU'].unique().tolist())
        sel_bu = st.selectbox("3. Phòng ban (BU):", bus)
        
        df_filtered = df_l2 if sel_bu == "Tất cả" else df_l2[df_l2['BU'] == sel_bu]
        
        st.markdown("---")
        st.write(f"Đang xem: **{len(df_filtered)}** chuyến")

    # --- KPI CALCULATION ---
    if df_filtered.empty:
        st.warning("⚠️ Không có dữ liệu cho bộ lọc này.")
        st.stop()

    total_cars = 21
    if 'HCM' in sel_loc or 'NAM' in sel_loc.upper(): total_cars = 16
    elif 'HN' in sel_loc or 'BAC' in sel_loc.upper(): total_cars = 5
    
    # Tính occupancy dựa trên ngày thực tế lọc
    days_in_filter = (df_filtered['Start'].max() - df_filtered['Start'].min()).days + 1
    days_in_filter = max(days_in_filter, 1)
    
    cap = total_cars * days_in_filter * 9
    used = df_filtered['Duration'].sum()
    occupancy = (used / cap * 100) if cap > 0 else 0
    
    counts = df_filtered['Tình trạng đơn yêu cầu'].fillna('Unknown').value_counts()
    total = len(df_filtered)
    cancel = counts.get('CANCELED', 0) + counts.get('CANCELLED', 0)
    reject = counts.get('REJECTED_BY_ADMIN', 0)
    completed = counts.get('CLOSED', 0) + counts.get('APPROVED', 0)
    
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
    t1, t2, t3 = st.tabs(["📊 Phân Tích Đơn Vị", "📈 Xu Hướng & Top", "📉 Chất Lượng Vận Hành"])
    
    with t1:
        c_left, c_right = st.columns([2, 1])
        with c_left:
            st.write("#### Phân tích theo Cấu trúc")
            if sel_comp == "Tất cả":
                df_g = df_filtered['Công ty'].value_counts().reset_index()
                df_g.columns = ['Công ty', 'Số chuyến']
                fig = px.bar(df_g, x='Số chuyến', y='Công ty', orientation='h', 
                             text='Số chuyến', title="Số chuyến theo Công ty",
                             color='Số chuyến', color_continuous_scale='Blues')
                fig.update_traces(textposition='outside')
                st.plotly_chart(fig, use_container_width=True)
            elif sel_bu == "Tất cả":
                df_g = df_filtered['BU'].value_counts().reset_index()
                df_g.columns = ['Phòng ban', 'Số chuyến']
                fig = px.bar(df_g, x='Số chuyến', y='Phòng ban', orientation='h', 
                             text='Số chuyến', title=f"Phòng ban thuộc {sel_comp}",
                             color='Số chuyến', color_continuous_scale='Teal')
                st.plotly_chart(fig, use_container_width=True)
            else:
                df_g = df_filtered['Người sử dụng xe'].value_counts().head(10).reset_index()
                df_g.columns = ['Nhân viên', 'Số chuyến']
                fig = px.bar(df_g, x='Số chuyến', y='Nhân viên', orientation='h', 
                             text='Số chuyến', title=f"Top nhân viên tại {sel_bu}",
                             color='Số chuyến', color_continuous_scale='Purples')
                st.plotly_chart(fig, use_container_width=True)
        with c_right:
             # Phạm vi di chuyển
            st.write("#### Phạm vi di chuyển")
            if 'Phạm Vi' in df_filtered.columns:
                 df_scope = df_filtered['Phạm Vi'].value_counts().reset_index()
                 df_scope.columns = ['Phạm vi', 'Số lượng']
                 fig_scope = px.pie(df_scope, values='Số lượng', names='Phạm vi', hole=0.5)
                 st.plotly_chart(fig_scope, use_container_width=True)

    with t2:
        c_trend, c_rank = st.columns([2, 1])
        with c_trend:
            st.write("#### Xu hướng theo thời gian")
            # Group by ngày hoặc tháng tùy theo filter
            if days_in_filter <= 31:
                 df_filtered['Date_Only'] = df_filtered['Start'].dt.date
                 df_trend = df_filtered.groupby('Date_Only').size().reset_index(name='Số chuyến')
                 x_axis = 'Date_Only'
            else:
                 df_trend = df_filtered.groupby('Tháng').size().reset_index(name='Số chuyến')
                 x_axis = 'Tháng'
                 
            fig_line = px.line(df_trend, x=x_axis, y='Số chuyến', markers=True, text='Số chuyến')
            fig_line.update_traces(textposition="top center")
            st.plotly_chart(fig_line, use_container_width=True)
        
        with c_rank:
            st.write("#### 🏆 Top Users")
            top_u = df_filtered['Người sử dụng xe'].value_counts().head(10).reset_index()
            top_u.columns = ['Tên', 'Chuyến']
            st.dataframe(top_u, use_container_width=True, hide_index=True)

    with t3:
        c1, c2 = st.columns(2)
        with c1:
            st.write("#### Tỷ lệ Trạng thái")
            df_st = counts.reset_index()
            df_st.columns = ['Status', 'Count']
            fig_pie = px.pie(df_st, values='Count', names='Status', hole=0.4, 
                             color='Status',
                             color_discrete_map={'CLOSED':'#107c10', 'CANCELED':'#d13438', 'REJECTED_BY_ADMIN':'#a80000'})
            fig_pie.update_traces(textinfo='percent+label') 
            st.plotly_chart(fig_pie, use_container_width=True)
            
        with c2:
            st.write("#### Danh sách Hủy/Từ chối")
            df_bad = df_filtered[df_filtered['Tình trạng đơn yêu cầu'].isin(['CANCELED', 'CANCELLED', 'REJECTED_BY_ADMIN'])]
            if not df_bad.empty:
                show_cols = ['Ngày khởi hành', 'Người sử dụng xe', 'Công ty', 'Tình trạng đơn yêu cầu', 'Note']
                # Lọc cột tồn tại
                actual_cols = [c for c in show_cols if c in df_bad.columns]
                st.dataframe(df_bad[actual_cols], use_container_width=True)
            else:
                st.success("Không có chuyến nào bị Hủy hoặc Từ chối trong bộ lọc này.")

    # --- PPTX BUTTON ---
    st.markdown("---")
    
    # Chuẩn bị dữ liệu cho PPTX
    # SỬA LỖI TẠI ĐÂY: Thêm check not empty cho df
    last_month_str = "N/A"
    if not df.empty and 'Tháng' in df.columns:
        valid_months = df['Tháng'].dropna()
        if not valid_months.empty:
            last_month_str = valid_months.max()

    kpi_exp = {
        'trips': total, 'hours': used, 'occupancy': occupancy, 
        'success_rate': suc_rate, 'cancel_rate': can_rate, 
        'reject_rate': rej_rate, 
        'last_month': last_month_str # Đã fix lỗi
    }
    
    df_comp_exp = df_filtered['Công ty'].value_counts().reset_index()
    df_comp_exp.columns=['Công ty', 'Số chuyến']
    
    df_status_exp = counts.reset_index()
    df_status_exp.columns = ['Trạng thái', 'Số lượng']
    
    # Chuẩn bị data cho slide bảng chi tiết (Đổi tên cột cho đẹp)
    df_bad_export = pd.DataFrame()
    if not df_bad.empty:
        df_bad_export = df_bad.copy()
        df_bad_export['Start_Str'] = df_bad_export['Start'].dt.strftime('%d/%m/%Y')
        df_bad_export = df_bad_export.rename(columns={'Người sử dụng xe': 'User', 'Tình trạng đơn yêu cầu': 'Status'})
        
    
    pptx_data = export_pptx(kpi_exp, df_status_exp, df_comp_exp, df_bad_export)
    
    col_dl1, col_dl2 = st.columns([1, 4])
    with col_dl1:
        st.download_button(
            "📥 Tải Báo Cáo PPTX", 
            pptx_data, 
            "Bao_Cao_Van_Hanh.pptx", 
            "application/vnd.openxmlformats-officedocument.presentationml.presentation", 
            type="primary"
        )
    with col_dl2:
        st.caption("💡 Báo cáo PPTX đã bao gồm biểu đồ và danh sách chi tiết các chuyến bị hủy/từ chối.")

else:
    st.info("👋 Upload file Excel để bắt đầu.")