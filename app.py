import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import matplotlib.pyplot as plt
from io import BytesIO
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.enum.text import PP_ALIGN

# --- 1. CẤU HÌNH TRANG & CSS PRO ---
st.set_page_config(page_title="Fleet Intelligence Hub", page_icon="📊", layout="wide")

# CSS giả lập giao diện Dashboard chuyên nghiệp
st.markdown("""
<style>
    /* Tổng thể */
    .block-container {padding-top: 1rem; padding-bottom: 3rem;}
    
    /* Sidebar */
    [data-testid="stSidebar"] {
        background-color: #f8f9fa;
        border-right: 1px solid #dee2e6;
    }
    
    /* KPI Cards - Power BI Style */
    .kpi-card {
        background-color: white;
        border-radius: 8px;
        padding: 15px 20px;
        box-shadow: 0 2px 5px rgba(0,0,0,0.05);
        border-left: 5px solid #0078d4; /* Màu xanh Power BI */
        margin-bottom: 10px;
    }
    .kpi-title {font-size: 13px; color: #605e5c; font-weight: 600; text-transform: uppercase;}
    .kpi-value {font-size: 32px; font-weight: 700; color: #201f1e; margin: 5px 0;}
    .kpi-note {font-size: 11px; color: #8a8886;}
    
    /* Section Headers */
    .section-title {
        font-size: 18px; font-weight: 700; color: #0078d4;
        margin-top: 20px; margin-bottom: 10px;
        display: flex; align-items: center;
    }
    .section-title::before {
        content: ""; display: inline-block; width: 6px; height: 24px;
        background-color: #0078d4; margin-right: 10px; border-radius: 2px;
    }
    
    /* Tabs */
    .stTabs [data-baseweb="tab-list"] {gap: 5px;}
    .stTabs [data-baseweb="tab"] {
        height: 40px; background-color: white; border-radius: 4px 4px 0 0;
        box-shadow: none; border: 1px solid #e1dfdd;
    }
    .stTabs [aria-selected="true"] {
        background-color: #eff6fc; color: #0078d4; border-bottom: 2px solid #0078d4;
    }
</style>
""", unsafe_allow_html=True)

# --- 2. HÀM XỬ LÝ DỮ LIỆU (Giữ nguyên logic thông minh) ---
@st.cache_data
def load_data_pro(file):
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

        # Clean
        df_bk.columns = df_bk.columns.str.strip()
        if not df_driver.empty:
            df_driver.columns = df_driver.columns.str.strip()
            df_driver = df_driver.drop_duplicates(subset=['Biển số xe'], keep='last')
        if not df_cbnv.empty:
            df_cbnv.columns = df_cbnv.columns.str.strip()
            # Map cột CBNV
            col_map = {}
            for c in df_cbnv.columns:
                if 'full name' in str(c).lower(): col_map[c] = 'Full Name'
                if 'công ty' in str(c).lower(): col_map[c] = 'Công ty'
                if 'bu' in str(c).lower(): col_map[c] = 'BU'
                if 'location' in str(c).lower(): col_map[c] = 'Location'
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

        # Calculate
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

# --- 3. HÀM TẠO HÌNH ẢNH BIỂU ĐỒ CHO PPTX (Dùng Matplotlib) ---
def generate_chart_image(data, x_col, y_col, kind='bar', title='Chart'):
    plt.figure(figsize=(6, 4))
    if kind == 'bar':
        plt.barh(data[x_col], data[y_col], color='#0078d4')
        plt.xlabel(y_col)
    elif kind == 'pie':
        plt.pie(data[y_col], labels=data[x_col], autopct='%1.1f%%', startangle=90, colors=plt.cm.Paired.colors)
    
    plt.title(title)
    plt.tight_layout()
    
    img_stream = BytesIO()
    plt.savefig(img_stream, format='png', dpi=100)
    plt.close()
    img_stream.seek(0)
    return img_stream

# --- 4. HÀM XUẤT PPTX ---
def create_pptx_pro(kpi_data, df_comp, df_status, df_loc):
    prs = Presentation()
    
    # 1. Slide Tiêu đề
    slide = prs.slides.add_slide(prs.slide_layouts[0])
    slide.shapes.title.text = "Báo Cáo Vận Hành Đội Xe"
    slide.placeholders[1].text = f"Dữ liệu tính đến tháng {kpi_data['last_month']}"

    # 2. Slide KPI
    slide = prs.slides.add_slide(prs.slide_layouts[1])
    slide.shapes.title.text = "Tổng Quan Hiệu Suất"
    content = slide.placeholders[1]
    content.text = f"""
    • Tổng số chuyến đi: {kpi_data['total_trips']}
    • Tổng giờ vận hành: {kpi_data['total_hours']:,.0f} giờ
    • Tỷ lệ lấp đầy (Occupancy): {kpi_data['occupancy']:.1f}%
    • Tỷ lệ Hủy/Từ chối: {kpi_data['cancel_rate'] + kpi_data['reject_rate']:.1f}%
    """

    # 3. Slide Biểu đồ Công ty (Có hình)
    slide = prs.slides.add_slide(prs.slide_layouts[5]) # Blank layout
    slide.shapes.title.text = "Phân Bổ Theo Công Ty"
    
    # Tạo hình biểu đồ
    img_stream = generate_chart_image(df_comp.head(10), 'Công ty', 'Số chuyến', 'bar', 'Top 10 Công Ty')
    slide.shapes.add_picture(img_stream, Inches(0.5), Inches(2), Inches(5), Inches(3.5))
    
    # Tạo bảng bên cạnh
    table = slide.shapes.add_table(min(len(df_comp), 10)+1, 2, Inches(6), Inches(2), Inches(3), Inches(3.5)).table
    table.cell(0, 0).text = "Công ty"; table.cell(0, 1).text = "Số chuyến"
    for i, row in enumerate(df_comp.head(10).itertuples(index=False)):
        table.cell(i+1, 0).text = str(row[0])
        table.cell(i+1, 1).text = str(row[1])

    # 4. Slide Trạng thái & Phạm vi
    slide = prs.slides.add_slide(prs.slide_layouts[5])
    slide.shapes.title.text = "Chất Lượng & Phạm Vi"
    
    # Hình Status (Pie)
    img_status = generate_chart_image(df_status, 'Trạng thái', 'Số lượng', 'pie', 'Trạng thái Chuyến')
    slide.shapes.add_picture(img_status, Inches(0.5), Inches(2), Inches(4), Inches(3))
    
    # Hình Phạm vi (Pie)
    img_loc = generate_chart_image(df_loc, 'Phạm Vi', 'Số chuyến', 'pie', 'Nội thành vs Đi Tỉnh')
    slide.shapes.add_picture(img_loc, Inches(5), Inches(2), Inches(4), Inches(3))

    # Save
    buffer = BytesIO()
    prs.save(buffer)
    buffer.seek(0)
    return buffer

# --- 5. GIAO DIỆN CHÍNH ---
st.title("📊 Fleet Intelligence Hub")

# Upload (Ẩn trong Expander cho gọn)
with st.expander("📂 QUẢN LÝ DỮ LIỆU ĐẦU VÀO", expanded=True):
    uploaded_file = st.file_uploader("Upload file Excel", type=['xlsx'], label_visibility="collapsed")

if uploaded_file:
    df = load_data_pro(uploaded_file)
    if isinstance(df, str): st.error(df); st.stop()
    
    # --- SIDEBAR: BỘ LỌC CHUYÊN NGHIỆP ---
    with st.sidebar:
        st.markdown("### 🌪️ BỘ LỌC")
        
        # Lọc Năm
        years = sorted(df['Năm'].dropna().unique().astype(int))
        sel_year = st.multiselect("Năm", years, default=years)
        df = df[df['Năm'].isin(sel_year)]

        # Lọc Vùng (Cascading)
        locs = sorted(df['Location'].unique())
        sel_loc = st.multiselect("Khu vực", locs, default=locs)
        df = df[df['Location'].isin(sel_loc)]
        
        # Lọc Công ty
        comps = sorted(df['Công ty'].unique())
        sel_comp = st.multiselect("Công ty", comps, default=comps)
        df = df[df['Công ty'].isin(sel_comp)]
        
        # Lọc Bộ phận
        bus = sorted(df['BU'].unique())
        sel_bu = st.multiselect("Bộ phận", bus, default=bus)
        df = df[df['BU'].isin(sel_bu)]
        
        st.markdown("---")
        st.caption(f"Đang hiển thị: **{len(df)}** chuyến")

    # --- KPI SECTION (GRID LAYOUT) ---
    # Tính toán
    total_cars = 21
    if len(sel_loc) == 1:
        if 'HCM' in str(sel_loc[0]) or 'NAM' in str(sel_loc[0]).upper(): total_cars = 16
        elif 'HN' in str(sel_loc[0]) or 'BAC' in str(sel_loc[0]).upper(): total_cars = 5
        
    days = (df['Start'].max() - df['Start'].min()).days + 1 if not df.empty else 1
    cap_hours = total_cars * max(days, 1) * 9
    used_hours = df['Duration'].sum()
    occupancy = (used_hours / cap_hours * 100) if cap_hours > 0 else 0
    
    # Tính Cancel rate
    counts = df['Tình trạng đơn yêu cầu'].fillna('Unknown').value_counts()
    cancel_cnt = counts.get('CANCELED', 0) + counts.get('CANCELLED', 0)
    reject_cnt = counts.get('REJECTED_BY_ADMIN', 0)
    cancel_rate = (cancel_cnt / len(df) * 100) if len(df) > 0 else 0
    reject_rate = (reject_cnt / len(df) * 100) if len(df) > 0 else 0

    # Hiển thị KPI Cards
    k1, k2, k3, k4 = st.columns(4)
    k1.markdown(f"<div class='kpi-card'><div class='kpi-title'>Tổng Chuyến</div><div class='kpi-value'>{len(df)}</div></div>", unsafe_allow_html=True)
    k2.markdown(f"<div class='kpi-card'><div class='kpi-title'>Tổng Giờ Vận Hành</div><div class='kpi-value'>{used_hours:,.0f}</div></div>", unsafe_allow_html=True)
    k3.markdown(f"<div class='kpi-card'><div class='kpi-title'>Tỷ Lệ Lấp Đầy</div><div class='kpi-value'>{occupancy:.1f}%</div><div class='kpi-note'>({total_cars} xe * {days} ngày * 9h)</div></div>", unsafe_allow_html=True)
    k4.markdown(f"<div class='kpi-card'><div class='kpi-title'>Tỷ Lệ Hủy/Từ Chối</div><div class='kpi-value' style='color:#d13438'>{cancel_rate + reject_rate:.1f}%</div></div>", unsafe_allow_html=True)

    # --- MAIN DASHBOARD ---
    st.markdown("<div class='section-title'>PHÂN TÍCH HIỆU SUẤT</div>", unsafe_allow_html=True)
    
    # Tab chính
    tab_overview, tab_struct, tab_rank = st.tabs(["📊 Tổng Quan & Xu Hướng", "🏢 Cấu Trúc Đơn Vị", "🏆 Xếp Hạng Top"])
    
    with tab_overview:
        c1, c2 = st.columns([2, 1])
        with c1:
            st.write("**Xu hướng theo tháng**")
            if 'Tháng' in df.columns:
                df_trend = df.groupby('Tháng').size().reset_index(name='Số chuyến')
                fig_trend = px.area(df_trend, x='Tháng', y='Số chuyến', markers=True, 
                                    line_shape='spline', color_discrete_sequence=['#0078d4'])
                fig_trend.update_layout(height=350, margin=dict(l=20, r=20, t=20, b=20))
                st.plotly_chart(fig_trend, use_container_width=True)
        
        with c2:
            st.write("**Chất lượng vận hành**")
            df_status = counts.reset_index()
            df_status.columns = ['Trạng thái', 'Số lượng']
            fig_pie = px.pie(df_status, values='Số lượng', names='Trạng thái', hole=0.6,
                             color='Trạng thái',
                             color_discrete_map={'CLOSED':'#107c10', 'CANCELED':'#d13438', 'REJECTED_BY_ADMIN':'#a80000'})
            fig_pie.update_layout(height=350, showlegend=False)
            st.plotly_chart(fig_pie, use_container_width=True)
            
        # Thêm biểu đồ tùy chọn
        st.write("**Phân tích tùy chọn**")
        opt_col, opt_chart = st.columns([1, 3])
        with opt_col:
            dim = st.selectbox("Phân tích theo:", ["Công ty", "Phạm Vi", "Loại Chuyến"])
            chart_kind = st.selectbox("Loại biểu đồ:", ["Bar (Cột)", "Pie (Tròn)", "Sunburst (Phân cấp)"])
        with opt_chart:
            if chart_kind == "Sunburst (Phân cấp)" and dim == "Công ty":
                 fig_sun = px.sunburst(df, path=['Vùng Miền' if 'Vùng Miền' in df else 'Location', 'Công ty', 'BU'], title="Phân cấp Vùng -> Công ty -> BU")
                 st.plotly_chart(fig_sun, use_container_width=True)
            else:
                df_agg = df[dim].value_counts().reset_index()
                df_agg.columns = [dim, 'Số lượng']
                if "Bar" in chart_kind:
                    fig_opt = px.bar(df_agg, x=dim, y='Số lượng', text='Số lượng', color='Số lượng')
                else:
                    fig_opt = px.pie(df_agg, names=dim, values='Số lượng', hole=0.4)
                st.plotly_chart(fig_opt, use_container_width=True)

    with tab_struct:
        st.write("**Biểu đồ phân cấp (Sunburst)**")
        st.info("Click vào vòng tròn trung tâm để mở rộng chi tiết.")
        # Gom nhóm cho Sunburst
        if not df.empty:
            fig_sun = px.sunburst(df, path=['Location', 'Công ty', 'BU'], color='Location',
                                  color_discrete_sequence=px.colors.qualitative.Prism)
            fig_sun.update_layout(height=600)
            st.plotly_chart(fig_sun, use_container_width=True)

    with tab_rank:
        c1, c2 = st.columns(2)
        with c1:
            st.markdown("###### 🥇 Top 10 Tài xế")
            if 'Tên tài xế' in df.columns:
                top_driver = df['Tên tài xế'].value_counts().head(10).reset_index()
                top_driver.columns = ['Tài xế', 'Số chuyến']
                st.dataframe(top_driver, use_container_width=True, hide_index=True)
        with c2:
            st.markdown("###### 🥇 Top 10 Người dùng")
            if 'Người sử dụng xe' in df.columns:
                top_user = df['Người sử dụng xe'].value_counts().head(10).reset_index()
                top_user.columns = ['Nhân viên', 'Số chuyến']
                st.dataframe(top_user, use_container_width=True, hide_index=True)

    # --- EXPORT SECTION ---
    st.markdown("---")
    st.markdown("### 📥 TẢI BÁO CÁO")
    
    # Chuẩn bị dữ liệu export
    kpi_export = {
        'total_trips': len(df), 'total_hours': used_hours, 'occupancy': occupancy,
        'cancel_rate': cancel_rate, 'reject_rate': reject_rate,
        'last_month': df['Tháng'].max() if 'Tháng' in df.columns else 'N/A'
    }
    df_comp_exp = df['Công ty'].value_counts().reset_index()
    df_comp_exp.columns = ['Công ty', 'Số chuyến']
    
    df_status_exp = df['Tình trạng đơn yêu cầu'].fillna('Unknown').value_counts().reset_index()
    df_status_exp.columns = ['Trạng thái', 'Số lượng']
    
    df_loc_exp = df['Phạm Vi'].value_counts().reset_index()
    df_loc_exp.columns = ['Phạm Vi', 'Số chuyến']

    pptx_buffer = create_pptx_pro(kpi_export, df_comp_exp, df_status_exp, df_loc_exp)
    
    c_dl1, c_dl2 = st.columns([1, 4])
    with c_dl1:
        st.download_button(
            label="📄 Xuất PPTX (Có Biểu Đồ)",
            data=pptx_buffer,
            file_name="Bao_Cao_Doi_Xe_Pro.pptx",
            mime="application/vnd.openxmlformats-officedocument.presentationml.presentation",
            type="primary"
        )
    with c_dl2:
        st.caption("File PPTX sẽ bao gồm các Slide KPI, Slide Biểu đồ cột (Top Công ty) và Slide Biểu đồ tròn (Chất lượng/Phạm vi).")

else:
    st.info("👋 Hãy upload file Excel để bắt đầu.")