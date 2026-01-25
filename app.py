import streamlit as st
import pandas as pd
import plotly.express as px
import matplotlib.pyplot as plt
from io import BytesIO
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor

# --- 1. CẤU HÌNH TRANG & CSS ---
st.set_page_config(page_title="Fleet Management Pro", page_icon="🚘", layout="wide")

st.markdown("""
<style>
    .block-container {padding-top: 1rem; padding-bottom: 3rem;}
    
    /* KPI Card Style */
    .kpi-card {
        background-color: white; border-radius: 8px; padding: 15px;
        box-shadow: 0 1px 3px rgba(0,0,0,0.1); border-left: 5px solid #0078d4;
        margin-bottom: 10px; height: 100%;
    }
    .kpi-title {
        font-size: 14px; color: #555; font-weight: 700; 
        text-transform: uppercase; margin-bottom: 5px;
    }
    .kpi-value {
        font-size: 26px; font-weight: 800; color: #222; margin: 0;
    }
    .kpi-formula {
        font-size: 11px; color: #888; font-style: italic; margin-top: 8px;
        border-top: 1px solid #eee; padding-top: 5px;
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

        # Load Data
        df_bk = smart_read(xl, sheet_booking, ['ngày khởi hành'])
        df_driver = smart_read(xl, sheet_driver, ['biển số xe']) if sheet_driver else pd.DataFrame()
        df_cbnv = smart_read(xl, sheet_cbnv, ['full name']) if sheet_cbnv else pd.DataFrame()

        df_bk.columns = df_bk.columns.str.strip()
        
        # Merge Driver
        df_final = df_bk
        if not df_driver.empty:
            df_driver.columns = df_driver.columns.str.strip()
            if 'Biển số xe' in df_driver.columns:
                df_driver = df_driver.drop_duplicates(subset=['Biển số xe'], keep='last')
                df_final = df_final.merge(df_driver[['Biển số xe', 'Tên tài xế']], on='Biển số xe', how='left', suffixes=('', '_D'))
                if 'Tên tài xế_D' in df_final.columns:
                    if 'Tên tài xế' not in df_final.columns:
                        df_final['Tên tài xế'] = df_final['Tên tài xế_D']
                    else:
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
        
        if 'Tên tài xế' not in df_final.columns: df_final['Tên tài xế'] = 'Chưa cập nhật'
        df_final['Tên tài xế'] = df_final['Tên tài xế'].fillna('Chưa cập nhật')

        df_final['Start'] = pd.to_datetime(df_final['Ngày khởi hành'].astype(str) + ' ' + df_final['Giờ khởi hành'].astype(str), errors='coerce')
        df_final['End'] = pd.to_datetime(df_final['Ngày khởi hành'].astype(str) + ' ' + df_final['Giờ kết thúc'].astype(str), errors='coerce')
        df_final.loc[df_final['End'] < df_final['Start'], 'End'] += pd.Timedelta(days=1)
        
        df_final['Duration'] = (df_final['End'] - df_final['Start']).dt.total_seconds() / 3600
        df_final['Tháng'] = df_final['Start'].dt.strftime('%Y-%m')
        
        # Scope
        def check_scope(r):
            s = str(r).lower()
            return "Đi Tỉnh" if any(x in s for x in ['tỉnh', 'tp.', 'bình dương', 'đồng nai', 'vũng tàu', 'hà nội']) else "Nội thành"
        df_final['Phạm Vi'] = df_final['Lộ trình'].apply(check_scope) if 'Lộ trình' in df_final.columns else 'Unknown'

        return df_final
    except Exception as e: return f"Lỗi: {str(e)}"

# --- 3. HÀM TẠO ẢNH CHO PPTX (SAFE MODE) ---
def get_chart_img(data, x, y, kind='bar', title='', color='#0078d4'):
    plt.figure(figsize=(6, 4))
    
    if x not in data.columns or y not in data.columns:
        plt.text(0.5, 0.5, 'Data Error', ha='center')
        img = BytesIO(); plt.savefig(img, format='png'); plt.close(); img.seek(0)
        return img

    if kind == 'bar':
        data = data.sort_values(by=x, ascending=True)
        plt.barh(data[y], data[x], color=color)
        plt.xlabel(x)
    elif kind == 'pie':
        plt.pie(data[x], labels=data[y], autopct='%1.1f%%', startangle=90, colors=['#107c10', '#d13438', '#0078d4'])
    
    plt.title(title, fontsize=12, fontweight='bold')
    plt.tight_layout()
    img = BytesIO(); plt.savefig(img, format='png', dpi=100); plt.close(); img.seek(0)
    return img

# --- 4. HÀM XUẤT PPTX ---
def export_pptx(kpi, df_comp, df_status, top_users, top_drivers, df_bad_trips, selected_options):
    prs = Presentation()
    
    def add_title(title, sub):
        slide = prs.slides.add_slide(prs.slide_layouts[0])
        slide.shapes.title.text = title
        slide.placeholders[1].text = sub
    
    add_title("BÁO CÁO VẬN HÀNH ĐỘI XE", f"Dữ liệu đến tháng: {kpi['last_month']}")
    
    slide = prs.slides.add_slide(prs.slide_layouts[1])
    slide.shapes.title.text = "TỔNG QUAN HIỆU SUẤT"
    tf = slide.shapes.placeholders[1].text_frame
    tf.text = f"• Tổng số chuyến: {kpi['trips']}"
    tf.add_paragraph().text = f"• Tổng giờ vận hành: {kpi['hours']:,.0f}h"
    tf.add_paragraph().text = f"• Công suất sử dụng (Occupancy): {kpi['occupancy']:.1f}%"
    tf.add_paragraph().text = f"• Tỷ lệ Hoàn thành: {kpi['success_rate']:.1f}%"
    tf.add_paragraph().text = f"• Tỷ lệ Hủy/Từ chối: {kpi['cancel_rate'] + kpi['reject_rate']:.1f}%"

    if "Biểu đồ Tổng quan" in selected_options:
        slide = prs.slides.add_slide(prs.slide_layouts[5])
        slide.shapes.title.text = "PHÂN BỔ THEO CÔNG TY & TRẠNG THÁI"
        
        img1 = get_chart_img(df_comp.head(8), 'Số chuyến', 'Công ty', 'bar', 'Top Công Ty')
        slide.shapes.add_picture(img1, Inches(0.5), Inches(2), Inches(4.5), Inches(3.5))
        
        img2 = get_chart_img(df_status, 'Số lượng', 'Trạng thái', 'pie', 'Trạng Thái Đơn')
        slide.shapes.add_picture(img2, Inches(5.2), Inches(2), Inches(4.5), Inches(3.5))

    if "Bảng Xếp Hạng (Top User/Driver)" in selected_options:
        slide_u = prs.slides.add_slide(prs.slide_layouts[5])
        slide_u.shapes.title.text = "TOP 10 NGƯỜI SỬ DỤNG NHIỀU NHẤT"
        img_u = get_chart_img(top_users.sort_values('Chuyến', ascending=False).head(10), 'Chuyến', 'Tên', 'bar', '', '#8764b8')
        slide_u.shapes.add_picture(img_u, Inches(1.5), Inches(2), Inches(7), Inches(4.5))
        
        slide_d = prs.slides.add_slide(prs.slide_layouts[5])
        slide_d.shapes.title.text = "TOP 10 TÀI XẾ HOẠT ĐỘNG NHIỀU NHẤT"
        img_d = get_chart_img(top_drivers.sort_values('Chuyến', ascending=False).head(10), 'Chuyến', 'Tên', 'bar', '', '#00cc6a')
        slide_d.shapes.add_picture(img_d, Inches(1.5), Inches(2), Inches(7), Inches(4.5))

    if "Danh sách Hủy/Từ chối" in selected_options:
        slide = prs.slides.add_slide(prs.slide_layouts[5])
        slide.shapes.title.text = "CHI TIẾT ĐƠN HỦY / TỪ CHỐI"
        if not df_bad_trips.empty:
            wanted_cols = ['Start_Str', 'User', 'Status', 'Note']
            avail_cols = [c for c in wanted_cols if c in df_bad_trips.columns]
            
            rows, cols = min(len(df_bad_trips)+1, 10), len(avail_cols)
            if cols > 0:
                table = slide.shapes.add_table(rows, cols, Inches(0.5), Inches(1.5), Inches(9), Inches(0.8)).table
                for i, h in enumerate(avail_cols):
                    cell = table.cell(0, i); cell.text = h
                    cell.fill.solid(); cell.fill.fore_color.rgb = RGBColor(0, 120, 212)
                    cell.text_frame.paragraphs[0].font.color.rgb = RGBColor(255, 255, 255)
                for i, row in enumerate(df_bad_trips.head(9).itertuples(), start=1):
                    for j, col_name in enumerate(avail_cols):
                        val = getattr(row, col_name, "")
                        table.cell(i, j).text = str(val) if str(val) != 'nan' else ""
        else:
            slide.shapes.add_textbox(Inches(1), Inches(2), Inches(5), Inches(1)).text_frame.text = "Không có dữ liệu."

    out = BytesIO(); prs.save(out); out.seek(0)
    return out

# --- 5. GIAO DIỆN CHÍNH ---
st.title("📊 Fleet Management Pro")
uploaded_file = st.file_uploader("Upload Excel", type=['xlsx'], label_visibility="collapsed")

if uploaded_file:
    df = load_data_final(uploaded_file)
    if isinstance(df, str): st.error(df); st.stop()
    
    # --- SIDEBAR FILTERS (FULL VERSION) ---
    with st.sidebar:
        st.header("🗂️ Bộ Lọc Dữ Liệu")
        
        # 1. Date Filter
        min_date, max_date = df['Start'].min().date(), df['Start'].max().date()
        date_range = st.date_input("Thời gian:", (min_date, max_date), min_value=min_date, max_value=max_date)
        if len(date_range) == 2:
            df_date_filtered = df[(df['Start'].dt.date >= date_range[0]) & (df['Start'].dt.date <= date_range[1])]
        else:
            df_date_filtered = df
            
        st.markdown("---")
        
        # 2. Hierarchy Filter
        st.caption("Lọc theo tổ chức (Drill-down):")
        
        # Level 1
        locs = ["Tất cả"] + sorted(df_date_filtered['Location'].unique().tolist())
        sel_loc = st.selectbox("1. Khu vực (Region):", locs)
        df_l1 = df_date_filtered if sel_loc == "Tất cả" else df_date_filtered[df_date_filtered['Location'] == sel_loc]
        
        # Level 2
        comps = ["Tất cả"] + sorted(df_l1['Công ty'].unique().tolist())
        sel_comp = st.selectbox("2. Công ty (Entity):", comps)
        df_l2 = df_l1 if sel_comp == "Tất cả" else df_l1[df_l1['Công ty'] == sel_comp]
        
        # Level 3
        bus = ["Tất cả"] + sorted(df_l2['BU'].unique().tolist())
        sel_bu = st.selectbox("3. Phòng ban (BU):", bus)
        df_filtered = df_l2 if sel_bu == "Tất cả" else df_l2[df_l2['BU'] == sel_bu]
        
        st.markdown("---")
        st.write(f"🔍 Đang xem: **{len(df_filtered)}** chuyến")

    if df_filtered.empty: st.warning("Không có dữ liệu."); st.stop()

    # --- KPI CALCULATION ---
    total_cars = 21
    if 'HCM' in sel_loc or 'NAM' in sel_loc.upper(): total_cars = 16
    elif 'HN' in sel_loc or 'BAC' in sel_loc.upper(): total_cars = 5
    
    days = max((df_filtered['Start'].max() - df_filtered['Start'].min()).days + 1, 1)
    
    total_trips = len(df_filtered)
    total_hours = df_filtered['Duration'].sum()
    occupancy = (total_hours / (total_cars * days * 9) * 100)
    
    counts = df_filtered['Tình trạng đơn yêu cầu'].fillna('Unknown').value_counts()
    completed = counts.get('CLOSED', 0) + counts.get('APPROVED', 0)
    canceled = counts.get('CANCELED', 0) + counts.get('CANCELLED', 0) + counts.get('REJECTED_BY_ADMIN', 0)
    
    suc_rate = (completed / total_trips * 100) if total_trips > 0 else 0
    fail_rate = (canceled / total_trips * 100) if total_trips > 0 else 0

    cols = st.columns(5)
    cards = [
        {"title": "Tổng Chuyến", "val": f"{total_trips}", "sub": "∑ Đếm số dòng", "color": "#0078d4"},
        {"title": "Giờ Vận Hành", "val": f"{total_hours:,.0f}", "sub": "∑ (Giờ về - Giờ đi)", "color": "#0078d4"},
        {"title": "Công Suất (Occupancy)", "val": f"{occupancy:.1f}%", "sub": f"Tổng Giờ / ({total_cars}xe * {days}ngày * 9h)", "color": "#0078d4"},
        {"title": "Hoàn Thành", "val": f"{suc_rate:.1f}%", "sub": "Số đơn xong / Tổng đơn", "color": "#107c10"},
        {"title": "Hủy / Từ Chối", "val": f"{fail_rate:.1f}%", "sub": "Số đơn hủy / Tổng đơn", "color": "#d13438"},
    ]

    for col, card in zip(cols, cards):
        col.markdown(f"""
        <div class="kpi-card" style="border-left: 5px solid {card['color']}">
            <div class="kpi-title">{card['title']}</div>
            <div class="kpi-value" style="color: {card['color']}">{card['val']}</div>
            <div class="kpi-formula">{card['sub']}</div>
        </div>
        """, unsafe_allow_html=True)

    t1, t2, t3 = st.tabs(["📊 Phân Tích Đơn Vị", "🏆 Bảng Xếp Hạng (Top)", "📉 Chi Tiết Chất Lượng"])
    
    with t1:
        c1, c2 = st.columns([2, 1])
        with c1:
            st.write("#### Phân tích theo Cấu trúc")
            if sel_comp == "Tất cả":
                df_g = df_filtered['Công ty'].value_counts().reset_index()
                df_g.columns = ['Công ty', 'Số chuyến']
                st.plotly_chart(px.bar(df_g, x='Số chuyến', y='Công ty', orientation='h', title="Theo Công Ty"), use_container_width=True)
            elif sel_bu == "Tất cả":
                df_g = df_filtered['BU'].value_counts().reset_index()
                df_g.columns = ['Phòng ban', 'Số chuyến']
                st.plotly_chart(px.bar(df_g, x='Số chuyến', y='Phòng ban', orientation='h', title=f"Theo Phòng Ban ({sel_comp})"), use_container_width=True)
            else:
                df_g = df_filtered['Người sử dụng xe'].value_counts().head(10).reset_index()
                df_g.columns = ['Nhân viên', 'Số chuyến']
                st.plotly_chart(px.bar(df_g, x='Số chuyến', y='Nhân viên', orientation='h', title=f"Top NV ({sel_bu})"), use_container_width=True)
        
        with c2:
            st.write("#### Phạm vi di chuyển")
            if 'Phạm Vi' in df_filtered.columns:
                df_sc = df_filtered['Phạm Vi'].value_counts().reset_index()
                df_sc.columns = ['Phạm vi', 'Số lượng']
                st.plotly_chart(px.pie(df_sc, values='Số lượng', names='Phạm vi', hole=0.5), use_container_width=True)

    with t2:
        top_user = df_filtered['Người sử dụng xe'].value_counts().reset_index()
        top_user.columns = ['Tên', 'Chuyến']
        
        top_driver = df_filtered['Tên tài xế'].value_counts().reset_index()
        top_driver.columns = ['Tên', 'Chuyến']

        c_u, c_d = st.columns(2)
        with c_u:
            st.write("##### 🥇 Top 10 Người dùng")
            st.dataframe(top_user.head(10), use_container_width=True, hide_index=True)
        with c_d:
            st.write("##### 🚘 Top 10 Tài xế")
            st.dataframe(top_driver.head(10), use_container_width=True, hide_index=True)

    with t3:
        bad_trips = df_filtered[df_filtered['Tình trạng đơn yêu cầu'].isin(['CANCELED', 'CANCELLED', 'REJECTED_BY_ADMIN'])].copy()
        if not bad_trips.empty:
            st.write(f"##### Danh sách {len(bad_trips)} chuyến bị Hủy/Từ chối")
            desired_cols = ['Ngày khởi hành', 'Người sử dụng xe', 'Tên tài xế', 'Lý do', 'Note', 'Tình trạng đơn yêu cầu']
            valid_cols = [c for c in desired_cols if c in bad_trips.columns]
            st.dataframe(bad_trips[valid_cols], use_container_width=True)
        else:
            st.success("Không có chuyến nào bị hủy trong giai đoạn này.")

    st.divider()
    st.subheader("📥 Xuất Báo Cáo PowerPoint")
    
    c_opt, c_btn = st.columns([2, 1])
    with c_opt:
        pptx_options = st.multiselect(
            "Chọn nội dung muốn đưa vào Slide:",
            ["Biểu đồ Tổng quan", "Bảng Xếp Hạng (Top User/Driver)", "Danh sách Hủy/Từ chối"],
            default=["Biểu đồ Tổng quan", "Bảng Xếp Hạng (Top User/Driver)"]
        )
    
    with c_btn:
        st.write("") 
        st.write("") 
        
        # --- FIX LỖI TYPE ERROR KHI TÍNH THÁNG MAX ---
        last_month_str = "N/A"
        try:
            if not df.empty and 'Tháng' in df.columns:
                valid_months = df['Tháng'].dropna()
                if not valid_months.empty:
                    last_month_str = valid_months.max()
        except: pass

        kpi_data = {
            'trips': total_trips, 'hours': total_hours, 'occupancy': occupancy,
            'success_rate': suc_rate, 'cancel_rate': fail_rate, 'reject_rate': 0,
            'last_month': last_month_str # Sử dụng biến an toàn đã tính ở trên
        }
        
        df_status_exp = counts.reset_index(); df_status_exp.columns = ['Trạng thái', 'Số lượng']
        df_comp_exp = df_filtered['Công ty'].value_counts().reset_index(); df_comp_exp.columns = ['Công ty', 'Số chuyến']

        df_bad_exp = pd.DataFrame()
        if not bad_trips.empty:
            df_bad_exp = bad_trips.copy()
            df_bad_exp['Start_Str'] = df_bad_exp['Start'].dt.strftime('%d/%m')
            df_bad_exp = df_bad_exp.rename(columns={'Người sử dụng xe': 'User', 'Tình trạng đơn yêu cầu': 'Status'})

        pptx_file = export_pptx(kpi_data, df_comp_exp, df_status_exp, top_user, top_driver, df_bad_exp, pptx_options)
        
        st.download_button(
            label="Tải file .PPTX ngay",
            data=pptx_file,
            file_name="Bao_Cao_Van_Hanh_Full.pptx",
            mime="application/vnd.openxmlformats-officedocument.presentationml.presentation",
            type="primary"
        )

else:
    st.info("👋 Vui lòng upload file Excel dữ liệu.")