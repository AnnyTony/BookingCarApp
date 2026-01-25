import streamlit as st
import pandas as pd
import plotly.express as px
import matplotlib.pyplot as plt
from io import BytesIO
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
import re

# --- 1. CẤU HÌNH TRANG & CSS ---
st.set_page_config(page_title="Hệ Thống Quản Trị & Tối Ưu Hóa Đội Xe", page_icon="🚘", layout="wide")

st.markdown("""
<style>
    .block-container {padding-top: 1rem; padding-bottom: 3rem;}
    .kpi-card {
        background-color: white; border-radius: 12px; padding: 20px;
        box-shadow: 0 4px 6px rgba(0,0,0,0.05);
        transition: transform 0.2s, box-shadow 0.2s;
        border: 1px solid #f0f2f6;
        height: 100%; display: flex; flex-direction: column; justify-content: space-between;
        min-height: 160px;
    }
    .kpi-card:hover { transform: translateY(-5px); box-shadow: 0 10px 15px rgba(0,0,0,0.1); }
    .kpi-header { display: flex; justify-content: space-between; align-items: center; margin-bottom: 10px; }
    .kpi-title { font-size: 14px; color: #6c757d; font-weight: 700; text-transform: uppercase; letter-spacing: 0.5px; }
    .kpi-icon { font-size: 20px; background: #f8f9fa; padding: 8px; border-radius: 8px; }
    .kpi-value { font-size: 32px; font-weight: 800; color: #212529; margin: 0; }
    .kpi-formula { font-size: 12px; color: #888; font-style: italic; margin-top: auto; padding-top: 10px; border-top: 1px dashed #eee; }
    .progress-bg { background-color: #e9ecef; border-radius: 4px; height: 6px; width: 100%; margin: 8px 0; overflow: hidden; }
    .progress-fill { height: 100%; border-radius: 4px; transition: width 0.5s ease-in-out; }
</style>
""", unsafe_allow_html=True)

# --- 2. HÀM XỬ LÝ DỮ LIỆU (LOGIC MỚI: UNION & NORMALIZE) ---
@st.cache_data
def load_data_final(file):
    try:
        xl = pd.ExcelFile(file, engine='openpyxl')
        
        # Tìm sheet linh hoạt
        sheet_driver = next((s for s in xl.sheet_names if 'driver' in s.lower()), None)
        sheet_booking = next((s for s in xl.sheet_names if 'booking' in s.lower()), None)
        sheet_cbnv = next((s for s in xl.sheet_names if 'cbnv' in s.lower()), None)
        
        if not sheet_booking: return "❌ Không tìm thấy sheet 'Booking car'.", [], pd.DataFrame()

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
        
        # --- HÀM CHUẨN HÓA & LỌC BIỂN SỐ ---
        def normalize_plate(plate):
            if not isinstance(plate, str): return ""
            # Loại bỏ chấm, gạch ngang, khoảng trắng, chuyển về chữ hoa
            clean = re.sub(r'[^A-Z0-9]', '', plate.upper())
            return clean

        def is_valid_plate(plate):
            s = str(plate).strip().upper()
            if len(s) > 15 or len(s) < 5: return False
            if ":" in s or "202" in s or "BIỂN SỐ" in s: return False # Loại bỏ ngày tháng/tiêu đề
            return any(char.isdigit() for char in s) # Phải có số
        
        # 1. Lấy xe từ Driver Sheet
        driver_cars = set()
        if not df_driver.empty:
            df_driver.columns = df_driver.columns.str.strip()
            if 'Biển số xe' in df_driver.columns:
                raw_driver = df_driver['Biển số xe'].dropna().unique()
                driver_cars = {normalize_plate(p) for p in raw_driver if is_valid_plate(p)}
                
                # Merge thông tin tài xế
                df_driver['Biển_Clean'] = df_driver['Biển số xe'].apply(normalize_plate)
                df_driver = df_driver.drop_duplicates(subset=['Biển_Clean'], keep='last')
                
                # Tạo cột Clean cho bảng chính để merge
                df_final['Biển_Clean'] = df_final['Biển số xe'].apply(normalize_plate)
                df_final = df_final.merge(df_driver[['Biển_Clean', 'Tên tài xế']], on='Biển_Clean', how='left', suffixes=('', '_D'))
                
                if 'Tên tài xế_D' in df_final.columns:
                    df_final['Tên tài xế'] = df_final['Tên tài xế'].fillna(df_final['Tên tài xế_D'])

        # 2. Lấy xe từ Booking History
        booking_cars = set()
        if 'Biển số xe' in df_final.columns:
            raw_booking = df_final['Biển số xe'].dropna().unique()
            booking_cars = {normalize_plate(p) for p in raw_booking if is_valid_plate(p)}

        # 3. Tổng hợp (Union) -> Danh sách xe duy nhất chuẩn hóa
        all_unique_cars = sorted(list(driver_cars.union(booking_cars)))
        
        # Tạo DataFrame chi tiết xe để đối soát (Quan trọng cho Tab 4)
        df_cars_check = pd.DataFrame({'Biển Số Chuẩn': all_unique_cars})
        df_cars_check['Nguồn'] = df_cars_check['Biển Số Chuẩn'].apply(
            lambda x: 'Cả hai' if (x in driver_cars and x in booking_cars) 
            else ('Chỉ có trong Driver' if x in driver_cars else 'Vãng lai (Chỉ có trong Booking)')
        )

        # Xử lý các cột khác
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
        
        df_final['Tên tài xế'] = df_final['Tên tài xế'].fillna('Chưa cập nhật')

        df_final['Start'] = pd.to_datetime(df_final['Ngày khởi hành'].astype(str) + ' ' + df_final['Giờ khởi hành'].astype(str), errors='coerce')
        df_final['End'] = pd.to_datetime(df_final['Ngày khởi hành'].astype(str) + ' ' + df_final['Giờ kết thúc'].astype(str), errors='coerce')
        df_final.loc[df_final['End'] < df_final['Start'], 'End'] += pd.Timedelta(days=1)
        df_final['Duration'] = (df_final['End'] - df_final['Start']).dt.total_seconds() / 3600
        df_final['Tháng'] = df_final['Start'].dt.strftime('%Y-%m')
        
        # Logic Đi Tỉnh
        def check_scope_v2(r):
            s = str(r).lower()
            provinces = ['bình dương', 'đồng nai', 'long an', 'bà rịa', 'vũng tàu', 'tây ninh', 'bình phước', 'tiền giang', 'bến tre', 'cần thơ', 'vĩnh long', 'an giang', 'bắc ninh', 'hưng yên', 'hải dương', 'hải phòng', 'vĩnh phúc', 'hà nam', 'nam định', 'thái bình', 'thái nguyên', 'hòa bình', 'bắc giang', 'phú thọ', 'thanh hóa', 'nghệ an']
            if any(p in s for p in provinces): return "Đi Tỉnh"
            return "Nội thành"

        df_final['Phạm Vi'] = df_final['Lộ trình'].apply(check_scope_v2) if 'Lộ trình' in df_final.columns else 'Unknown'
        
        # Thêm cột Biển_Clean vào df_final để filter sau này
        if 'Biển_Clean' not in df_final.columns:
             df_final['Biển_Clean'] = df_final['Biển số xe'].apply(normalize_plate)

        return df_final, all_unique_cars, df_cars_check
    except Exception as e: return f"Lỗi: {str(e)}", [], pd.DataFrame()

# --- 3. HÀM TẠO ẢNH CHO PPTX ---
def get_chart_img(data, x, y, kind='bar', title='', color='#0078d4'):
    plt.figure(figsize=(7, 4.5))
    if x not in data.columns or y not in data.columns:
        plt.text(0.5, 0.5, 'No Data', ha='center'); img = BytesIO(); plt.savefig(img, format='png'); plt.close(); img.seek(0); return img

    if kind == 'bar': 
        data = data.sort_values(by=x, ascending=True)
        bars = plt.barh(data[y], data[x], color=color); plt.xlabel(x); plt.bar_label(bars, fmt='%g')
    elif kind == 'column': 
        bars = plt.bar(data[y], data[x], color=color); plt.ylabel(x); plt.xticks(rotation=45, ha='right'); plt.bar_label(bars, fmt='%g')
    elif kind == 'pie': 
        plt.pie(data[x], labels=data[y], autopct='%1.1f%%', startangle=90, colors=['#107c10', '#d13438', '#0078d4', '#ffc107', '#8764b8'])
    
    plt.title(title, pad=15, fontweight='bold', fontsize=12, color='#333'); plt.tight_layout()
    img = BytesIO(); plt.savefig(img, format='png', dpi=120); plt.close(); img.seek(0)
    return img

# --- 4. HÀM XUẤT PPTX ---
def export_pptx(kpi, df_comp, df_status, top_users, top_drivers, df_bad_trips, selected_options, chart_prefs, df_scope):
    prs = Presentation()
    
    def add_kpi_shape(slide, left, top, width, height, title, value, sub, color_rgb):
        shape = slide.shapes.add_shape(1, left, top, width, height)
        shape.fill.solid(); shape.fill.fore_color.rgb = RGBColor(255, 255, 255)
        shape.line.color.rgb = color_rgb; shape.line.width = Pt(2.5)
        
        tb = slide.shapes.add_textbox(left + Inches(0.1), top + Inches(0.1), width - Inches(0.2), Inches(0.3))
        tb.text_frame.text = title; tb.text_frame.paragraphs[0].font.bold = True; tb.text_frame.paragraphs[0].font.color.rgb = RGBColor(100, 100, 100)
        
        tb_v = slide.shapes.add_textbox(left + Inches(0.1), top + Inches(0.4), width - Inches(0.2), Inches(0.5))
        p_v = tb_v.text_frame.paragraphs[0]; p_v.text = str(value); p_v.font.size = Pt(24); p_v.font.bold = True
        
        tb_s = slide.shapes.add_textbox(left + Inches(0.1), top + height - Inches(0.4), width - Inches(0.2), Inches(0.3))
        p_s = tb_s.text_frame.paragraphs[0]; p_s.text = sub; p_s.font.size = Pt(9); p_s.font.italic = True; p_s.font.color.rgb = RGBColor(150, 150, 150)

    slide = prs.slides.add_slide(prs.slide_layouts[0])
    slide.shapes.title.text = "BÁO CÁO VẬN HÀNH ĐỘI XE"; slide.placeholders[1].text = f"Cập nhật đến tháng: {kpi['last_month']}"

    slide = prs.slides.add_slide(prs.slide_layouts[5]); slide.shapes.title.text = "TỔNG QUAN HIỆU SUẤT"
    add_kpi_shape(slide, Inches(0.5), Inches(2.5), Inches(1.8), Inches(1.5), "TỔNG CHUYẾN", f"{kpi['trips']}", "Số chuyến", RGBColor(0, 120, 212))
    add_kpi_shape(slide, Inches(2.4), Inches(2.5), Inches(1.8), Inches(1.5), "GIỜ VẬN HÀNH", f"{kpi['hours']:,.0f}", "Tổng giờ", RGBColor(0, 120, 212))
    add_kpi_shape(slide, Inches(4.3), Inches(2.5), Inches(1.8), Inches(1.5), "CÔNG SUẤT", f"{kpi['occupancy']:.1f}%", "Mục tiêu >50%", RGBColor(0, 120, 212))
    add_kpi_shape(slide, Inches(6.2), Inches(2.5), Inches(1.8), Inches(1.5), "HOÀN THÀNH", f"{kpi['success_rate']:.1f}%", "Tỷ lệ OK", RGBColor(16, 124, 16))
    add_kpi_shape(slide, Inches(8.1), Inches(2.5), Inches(1.8), Inches(1.5), "HỦY/TỪ CHỐI", f"{kpi['cancel_rate'] + kpi['reject_rate']:.1f}%", "Tỷ lệ Fail", RGBColor(209, 52, 56))

    if "Biểu đồ Tổng quan" in selected_options:
        slide = prs.slides.add_slide(prs.slide_layouts[5]); slide.shapes.title.text = "PHÂN TÍCH CẤU TRÚC SỬ DỤNG"
        img1 = get_chart_img(df_comp.head(8), 'Value', 'Category', kind=chart_prefs.get('structure', 'bar'), title='Top Đơn Vị')
        slide.shapes.add_picture(img1, Inches(0.5), Inches(1.8), Inches(4.5), Inches(3.5))
        img2 = get_chart_img(df_scope, 'Số lượng', 'Phạm vi', kind=chart_prefs.get('scope', 'pie'), title='Phạm Vi')
        slide.shapes.add_picture(img2, Inches(5.2), Inches(1.8), Inches(4.5), Inches(3.5))

    if "Bảng Xếp Hạng (Top User/Driver)" in selected_options:
        slide = prs.slides.add_slide(prs.slide_layouts[5]); slide.shapes.title.text = "BẢNG XẾP HẠNG HOẠT ĐỘNG"
        img_u = get_chart_img(top_users.head(8), 'Số_chuyến', 'Người sử dụng xe', kind=chart_prefs.get('top_user', 'bar'), title='Top Users', color='#8764b8')
        slide.shapes.add_picture(img_u, Inches(0.5), Inches(1.8), Inches(4.5), Inches(3.5))
        img_d = get_chart_img(top_drivers.head(8), 'Số_chuyến', 'Tên tài xế', kind=chart_prefs.get('top_driver', 'bar'), title='Top Drivers', color='#00cc6a')
        slide.shapes.add_picture(img_d, Inches(5.2), Inches(1.8), Inches(4.5), Inches(3.5))

    if "Danh sách Hủy/Từ chối" in selected_options:
        slide = prs.slides.add_slide(prs.slide_layouts[5]); slide.shapes.title.text = "CHI TIẾT ĐƠN HỦY / TỪ CHỐI"
        if not df_bad_trips.empty:
            wanted_cols = ['Start_Str', 'User', 'Status', 'Note']
            avail_cols = [c for c in wanted_cols if c in df_bad_trips.columns]
            rows, cols = min(len(df_bad_trips)+1, 10), len(avail_cols)
            if cols > 0:
                table = slide.shapes.add_table(rows, cols, Inches(0.5), Inches(1.5), Inches(9), Inches(0.8)).table
                for i, h in enumerate(avail_cols):
                    cell = table.cell(0, i); cell.text = h
                for i, row in enumerate(df_bad_trips.head(9).itertuples(), start=1):
                    for j, col_name in enumerate(avail_cols):
                        val = getattr(row, col_name, ""); table.cell(i, j).text = str(val)[:30]

    out = BytesIO(); prs.save(out); out.seek(0); return out

# --- 5. GIAO DIỆN CHÍNH ---
st.title("📊 Phước Minh - Hệ Thống Quản Trị & Tối Ưu Hóa Đội Xe")
uploaded_file = st.file_uploader("Upload Excel", type=['xlsx'], label_visibility="collapsed")

if uploaded_file:
    # Load data
    df, all_unique_cars, df_cars_check = load_data_final(uploaded_file)
    if isinstance(df, str): st.error(df); st.stop()
    
    with st.sidebar:
        st.header("🗂️ Bộ Lọc Dữ Liệu")
        min_date, max_date = df['Start'].min().date(), df['Start'].max().date()
        date_range = st.date_input("Thời gian:", (min_date, max_date), min_value=min_date, max_value=max_date)
        if len(date_range) == 2:
            df_date_filtered = df[(df['Start'].dt.date >= date_range[0]) & (df['Start'].dt.date <= date_range[1])]
        else:
            df_date_filtered = df
            
        st.markdown("---")
        st.caption("Lọc theo tổ chức (Drill-down):")
        locs = ["Tất cả"] + sorted(df_date_filtered['Location'].unique().tolist())
        sel_loc = st.selectbox("1. Khu vực (Region):", locs)
        df_l1 = df_date_filtered if sel_loc == "Tất cả" else df_date_filtered[df_date_filtered['Location'] == sel_loc]
        comps = ["Tất cả"] + sorted(df_l1['Công ty'].unique().tolist())
        sel_comp = st.selectbox("2. Công ty (Entity):", comps)
        df_l2 = df_l1 if sel_comp == "Tất cả" else df_l1[df_l1['Công ty'] == sel_comp]
        bus = ["Tất cả"] + sorted(df_l2['BU'].unique().tolist())
        sel_bu = st.selectbox("3. Phòng ban (BU):", bus)
        df_filtered = df_l2 if sel_bu == "Tất cả" else df_l2[df_l2['BU'] == sel_bu]
        st.markdown("---"); st.write(f"🔍 Đang xem: **{len(df_filtered)}** chuyến")

    if df_filtered.empty: st.warning("Không có dữ liệu."); st.stop()

    # --- KPI CALCULATION ---
    # Tự động tính số xe dựa trên bộ lọc
    if sel_loc == "Tất cả" and sel_comp == "Tất cả" and sel_bu == "Tất cả":
        total_cars_kpi = len(all_unique_cars) # Lấy tổng xe đã chuẩn hóa
        cars_display = all_unique_cars
    else:
        # Lấy danh sách xe trong vùng filter hiện tại
        active_raw = df_filtered['Biển_Clean'].dropna().unique().tolist()
        cars_display = sorted(active_raw)
        total_cars_kpi = len(cars_display)
        if total_cars_kpi == 0: total_cars_kpi = 1

    days = max((df_filtered['Start'].max() - df_filtered['Start'].min()).days + 1, 1)
    total_trips = len(df_filtered)
    total_hours = df_filtered['Duration'].sum()
    
    occupancy_cap = total_cars_kpi * days * 8
    occupancy = (total_hours / occupancy_cap * 100) if occupancy_cap > 0 else 0
    
    counts = df_filtered['Tình trạng đơn yêu cầu'].fillna('Unknown').value_counts()
    suc_rate = ((counts.get('CLOSED', 0) + counts.get('APPROVED', 0)) / total_trips * 100) if total_trips > 0 else 0
    fail_rate = ((counts.get('CANCELED', 0) + counts.get('CANCELLED', 0) + counts.get('REJECTED_BY_ADMIN', 0)) / total_trips * 100) if total_trips > 0 else 0

    # --- KPI UI ---
    cols = st.columns(5)
    cards = [
        {"title": "Tổng Chuyến", "val": f"{total_trips}", "sub": "∑ Đếm số dòng", "color": "#0078d4", "icon": "🚘", "is_percent": False},
        {"title": "Giờ Vận Hành", "val": f"{total_hours:,.0f}", "sub": "∑ (Giờ về - Giờ đi)", "color": "#0078d4", "icon": "⏱️", "is_percent": False},
        {"title": "Công Suất", "val": f"{occupancy:.1f}%", "sub": f"Giờ / ({total_cars_kpi}xe * {days}ngày * 8h)", "color": "#0078d4", "icon": "📉", "is_percent": True, "pct_val": min(occupancy, 100)},
        {"title": "Hoàn Thành", "val": f"{suc_rate:.1f}%", "sub": "Tỷ lệ thành công", "color": "#107c10", "icon": "✅", "is_percent": True, "pct_val": suc_rate},
        {"title": "Hủy / Từ Chối", "val": f"{fail_rate:.1f}%", "sub": "Tỷ lệ thất bại", "color": "#d13438", "icon": "🚫", "is_percent": True, "pct_val": fail_rate},
    ]

    for col, card in zip(cols, cards):
        progress_html = f'<div class="progress-bg"><div class="progress-fill" style="width: {card["pct_val"]}%; background-color: {card["color"]}"></div></div>' if card["is_percent"] else '<div style="height: 24px;"></div>'
        html_code = f"""<div class="kpi-card" style="border-top: 4px solid {card['color']}">
<div class="kpi-header"><span class="kpi-title" style="color: {card['color']}">{card['title']}</span><span class="kpi-icon">{card['icon']}</span></div>
<div class="kpi-value">{card['val']}</div>{progress_html}<div class="kpi-formula">{card['sub']}</div></div>"""
        col.markdown(html_code, unsafe_allow_html=True)

    # --- TABS ---
    t1, t2, t3, t4 = st.tabs(["📊 Phân Tích Đơn Vị", "🏆 Bảng Xếp Hạng", "📉 Chất Lượng", "⚙️ Chi Tiết & Đối Soát"])
    
    chart_prefs = {}
    kind_map = {"Thanh ngang (Bar)": "bar", "Thanh dọc (Column)": "column", "Tròn (Pie)": "pie"}

    with t1:
        c1, c2 = st.columns([2, 1])
        with c1:
            chart_type_struct = st.selectbox("Kiểu biểu đồ Cấu trúc:", list(kind_map.keys()), index=0, key="c_struct")
            chart_prefs['structure'] = kind_map[chart_type_struct]
            if sel_comp == "Tất cả": df_g = df_filtered['Công ty'].value_counts().reset_index(); df_g.columns = ['Category', 'Value']; title_c = "Theo Công Ty"
            elif sel_bu == "Tất cả": df_g = df_filtered['BU'].value_counts().reset_index(); df_g.columns = ['Category', 'Value']; title_c = f"Theo Phòng Ban ({sel_comp})"
            else: df_g = df_filtered['Người sử dụng xe'].value_counts().head(10).reset_index(); df_g.columns = ['Category', 'Value']; title_c = f"Top NV ({sel_bu})"
            
            if chart_prefs['structure'] == "bar": fig = px.bar(df_g, x='Value', y='Category', orientation='h', text='Value', title=title_c)
            elif chart_prefs['structure'] == "column": fig = px.bar(df_g, x='Category', y='Value', text='Value', title=title_c)
            else: fig = px.pie(df_g, values='Value', names='Category', title=title_c)
            st.plotly_chart(fig, use_container_width=True)
        
        with c2:
            chart_type_scope = st.selectbox("Kiểu biểu đồ Phạm vi:", list(kind_map.keys()), index=2, key="c_scope")
            chart_prefs['scope'] = kind_map[chart_type_scope]
            if 'Phạm Vi' in df_filtered.columns:
                df_sc = df_filtered['Phạm Vi'].value_counts().reset_index(); df_sc.columns = ['Phạm vi', 'Số lượng']
                if chart_prefs['scope'] == "bar": fig_s = px.bar(df_sc, x='Số lượng', y='Phạm vi', orientation='h', text='Số lượng', title="Phạm Vi Di Chuyển")
                elif chart_prefs['scope'] == "column": fig_s = px.bar(df_sc, x='Phạm vi', y='Số lượng', text='Số lượng', title="Phạm Vi Di Chuyển")
                else: fig_s = px.pie(df_sc, values='Số lượng', names='Phạm vi', hole=0.5, title="Phạm Vi Di Chuyển")
                st.plotly_chart(fig_s, use_container_width=True)
                
                with st.expander("🔍 Kiểm tra chi tiết Phạm Vi (Xem tại đây)"):
                    st.write("Dữ liệu Lộ trình & Phân loại:")
                    st.dataframe(df_filtered[['Ngày khởi hành', 'Lộ trình', 'Phạm Vi']], use_container_width=True)

    with t2:
        df_user_stats = df_filtered.groupby('Người sử dụng xe').agg(Số_chuyến=('Start', 'count'), Công_ty=('Công ty', lambda x: x.mode()[0] if not x.mode().empty else 'Unknown')).reset_index().sort_values('Số_chuyến', ascending=False)
        df_driver_stats = df_filtered.groupby('Tên tài xế').agg(Số_chuyến=('Start', 'count'), Tuyến_hay_chạy=('Lộ trình', lambda x: x.mode()[0] if not x.mode().empty else 'N/A')).reset_index().sort_values('Số_chuyến', ascending=False)

        c_u, c_d = st.columns(2)
        with c_u:
            type_u = st.selectbox("Biểu đồ Top User:", list(kind_map.keys()), index=0, key="c_user")
            chart_prefs['top_user'] = kind_map[type_u]
            st.write("##### 🥇 Top User (Kèm Công ty)")
            if chart_prefs['top_user'] == "bar": fig_u = px.bar(df_user_stats.head(10), x='Số_chuyến', y='Người sử dụng xe', orientation='h', text='Số_chuyến', hover_data=['Công_ty'], color_discrete_sequence=['#8764b8'])
            elif chart_prefs['top_user'] == "column": fig_u = px.bar(df_user_stats.head(10), x='Người sử dụng xe', y='Số_chuyến', text='Số_chuyến', hover_data=['Công_ty'], color_discrete_sequence=['#8764b8'])
            else: fig_u = px.pie(df_user_stats.head(10), values='Số_chuyến', names='Người sử dụng xe', hover_data=['Công_ty'])
            st.plotly_chart(fig_u, use_container_width=True)
            st.dataframe(df_user_stats.head(10), use_container_width=True, hide_index=True)

        with c_d:
            type_d = st.selectbox("Biểu đồ Top Driver:", list(kind_map.keys()), index=0, key="c_driver")
            chart_prefs['top_driver'] = kind_map[type_d]
            st.write("##### 🚘 Top Driver (Kèm Tuyến phổ biến)")
            if chart_prefs['top_driver'] == "bar": fig_d = px.bar(df_driver_stats.head(10), x='Số_chuyến', y='Tên tài xế', orientation='h', text='Số_chuyến', hover_data=['Tuyến_hay_chạy'], color_discrete_sequence=['#00cc6a'])
            elif chart_prefs['top_driver'] == "column": fig_d = px.bar(df_driver_stats.head(10), x='Tên tài xế', y='Số_chuyến', text='Số_chuyến', hover_data=['Tuyến_hay_chạy'], color_discrete_sequence=['#00cc6a'])
            else: fig_d = px.pie(df_driver_stats.head(10), values='Số_chuyến', names='Tên tài xế', hover_data=['Tuyến_hay_chạy'])
            st.plotly_chart(fig_d, use_container_width=True)
            st.dataframe(df_driver_stats.head(10), use_container_width=True, hide_index=True)

    with t3:
        c_status_left, c_status_right = st.columns(2)
        with c_status_left:
            chart_type_status = st.selectbox("Kiểu biểu đồ Trạng thái:", list(kind_map.keys()), index=2, key="c_status")
            chart_prefs['status'] = kind_map[chart_type_status]
            st.write("#### Tỷ lệ Trạng thái")
            df_st = counts.reset_index(); df_st.columns = ['Status', 'Count']
            if chart_prefs['status'] == "pie": fig_st = px.pie(df_st, values='Count', names='Status', hole=0.4, color='Status', color_discrete_map={'CLOSED':'#107c10', 'CANCELED':'#d13438', 'REJECTED_BY_ADMIN':'#a80000'})
            elif chart_prefs['status'] == "bar": fig_st = px.bar(df_st, x='Count', y='Status', orientation='h', text='Count', color='Status')
            else: fig_st = px.bar(df_st, x='Status', y='Count', text='Count', color='Status')
            st.plotly_chart(fig_st, use_container_width=True)

        with c_status_right:
            bad_trips = df_filtered[df_filtered['Tình trạng đơn yêu cầu'].isin(['CANCELED', 'CANCELLED', 'REJECTED_BY_ADMIN'])].copy()
            if not bad_trips.empty:
                st.write(f"##### Danh sách {len(bad_trips)} chuyến bị Hủy/Từ chối")
                wanted = ['Ngày khởi hành', 'Người sử dụng xe', 'Tên tài xế', 'Lý do', 'Note', 'Tình trạng đơn yêu cầu']
                actual = [c for c in wanted if c in bad_trips.columns]
                st.dataframe(bad_trips[actual], use_container_width=True)
            else: st.success("Không có chuyến nào bị hủy.")

    with t4:
        st.subheader("⚙️ Đối Soát Công Thức & Dữ Liệu")
        st.info("Tab này dùng để kiểm tra tính chính xác của các chỉ số KPI.")
        c_kpi_check, c_chart_check = st.columns(2)
        
        with c_kpi_check:
            st.write("#### 1. Các tham số tính Công Suất")
            st.write(f"- **Tổng số xe ($N$):** {total_cars_kpi} xe")
            
            # --- SHOW LIST XE ---
            with st.expander(f"🚗 Xem danh sách {len(cars_display)} xe đã chuẩn hóa (Click để mở)"):
                st.write("Danh sách biển số xe sau khi loại bỏ trùng lặp và làm sạch:")
                df_disp = pd.DataFrame(cars_display, columns=["Biển Số"])
                # Nếu đang ở chế độ 'Tất cả' (không filter), hiển thị thêm cột Nguồn gốc
                if not df_cars_check.empty and len(cars_display) == len(all_unique_cars):
                     st.dataframe(df_cars_check, use_container_width=True)
                else:
                     st.dataframe(df_disp, use_container_width=True)
            # --------------------

            st.write(f"- **Số ngày trong kỳ lọc ($D$):** {days} ngày (từ {df_filtered['Start'].min().date()} đến {df_filtered['Start'].max().date()})")
            st.write(f"- **Giờ tiêu chuẩn/ngày:** 8 giờ")
            st.markdown("---")
            st.write(f"👉 **Năng lực tối đa (Capacity):** {total_cars_kpi} * {days} * 8 = **{occupancy_cap:,.0f} giờ**")
            st.write(f"👉 **Thực tế sử dụng (Actual):** **{total_hours:,.0f} giờ**")
            st.metric("Kết quả Occupancy", f"{occupancy:.2f}%")
            
        with c_chart_check:
            st.write("#### 2. Biểu đồ So Sánh Năng Lực")
            df_check = pd.DataFrame({'Loại': ['Năng Lực Tối Đa', 'Thực Tế Sử Dụng'], 'Giờ': [occupancy_cap, total_hours]})
            fig_check = px.bar(df_check, x='Loại', y='Giờ', text='Giờ', color='Loại', color_discrete_map={'Năng Lực Tối Đa': '#e9ecef', 'Thực Tế Sử Dụng': '#0078d4'})
            st.plotly_chart(fig_check, use_container_width=True)

    st.divider()
    st.subheader("📥 Xuất Báo Cáo PowerPoint")
    c_opt, c_btn = st.columns([2, 1])
    with c_opt:
        pptx_options = st.multiselect("Chọn nội dung Slide:", ["Biểu đồ Tổng quan", "Bảng Xếp Hạng (Top User/Driver)", "Danh sách Hủy/Từ chối"], default=["Biểu đồ Tổng quan", "Bảng Xếp Hạng (Top User/Driver)"])
    with c_btn:
        st.write(""); st.write("")
        last_month_str = "N/A"
        try:
            if not df.empty and 'Tháng' in df.columns:
                valid_months = df['Tháng'].dropna()
                if not valid_months.empty: last_month_str = valid_months.max()
        except: pass

        kpi_data = {'trips': total_trips, 'hours': total_hours, 'occupancy': occupancy, 'success_rate': suc_rate, 'cancel_rate': fail_rate, 'reject_rate': 0, 'last_month': last_month_str}
        df_status_exp = counts.reset_index(); df_status_exp.columns = ['Trạng thái', 'Số lượng']
        if sel_comp == "Tất cả": df_comp_exp = df_filtered['Công ty'].value_counts().reset_index(); df_comp_exp.columns=['Category', 'Value']
        elif sel_bu == "Tất cả": df_comp_exp = df_filtered['BU'].value_counts().reset_index(); df_comp_exp.columns=['Category', 'Value']
        else: df_comp_exp = df_filtered['Người sử dụng xe'].value_counts().head(10).reset_index(); df_comp_exp.columns=['Category', 'Value']
        
        if 'Phạm Vi' in df_filtered.columns: df_scope_exp = df_filtered['Phạm Vi'].value_counts().reset_index(); df_scope_exp.columns = ['Phạm vi', 'Số lượng']
        else: df_scope_exp = pd.DataFrame(columns=['Phạm vi', 'Số lượng'])
        
        df_bad_exp = pd.DataFrame()
        if not bad_trips.empty:
            df_bad_exp = bad_trips.copy()
            df_bad_exp['Start_Str'] = df_bad_exp['Start'].dt.strftime('%d/%m')
            df_bad_exp = df_bad_exp.rename(columns={'Người sử dụng xe': 'User', 'Tình trạng đơn yêu cầu': 'Status'})

        pptx_file = export_pptx(kpi_data, df_comp_exp, df_status_exp, df_user_stats, df_driver_stats, df_bad_exp, pptx_options, chart_prefs, df_scope_exp)
        st.download_button(label="Tải file .PPTX ngay", data=pptx_file, file_name="Bao_Cao_Van_Hanh_Full.pptx", mime="application/vnd.openxmlformats-officedocument.presentationml.presentation", type="primary")

else:
    st.info("👋 Vui lòng upload file Excel dữ liệu.")