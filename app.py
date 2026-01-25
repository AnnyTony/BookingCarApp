import streamlit as st
import pandas as pd
import plotly.express as px
import matplotlib.pyplot as plt
from io import BytesIO
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
import re

# --- 1. CẤU HÌNH TRANG ---
st.set_page_config(page_title="Hệ Thống Quản Trị Đội Xe", page_icon="🚘", layout="wide")

st.markdown("""
<style>
    .block-container {padding-top: 1rem; padding-bottom: 3rem;}
    .kpi-card {
        background-color: white; border-radius: 12px; padding: 20px;
        box-shadow: 0 4px 6px rgba(0,0,0,0.05); border: 1px solid #f0f2f6;
        height: 100%; display: flex; flex-direction: column; justify-content: space-between;
        min-height: 160px;
    }
    .kpi-card:hover { transform: translateY(-5px); box-shadow: 0 10px 15px rgba(0,0,0,0.1); }
    .kpi-header { display: flex; justify-content: space-between; align-items: center; margin-bottom: 10px; }
    .kpi-title { font-size: 14px; color: #6c757d; font-weight: 700; text-transform: uppercase; }
    .kpi-icon { font-size: 20px; background: #f8f9fa; padding: 8px; border-radius: 8px; }
    .kpi-value { font-size: 32px; font-weight: 800; color: #212529; margin: 0; }
    .kpi-formula { font-size: 12px; color: #888; font-style: italic; margin-top: 10px; border-top: 1px dashed #eee; padding-top: 5px; }
    .progress-bg { background-color: #e9ecef; border-radius: 4px; height: 6px; width: 100%; margin: 8px 0; overflow: hidden; }
    .progress-fill { height: 100%; border-radius: 4px; transition: width 0.5s ease-in-out; }
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
        
        if not sheet_booking: return "❌ Không tìm thấy sheet 'Booking car'.", pd.DataFrame(), {}, pd.DataFrame()

        def read_sheet_smart(sheet_name, key_col):
            df_preview = xl.parse(sheet_name, header=None, nrows=15)
            header_idx = -1
            for idx, row in df_preview.iterrows():
                row_str = row.astype(str).str.lower().tolist()
                if any(key_col.lower() in s for s in row_str):
                    header_idx = idx
                    break
            if header_idx == -1: return pd.DataFrame()
            return xl.parse(sheet_name, header=header_idx)

        df_bk = read_sheet_smart(sheet_booking, 'ngày khởi hành')
        df_driver = read_sheet_smart(sheet_driver, 'biển số xe')
        df_cbnv = read_sheet_smart(sheet_cbnv, 'full name')

        df_bk.columns = df_bk.columns.str.strip()
        if not df_driver.empty: df_driver.columns = df_driver.columns.str.strip()
        if not df_cbnv.empty: df_cbnv.columns = df_cbnv.columns.str.strip()

        # --- HÀM CHUẨN HÓA ---
        def normalize_plate(plate):
            if not isinstance(plate, str): return ""
            return re.sub(r'[^A-Z0-9]', '', plate.upper())

        def is_valid_plate(plate):
            s = str(plate).strip().upper()
            if len(s) > 15 or len(s) < 5: return False
            if ":" in s or "202" in s or "BIỂN SỐ" in s: return False
            return any(char.isdigit() for char in s)

        # 1. Xử lý Master Data (Driver)
        driver_cars_map = {}
        duplicates_check = []
        
        if not df_driver.empty and 'Biển số xe' in df_driver.columns:
            cc_col = next((c for c in df_driver.columns if 'cost' in c.lower()), None)
            
            for idx, row in df_driver.iterrows():
                raw_plate = row['Biển số xe']
                if is_valid_plate(raw_plate):
                    clean_plate = normalize_plate(raw_plate)
                    if clean_plate in driver_cars_map:
                        duplicates_check.append(raw_plate)
                    
                    driver_cars_map[clean_plate] = {
                        'Raw': raw_plate,
                        'Driver_Name': row.get('Tên tài xế', 'Unknown'),
                        'Cost_Center': row.get(cc_col, 'Unknown') if cc_col else 'Unknown'
                    }

        # 2. Xử lý Transaction Data (Booking)
        if 'Biển số xe' not in df_bk.columns: return "Lỗi: Không tìm thấy cột 'Biển số xe' trong sheet Booking.", pd.DataFrame(), {}, pd.DataFrame()
        
        df_bk['Biển_Clean'] = df_bk['Biển số xe'].apply(normalize_plate)
        
        def classify_car(clean_plate):
            if not clean_plate or not is_valid_plate(clean_plate): return "Unknown"
            if clean_plate in driver_cars_map: return "Xe Nội bộ"
            return "Xe Vãng lai"

        def get_driver_info(clean_plate, col_type):
            if clean_plate in driver_cars_map:
                return driver_cars_map[clean_plate].get(col_type, 'Unknown')
            return None

        df_bk['Phân Loại Xe'] = df_bk['Biển_Clean'].apply(classify_car)
        
        # Merge thông tin chuẩn
        df_bk['Tên tài xế chuẩn'] = df_bk.apply(lambda x: get_driver_info(x['Biển_Clean'], 'Driver_Name') if x['Phân Loại Xe'] == 'Xe Nội bộ' else x['Tên tài xế'], axis=1)
        df_bk['Tên tài xế'] = df_bk['Tên tài xế chuẩn'].fillna('Chưa cập nhật')
        
        # 3. Merge CBNV
        if not df_cbnv.empty:
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
                df_bk = df_bk.merge(df_cbnv, left_on='Người sử dụng xe', right_on='Full Name', how='left')

        for c in ['Công ty', 'BU', 'Location']:
            if c not in df_bk.columns: df_bk[c] = 'Unknown'
            else: df_bk[c] = df_bk[c].fillna('Unknown').astype(str)

        # 4. Thời gian
        df_bk['Start'] = pd.to_datetime(df_bk['Ngày khởi hành'].astype(str) + ' ' + df_bk['Giờ khởi hành'].astype(str), errors='coerce')
        df_bk['End'] = pd.to_datetime(df_bk['Ngày khởi hành'].astype(str) + ' ' + df_bk['Giờ kết thúc'].astype(str), errors='coerce')
        df_bk.loc[df_bk['End'] < df_bk['Start'], 'End'] += pd.Timedelta(days=1)
        df_bk['Duration'] = (df_bk['End'] - df_bk['Start']).dt.total_seconds() / 3600
        df_bk['Tháng'] = df_bk['Start'].dt.strftime('%Y-%m')

        def check_scope(r):
            s = str(r).lower()
            provinces = ['bình dương', 'đồng nai', 'long an', 'bà rịa', 'vũng tàu', 'tây ninh', 'bình phước', 'tiền giang', 'bến tre', 'cần thơ', 'vĩnh long', 'an giang', 'bắc ninh', 'hưng yên', 'hải dương', 'hải phòng', 'vĩnh phúc', 'hà nam', 'nam định', 'thái bình', 'thái nguyên', 'hòa bình', 'bắc giang', 'phú thọ', 'thanh hóa', 'nghệ an']
            if any(p in s for p in provinces): return "Đi Tỉnh"
            return "Nội thành"
        
        if 'Lộ trình' in df_bk.columns:
            df_bk['Phạm Vi'] = df_bk['Lộ trình'].apply(check_scope)
        else:
            df_bk['Phạm Vi'] = 'Unknown'

        report_info = {
            'driver_cars_count': len(driver_cars_map),
            'duplicates_list': duplicates_check,
            'driver_cars_map': driver_cars_map
        }

        return df_bk, report_info, df_driver
    except Exception as e: return f"Lỗi: {str(e)}", pd.DataFrame(), {}, pd.DataFrame()

# --- 3. HÀM TẠO ẢNH CHO PPTX ---
def get_chart_img(data, x, y, kind='bar', title='', color='#0078d4'):
    plt.figure(figsize=(7, 4.5))
    
    # Check safe columns for plotting
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

# --- 4. EXPORT PPTX ---
def export_pptx(kpi, df_comp, df_status, top_users, top_drivers, df_bad_trips, chart_prefs, df_scope):
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

    # Slide 1
    slide = prs.slides.add_slide(prs.slide_layouts[0])
    slide.shapes.title.text = "BÁO CÁO VẬN HÀNH ĐỘI XE"
    slide.placeholders[1].text = f"Cập nhật: {kpi['last_month']}"
    
    # Slide 2: KPI
    slide = prs.slides.add_slide(prs.slide_layouts[5]); slide.shapes.title.text = "TỔNG QUAN HIỆU SUẤT"
    add_kpi_shape(slide, Inches(0.5), Inches(2.5), Inches(1.8), Inches(1.5), "TỔNG CHUYẾN", f"{kpi['trips']}", "Số chuyến", RGBColor(0, 120, 212))
    add_kpi_shape(slide, Inches(2.4), Inches(2.5), Inches(1.8), Inches(1.5), "GIỜ VẬN HÀNH", f"{kpi['hours']:,.0f}", "Tổng giờ", RGBColor(0, 120, 212))
    add_kpi_shape(slide, Inches(4.3), Inches(2.5), Inches(1.8), Inches(1.5), "CÔNG SUẤT", kpi['occupancy_text'], "Mục tiêu >50%", RGBColor(0, 120, 212))
    add_kpi_shape(slide, Inches(6.2), Inches(2.5), Inches(1.8), Inches(1.5), "HOÀN THÀNH", f"{kpi['success_rate']:.1f}%", "Tỷ lệ OK", RGBColor(16, 124, 16))
    add_kpi_shape(slide, Inches(8.1), Inches(2.5), Inches(1.8), Inches(1.5), "HỦY/TỪ CHỐI", f"{kpi['cancel_rate'] + kpi['reject_rate']:.1f}%", "Tỷ lệ Fail", RGBColor(209, 52, 56))

    # Slide 3: Charts
    slide = prs.slides.add_slide(prs.slide_layouts[5]); slide.shapes.title.text = "PHÂN TÍCH CẤU TRÚC SỬ DỤNG"
    if not df_comp.empty:
        img1 = get_chart_img(df_comp.head(8), 'Value', 'Category', kind=chart_prefs.get('structure', 'bar'), title='Top Đơn Vị')
        slide.shapes.add_picture(img1, Inches(0.5), Inches(1.8), Inches(4.5), Inches(3.5))
    if not df_scope.empty:
        img2 = get_chart_img(df_scope, 'Số lượng', 'Phạm vi', kind=chart_prefs.get('scope', 'pie'), title='Phạm Vi')
        slide.shapes.add_picture(img2, Inches(5.2), Inches(1.8), Inches(4.5), Inches(3.5))

    # Slide 4: Ranking
    slide = prs.slides.add_slide(prs.slide_layouts[5]); slide.shapes.title.text = "BẢNG XẾP HẠNG HOẠT ĐỘNG"
    if not top_users.empty:
        img_u = get_chart_img(top_users.head(8), 'Số chuyến', 'Người sử dụng xe', kind=chart_prefs.get('top_user', 'bar'), title='Top Users', color='#8764b8')
        slide.shapes.add_picture(img_u, Inches(0.5), Inches(1.8), Inches(4.5), Inches(3.5))
    if not top_drivers.empty:
        img_d = get_chart_img(top_drivers.head(8), 'Số chuyến', 'Tên tài xế', kind=chart_prefs.get('top_driver', 'bar'), title='Top Drivers', color='#00cc6a')
        slide.shapes.add_picture(img_d, Inches(5.2), Inches(1.8), Inches(4.5), Inches(3.5))

    # Slide 5: Bad Trips
    if not df_bad_trips.empty:
        slide = prs.slides.add_slide(prs.slide_layouts[5]); slide.shapes.title.text = "CHI TIẾT ĐƠN HỦY / TỪ CHỐI"
        
        # --- SAFE COLUMN SELECTION FOR PPTX ---
        wanted_cols = ['Start_Str', 'User', 'Status', 'Note', 'Lý do']
        avail_cols = [c for c in wanted_cols if c in df_bad_trips.columns]
        
        rows, cols = min(len(df_bad_trips)+1, 10), len(avail_cols)
        if cols > 0:
            table = slide.shapes.add_table(rows, cols, Inches(0.5), Inches(1.5), Inches(9), Inches(0.8)).table
            for i, h in enumerate(avail_cols):
                cell = table.cell(0, i); cell.text = h
                cell.fill.solid(); cell.fill.fore_color.rgb = RGBColor(0, 120, 212)
            for i, row in enumerate(df_bad_trips.head(9).itertuples(), start=1):
                for j, col_name in enumerate(avail_cols):
                    val = getattr(row, col_name, "")
                    table.cell(i, j).text = str(val)[:30] # Cắt ngắn nếu dài

    out = BytesIO(); prs.save(out); out.seek(0); return out

# --- 5. GIAO DIỆN CHÍNH ---
st.title("📊 Phước Minh - Hệ Thống Quản Trị & Tối Ưu Hóa Đội Xe")
uploaded_file = st.file_uploader("Upload Excel", type=['xlsx'], label_visibility="collapsed")

if uploaded_file:
    df, report_info, df_driver_raw = load_data_final(uploaded_file)
    if isinstance(df, str): st.error(df); st.stop()
    
    # SIDEBAR
    with st.sidebar:
        st.header("🗂️ Bộ Lọc Dữ Liệu")
        type_filter = st.multiselect("Loại Xe:", ["Xe Nội bộ", "Xe Vãng lai"], default=["Xe Nội bộ", "Xe Vãng lai"])
        
        min_date, max_date = df['Start'].min().date(), df['Start'].max().date()
        date_range = st.date_input("Thời gian:", (min_date, max_date), min_value=min_date, max_value=max_date)
        
        df_filtered = df.copy()
        if len(date_range) == 2:
            df_filtered = df_filtered[(df_filtered['Start'].dt.date >= date_range[0]) & (df_filtered['Start'].dt.date <= date_range[1])]
        
        if type_filter:
            df_filtered = df_filtered[df_filtered['Phân Loại Xe'].isin(type_filter)]

        st.markdown("---")
        st.caption("Drill-down:")
        locs = ["Tất cả"] + sorted(df_filtered['Location'].unique().tolist())
        sel_loc = st.selectbox("Khu vực:", locs)
        if sel_loc != "Tất cả": df_filtered = df_filtered[df_filtered['Location'] == sel_loc]
        
        comps = ["Tất cả"] + sorted(df_filtered['Công ty'].unique().tolist())
        sel_comp = st.selectbox("Công ty:", comps)
        if sel_comp != "Tất cả": df_filtered = df_filtered[df_filtered['Công ty'] == sel_comp]
        
        st.write(f"🔍 Đang xem: **{len(df_filtered)}** chuyến")

    if df_filtered.empty: st.warning("Không có dữ liệu."); st.stop()

    # --- KPI CALCULATION ---
    total_trips = len(df_filtered)
    total_hours = df_filtered['Duration'].sum()
    
    # Tính Occupancy CHỈ CHO XE NỘI BỘ
    internal_trips = df_filtered[df_filtered['Phân Loại Xe'] == 'Xe Nội bộ']
    hours_internal = internal_trips['Duration'].sum()
    
    if 'Xe Nội bộ' in type_filter:
        total_internal_fleet = report_info['driver_cars_count'] 
        active_internal_in_filter = internal_trips['Biển_Clean'].nunique()
        capacity_cars = total_internal_fleet if sel_loc == "Tất cả" else active_internal_in_filter
        if capacity_cars == 0: capacity_cars = 1
        
        days = max((df_filtered['Start'].max() - df_filtered['Start'].min()).days + 1, 1)
        cap = capacity_cars * days * 8
        occupancy_pct = (hours_internal / cap * 100) if cap > 0 else 0
        occupancy_text = f"{occupancy_pct:.1f}%"
    else:
        occupancy_text = "N/A"
        occupancy_pct = 0

    counts = df_filtered['Tình trạng đơn yêu cầu'].fillna('Unknown').value_counts()
    suc_rate = ((counts.get('CLOSED', 0) + counts.get('APPROVED', 0)) / total_trips * 100) if total_trips > 0 else 0
    fail_rate = ((counts.get('CANCELED', 0) + counts.get('CANCELLED', 0) + counts.get('REJECTED_BY_ADMIN', 0)) / total_trips * 100) if total_trips > 0 else 0

    # --- KPI UI ---
    cols = st.columns(5)
    cards = [
        {"title": "Tổng Chuyến", "val": f"{total_trips}", "sub": "∑ Đếm số dòng", "color": "#0078d4", "icon": "🚘", "is_percent": False},
        {"title": "Giờ Vận Hành", "val": f"{total_hours:,.0f}", "sub": "∑ (Giờ về - Giờ đi)", "color": "#0078d4", "icon": "⏱️", "is_percent": False},
        {"title": "Công Suất (Nội bộ)", "val": occupancy_text, "sub": f"Chỉ tính trên xe Cty", "color": "#0078d4", "icon": "📉", "is_percent": True, "pct_val": min(occupancy_pct, 100)},
        {"title": "Hoàn Thành", "val": f"{suc_rate:.1f}%", "sub": "Tỷ lệ thành công", "color": "#107c10", "icon": "✅", "is_percent": True, "pct_val": suc_rate},
        {"title": "Hủy / Từ Chối", "val": f"{fail_rate:.1f}%", "sub": "Tỷ lệ thất bại", "color": "#d13438", "icon": "🚫", "is_percent": True, "pct_val": fail_rate},
    ]
    for col, card in zip(cols, cards):
        progress_html = f'<div class="progress-bg"><div class="progress-fill" style="width: {card["pct_val"]}%; background-color: {card["color"]}"></div></div>' if card["is_percent"] else '<div style="height: 24px;"></div>'
        col.markdown(f"""<div class="kpi-card" style="border-top: 4px solid {card['color']}"><div class="kpi-header"><span class="kpi-title" style="color: {card['color']}">{card['title']}</span><span class="kpi-icon">{card['icon']}</span></div><div class="kpi-value">{card['val']}</div>{progress_html}<div class="kpi-formula">{card['sub']}</div></div>""", unsafe_allow_html=True)

    # --- TABS ---
    t1, t2, t3, t4 = st.tabs(["📊 Phân Tích", "🏆 Bảng Xếp Hạng", "📉 Chất Lượng", "⚙️ Đối Soát & Kiểm Tra"])
    
    chart_prefs = {} 
    kind_map = {"Thanh ngang (Bar)": "bar", "Thanh dọc (Column)": "column", "Tròn (Pie)": "pie"}

    with t1:
        c1, c2 = st.columns([2, 1])
        with c1:
            st.write("#### Phân bổ Loại Xe & Cấu trúc")
            chart_type_struct = st.selectbox("Kiểu biểu đồ Cấu trúc:", list(kind_map.keys()), index=0, key="c_struct")
            chart_prefs['structure'] = kind_map[chart_type_struct]
            
            if sel_comp == "Tất cả": df_g = df_filtered['Công ty'].value_counts().reset_index(); df_g.columns = ['Category', 'Value']; title_c = "Theo Công Ty"
            else: df_g = df_filtered['BU'].value_counts().reset_index(); df_g.columns = ['Category', 'Value']; title_c = f"Theo Phòng Ban ({sel_comp})"
            
            if chart_prefs['structure'] == "bar": fig = px.bar(df_g, x='Value', y='Category', orientation='h', text='Value', title=title_c)
            elif chart_prefs['structure'] == "column": fig = px.bar(df_g, x='Category', y='Value', text='Value', title=title_c)
            else: fig = px.pie(df_g, values='Value', names='Category', title=title_c)
            st.plotly_chart(fig, use_container_width=True)
            
        with c2:
            st.write("#### Phạm vi di chuyển")
            chart_type_scope = st.selectbox("Kiểu biểu đồ Phạm vi:", list(kind_map.keys()), index=2, key="c_scope")
            chart_prefs['scope'] = kind_map[chart_type_scope]
            
            if 'Phạm Vi' in df_filtered.columns:
                df_sc = df_filtered['Phạm Vi'].value_counts().reset_index(); df_sc.columns = ['Phạm vi', 'Số lượng']
                if chart_prefs['scope'] == "pie": fig_s = px.pie(df_sc, values='Số lượng', names='Phạm vi', hole=0.5)
                elif chart_prefs['scope'] == "bar": fig_s = px.bar(df_sc, x='Số lượng', y='Phạm vi', orientation='h', text='Số lượng')
                else: fig_s = px.bar(df_sc, x='Phạm vi', y='Số lượng', text='Số lượng')
                st.plotly_chart(fig_s, use_container_width=True)

    with t2:
        c_u, c_d = st.columns(2)
        with c_u:
            type_u = st.selectbox("Biểu đồ Top User:", list(kind_map.keys()), index=0, key="c_user")
            chart_prefs['top_user'] = kind_map[type_u]
            top_u = df_filtered.groupby(['Người sử dụng xe', 'Công ty']).size().reset_index(name='Số chuyến').sort_values('Số chuyến', ascending=False).head(10)
            st.write("##### 🥇 Top User")
            
            if chart_prefs['top_user'] == "bar": fig_u = px.bar(top_u, x='Số chuyến', y='Người sử dụng xe', orientation='h', text='Số chuyến', hover_data=['Công ty'])
            elif chart_prefs['top_user'] == "column": fig_u = px.bar(top_u, x='Người sử dụng xe', y='Số chuyến', text='Số chuyến')
            else: fig_u = px.pie(top_u, values='Số chuyến', names='Người sử dụng xe')
            st.plotly_chart(fig_u, use_container_width=True)
            st.dataframe(top_u, use_container_width=True)
            
        with c_d:
            type_d = st.selectbox("Biểu đồ Top Driver:", list(kind_map.keys()), index=0, key="c_driver")
            chart_prefs['top_driver'] = kind_map[type_d]
            top_d = df_filtered.groupby(['Tên tài xế', 'Phân Loại Xe']).size().reset_index(name='Số chuyến').sort_values('Số chuyến', ascending=False).head(10)
            st.write("##### 🚘 Top Driver")
            
            if chart_prefs['top_driver'] == "bar": fig_d = px.bar(top_d, x='Số chuyến', y='Tên tài xế', orientation='h', text='Số chuyến', hover_data=['Phân Loại Xe'])
            elif chart_prefs['top_driver'] == "column": fig_d = px.bar(top_d, x='Tên tài xế', y='Số chuyến', text='Số chuyến')
            else: fig_d = px.pie(top_d, values='Số chuyến', names='Tên tài xế')
            st.plotly_chart(fig_d, use_container_width=True)
            st.dataframe(top_d, use_container_width=True)

    with t3:
        st.write("#### Chi tiết Hủy / Từ chối")
        bad = df_filtered[df_filtered['Tình trạng đơn yêu cầu'].isin(['CANCELED', 'CANCELLED', 'REJECTED_BY_ADMIN'])]
        
        # --- SAFE COLUMN SELECTION (FIX KEYERROR) ---
        desired_cols = ['Ngày khởi hành', 'Biển số xe', 'Phân Loại Xe', 'Lý do', 'Note', 'Tình trạng đơn yêu cầu']
        actual_cols = [c for c in desired_cols if c in bad.columns]
        
        if not bad.empty:
            st.dataframe(bad[actual_cols], use_container_width=True)
        else:
            st.success("Không có chuyến nào bị hủy trong giai đoạn này.")

    with t4:
        st.subheader("⚙️ Chi Tiết Đối Soát Dữ Liệu")
        
        with st.expander("🚨 Kiểm tra Trùng lặp trong danh sách Driver", expanded=True):
            if report_info['duplicates_list']:
                st.error(f"Phát hiện {len(report_info['duplicates_list'])} biển số bị nhập trùng trong file Driver!")
                st.write(report_info['duplicates_list'])
            else:
                st.success("Dữ liệu Driver sạch, không có biển số trùng.")

        with st.expander(f"🚗 Danh sách Xe Vãng Lai"):
            vang_lai = df_filtered[df_filtered['Phân Loại Xe'] == 'Xe Vãng lai']['Biển số xe'].unique()
            st.write(f"Tìm thấy **{len(vang_lai)}** xe vãng lai trong bộ lọc hiện tại:")
            st.table(pd.DataFrame(vang_lai, columns=['Biển số Vãng lai']))

        with st.expander(f"🚙 Danh sách Xe Nội Bộ"):
            noi_bo = df_filtered[df_filtered['Phân Loại Xe'] == 'Xe Nội bộ']['Biển số xe'].unique()
            st.write(f"Tìm thấy **{len(noi_bo)}** xe nội bộ hoạt động:")
            st.table(pd.DataFrame(noi_bo, columns=['Biển số Nội bộ']))

    # --- PPTX ---
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

        kpi_data = {'trips': total_trips, 'hours': total_hours, 'occupancy': occupancy_pct, 'occupancy_text': occupancy_text, 'success_rate': suc_rate, 'cancel_rate': fail_rate, 'reject_rate': 0, 'last_month': last_month_str}
        
        df_status_exp = counts.reset_index(); df_status_exp.columns = ['Trạng thái', 'Số lượng']
        if sel_comp == "Tất cả": df_comp_exp = df_filtered['Công ty'].value_counts().reset_index(); df_comp_exp.columns=['Category', 'Value']
        else: df_comp_exp = df_filtered['BU'].value_counts().reset_index(); df_comp_exp.columns=['Category', 'Value']
        
        if 'Phạm Vi' in df_filtered.columns: df_scope_exp = df_filtered['Phạm Vi'].value_counts().reset_index(); df_scope_exp.columns = ['Phạm vi', 'Số lượng']
        else: df_scope_exp = pd.DataFrame(columns=['Phạm vi', 'Số lượng'])
        
        df_bad_exp = pd.DataFrame()
        if not bad.empty:
            df_bad_exp = bad.copy()
            df_bad_exp['Start_Str'] = df_bad_exp['Start'].dt.strftime('%d/%m')
            df_bad_exp = df_bad_exp.rename(columns={'Người sử dụng xe': 'User', 'Tình trạng đơn yêu cầu': 'Status'})

        pptx_file = export_pptx(kpi_data, df_comp_exp, df_status_exp, top_u, top_d, df_bad_exp, pptx_options, chart_prefs, df_scope_exp)
        st.download_button(label="Tải file .PPTX ngay", data=pptx_file, file_name="Bao_Cao_Van_Hanh_Full.pptx", mime="application/vnd.openxmlformats-officedocument.presentationml.presentation", type="primary")

else:
    st.info("👋 Vui lòng upload file Excel dữ liệu.")