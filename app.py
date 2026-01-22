import streamlit as st
import pandas as pd
import plotly.express as px
import numpy as np

# --- CẤU HÌNH GIAO DIỆN ---
st.set_page_config(page_title="Smart Fleet Dashboard", page_icon="🚀", layout="wide")
st.markdown("""
<style>
    .header-style {font-size: 26px; font-weight: bold; color: #2c3e50;}
    .sub-header {font-size: 18px; color: #7f8c8d;}
    div[data-testid="stMetricValue"] {color: #2980b9;}
</style>
""", unsafe_allow_html=True)

st.markdown("<div class='header-style'>🚀 Dashboard Đội Xe Thông Minh (AI Powered)</div>", unsafe_allow_html=True)
st.markdown("---")

# --- HÀM XỬ LÝ DỮ LIỆU ---
@st.cache_data
def load_and_process_data(file):
    try:
        if file.name.endswith('.csv'):
            df = pd.read_csv(file)
        else:
            df = pd.read_excel(file, engine='openpyxl')
        
        # Chuẩn hóa tên cột
        df.columns = df.columns.str.strip()
        
        # Xử lý Ngày Giờ (Bắt buộc phải có)
        try:
            df['Start_Datetime'] = pd.to_datetime(df['Ngày khởi hành'].astype(str) + ' ' + df['Giờ khởi hành'].astype(str), errors='coerce')
            df['End_Datetime'] = pd.to_datetime(df['Ngày khởi hành'].astype(str) + ' ' + df['Giờ kết thúc'].astype(str), errors='coerce')
            
            mask_overnight = df['End_Datetime'] < df['Start_Datetime']
            df.loc[mask_overnight, 'End_Datetime'] += pd.Timedelta(days=1)
            
            df['Thời lượng (Giờ)'] = (df['End_Datetime'] - df['Start_Datetime']).dt.total_seconds() / 3600
            df['Tháng'] = df['Start_Datetime'].dt.to_period('M').astype(str)
        except:
            pass # Nếu lỗi ngày giờ thì bỏ qua, vẫn load các cột khác để tính toán
            
        return df
    except Exception as e:
        return str(e)

# --- UPLOAD ---
uploaded_file = st.file_uploader("📂 Upload file Excel/CSV", type=['xlsx', 'csv'])

if uploaded_file is not None:
    df = load_and_process_data(uploaded_file)
    if isinstance(df, str):
        st.error(f"Lỗi: {df}")
        st.stop()

    # --- SIDEBAR ---
    with st.sidebar:
        st.header("🔍 Bộ Lọc Nhanh")
        if 'Start_Datetime' in df.columns:
            min_d = df['Start_Datetime'].min().date()
            max_d = df['End_Datetime'].max().date()
            d_range = st.date_input("Thời gian:", (min_d, max_d))
            # Lọc dataframe
            if len(d_range) == 2:
                 df = df[(df['Start_Datetime'].dt.date >= d_range[0]) & (df['Start_Datetime'].dt.date <= d_range[1])]
        
        st.info(f"Đang xử lý: {len(df)} dòng dữ liệu")

    # --- TABS CHÍNH ---
    tab1, tab2, tab3, tab4 = st.tabs(["📊 Tổng Quan", "🏢 Đơn Vị & User", "⚠️ Check Trùng", "🧮 Tự Tính Toán (New)"])

    # --- TAB 1: TỔNG QUAN ---
    with tab1:
        if 'Thời lượng (Giờ)' in df.columns:
            col1, col2 = st.columns(2)
            col1.metric("Tổng số chuyến", len(df))
            col1.metric("Tổng giờ chạy", f"{df['Thời lượng (Giờ)'].sum():,.1f}h")
            
            # Biểu đồ diễn biến
            daily_usage = df.groupby('Tháng')['Thời lượng (Giờ)'].sum().reset_index()
            fig = px.bar(daily_usage, x='Tháng', y='Thời lượng (Giờ)', title="Xu hướng sử dụng theo tháng")
            st.plotly_chart(fig, use_container_width=True)
        else:
            st.warning("Không tính được KPI do thiếu cột Ngày/Giờ chuẩn.")

    # --- TAB 2: ĐƠN VỊ ---
    with tab2:
        # Tự động tìm cột Bộ phận / Công ty
        cols_to_plot = [c for c in df.columns if c in ['Bộ phận', 'Công ty', 'Cost center', 'Người sử dụng xe']]
        if cols_to_plot:
            selected_col = st.selectbox("Chọn tiêu chí vẽ biểu đồ:", cols_to_plot)
            counts = df[selected_col].value_counts().reset_index().head(10)
            counts.columns = [selected_col, 'Số chuyến']
            fig2 = px.bar(counts, x='Số chuyến', y=selected_col, orientation='h', title=f"Top {selected_col}")
            st.plotly_chart(fig2, use_container_width=True)
        else:
            st.info("Chưa tìm thấy cột Bộ phận/Công ty/Người dùng phù hợp.")

    # --- TAB 3: CHECK TRÙNG ---
    with tab3:
        if 'Biển số xe' in df.columns and 'Start_Datetime' in df.columns:
            df_s = df.dropna(subset=['Biển số xe']).sort_values(['Biển số xe', 'Start_Datetime'])
            df_s['Prev_End'] = df_s.groupby('Biển số xe')['End_Datetime'].shift(1)
            overlaps = df_s[df_s['Start_Datetime'] < df_s['Prev_End']]
            
            if not overlaps.empty:
                st.error(f"Phát hiện {len(overlaps)} chuyến bị trùng!")
                st.dataframe(overlaps[['Ngày khởi hành', 'Biển số xe', 'Tên tài xế', 'Start_Datetime', 'End_Datetime', 'Prev_End']])
            else:
                st.success("Không có chuyến nào bị trùng.")

    # --- TAB 4: TỰ TÍNH TOÁN (TÍNH NĂNG MỚI) ---
    with tab4:
        st.subheader("🛠️ Công cụ Tự Tạo Công Thức")
        st.markdown("Bạn có thể tự chọn 2 cột số bất kỳ để cộng trừ nhân chia và xem kết quả.")
        
        # 1. Lọc ra các cột chứa số (Numeric columns only)
        numeric_cols = df.select_dtypes(include=[np.number]).columns.tolist()
        
        if len(numeric_cols) < 2:
            st.warning("File không đủ các cột dữ liệu số để tính toán (Cần ít nhất 2 cột số).")
        else:
            c1, c2, c3, c4 = st.columns([3, 1, 3, 2])
            
            with c1:
                col_a = st.selectbox("Chọn Cột A (Số):", numeric_cols, index=0)
            with c2:
                operator = st.selectbox("Phép tính:", ["+", "-", "*", "/"])
            with c3:
                # Cho phép chọn Cột B hoặc nhập một số cố định
                input_mode = st.radio("Cột B là:", ["Một Cột Khác", "Số Cố Định (VD: 1000)"], horizontal=True)
                if input_mode == "Một Cột Khác":
                    col_b = st.selectbox("Chọn Cột B (Số):", numeric_cols, index=1 if len(numeric_cols)>1 else 0)
                    val_b = None
                else:
                    col_b = None
                    val_b = st.number_input("Nhập số:", value=1.0)
            
            with c4:
                st.write("") # Spacer
                st.write("")
                calc_btn = st.button("🚀 Tính & Vẽ Biểu Đồ", type="primary")

            # Xử lý tính toán khi bấm nút
            if calc_btn:
                new_col_name = f"Kết quả ({col_a} {operator} {col_b if col_b else val_b})"
                
                try:
                    # Thực hiện phép tính
                    if operator == "+":
                        res = df[col_a] + (df[col_b] if col_b else val_b)
                    elif operator == "-":
                        res = df[col_a] - (df[col_b] if col_b else val_b)
                    elif operator == "*":
                        res = df[col_a] * (df[col_b] if col_b else val_b)
                    elif operator == "/":
                        # Xử lý chia cho 0
                        divisor = df[col_b] if col_b else val_b
                        res = df[col_a] / divisor.replace(0, np.nan)
                    
                    # Thêm vào dataframe tạm
                    df[new_col_name] = res
                    
                    st.success(f"Đã tính xong! Tạo cột mới: '{new_col_name}'")
                    
                    # Hiển thị thống kê
                    m1, m2 = st.columns(2)
                    m1.metric("Tổng cộng (Sum)", f"{res.sum():,.2f}")
                    m2.metric("Trung bình (Mean)", f"{res.mean():,.2f}")
                    
                    # Vẽ biểu đồ kết quả
                    st.subheader("Biểu đồ phân bố kết quả")
                    
                    # Cho chọn trục X để vẽ (ví dụ theo Tháng hoặc theo Công ty)
                    x_axis_options = [c for c in df.columns if df[c].dtype == 'object'] # Cột chữ
                    if not x_axis_options: x_axis_options = ['index']
                    
                    x_axis = st.selectbox("Chọn trục hoành (X) để nhóm dữ liệu:", x_axis_options, index=0)
                    
                    # Gom nhóm và vẽ
                    chart_data = df.groupby(x_axis)[new_col_name].sum().reset_index()
                    fig_calc = px.bar(chart_data, x=x_axis, y=new_col_name, title=f"Biểu đồ {new_col_name} theo {x_axis}")
                    st.plotly_chart(fig_calc, use_container_width=True)
                    
                    # Hiện bảng dữ liệu chi tiết
                    with st.expander("Xem bảng dữ liệu chi tiết"):
                        st.dataframe(df[[col_a, col_b] + [new_col_name] if col_b else df[[col_a, new_col_name]]])

                except Exception as e:
                    st.error(f"Lỗi tính toán: {e}")

else:
    st.info("👈 Hãy upload file để trải nghiệm tính năng AI")