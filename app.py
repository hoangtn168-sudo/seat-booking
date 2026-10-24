import streamlit as st
import pandas as pd
import os

# Cấu hình giao diện trang web rộng
st.set_page_config(page_title="Hệ Thống Đăng Ký Chỗ Ngồi Online", layout="wide", initial_sidebar_state="expanded")

FILE_PATH = "Danh_Sach_Cho_Ngoi.xlsx"

# Đường dẫn các file ảnh sơ đồ tổng quan & chi tiết
IMG_CHUNG = "so_do_chung.png"
IMG_ZONE_A = "zone_a.png"
IMG_ZONE_B = "zone_b.png"
IMG_ZONE_C = "zone_c.png"

# Danh sách đầy đủ 9 phòng ban (BU)
LIST_BU = [
    'IOT', 
    'C&D', 
    'Product Engineering', 
    'SAP', 
    'Workday', 
    'Japan Desk', 
    'GTM Sales Team', 
    'IT Service Desk', 
    'Facilities Real Estate Support'
]

# Hàm đọc và chuẩn hóa dữ liệu từ file Excel
def load_data():
    if not os.path.exists(FILE_PATH):
        seats = []
        for i in range(1, 29):
            seats.append({'Mã Chỗ Ngồi': f'A-{i:02d}', 'Phân Khu Mặt Bằng': 'Zone A', 'Họ Và Tên': '', 'LDAP ID': '', 'Phòng Ban (BU)': 'IOT', 'Ghi Chú Bố Trí': ''})
        for i in range(1, 27):
            seats.append({'Mã Chỗ Ngồi': f'B-{i:02d}', 'Phân Khu Mặt Bằng': 'Zone B', 'Họ Và Tên': '', 'LDAP ID': '', 'Phòng Ban (BU)': 'C&D', 'Ghi Chú Bố Trí': ''})
        for i in range(1, 26):
            seats.append({'Mã Chỗ Ngồi': f'C-{i:02d}', 'Phân Khu Mặt Bằng': 'Zone C', 'Họ Và Tên': '', 'LDAP ID': '', 'Phòng Ban (BU)': 'C&D', 'Ghi Chú Bố Trí': ''})
        df_init = pd.DataFrame(seats)
        df_init.to_excel(FILE_PATH, index=False)

    df = pd.read_excel(FILE_PATH, dtype=str)
    
    column_mapping = {
        'Phân Khu Mặt Bằng': 'Phân Khu',
        'Phân khu': 'Phân Khu',
        'Mã chỗ ngồi': 'Mã Chỗ Ngồi',
        'Họ và tên': 'Họ Và Tên',
        'Phòng ban (BU)': 'Phòng Ban (BU)',
        'Phòng Ban': 'Phòng Ban (BU)',
        'Ghi Chú Bố Trí': 'Ghi Chú',
        'Ghi chú': 'Ghi Chú'
    }
    df = df.rename(columns=column_mapping)
    
    for col in ['Họ Và Tên', 'LDAP ID', 'Phòng Ban (BU)', 'Ghi Chú']:
        if col in df.columns:
            df[col] = df[col].fillna('').astype(str).str.strip()
            df[col] = df[col].replace(['nan', 'None', 'NaN', 'nat', '<NA>', '0'], '')
        else:
            df[col] = ''

    df['LDAP ID'] = df['LDAP ID'].apply(lambda x: x[:-2] if x.endswith('.0') else x)

    if 'Phân Khu' not in df.columns:
        def detect_zone(code):
            code_str = str(code).upper()
            if code_str.startswith('A'): return 'Zone A'
            elif code_str.startswith('B'): return 'Zone B'
            elif code_str.startswith('C'): return 'Zone C'
            return 'Zone A'
        df['Phân Khu'] = df['Mã Chỗ Ngồi'].apply(detect_zone)
        
    return df

def save_data(df):
    df.to_excel(FILE_PATH, index=False)

df = load_data()

st.title("HỆ THỐNG ĐĂNG KÝ CHỖ NGỒI VĂN PHÒNG")
st.info("Hướng dẫn: Xem sơ đồ dạng bảng ma trận bên dưới -> Bấm nút vào ô ghế muốn chọn -> Điền thông tin ở khung đăng ký hiện ra.")

# --- HIỂN THỊ SƠ ĐỒ CHUNG TRỰC QUAN ---
if os.path.exists(IMG_CHUNG):
    with st.expander("Xem sơ đồ tổng quan văn phòng (Ảnh gốc)", expanded=False):
        st.image(IMG_CHUNG, caption="Sơ đồ tổng quan toàn văn phòng", use_container_width=True)

# Xác định trạng thái của ghế dựa trên Ghi Chú và Họ Tên
def get_status(row):
    note = str(row['Ghi Chú']).lower()
    name = str(row['Họ Và Tên'])
    
    if any(k in note for k in ["khóa", "hỏng", "không"]):
        return "XÁM"
    elif name != "" and name.lower() not in ["nan", "none", "", "0"]:
        return "XANH"
    else:
        return "VÀNG"

# Hàm render một ô ghế chuẩn mẫu ma trận
def render_seat_box(row):
    if row is None:
        st.markdown("<div style='min-height: 85px;'></div>", unsafe_allow_html=True)
        return
        
    code = row['Mã Chỗ Ngồi']
    name = row['Họ Và Tên'] if row['Họ Và Tên'] not in ["", "nan", "None", "0"] else "Trống (0)"
    bu = row['Phòng Ban (BU)'] if row['Phòng Ban (BU)'] not in ["", "nan", "None"] else "Chưa chọn BU"
    status = get_status(row)
    
    if status == "XANH":
        st.markdown(f"""
        <div style="border: 2px solid #28a745; background-color: #e8f5e9; padding: 6px; border-radius: 6px; text-align: center; margin-bottom: 4px; min-height: 80px;">
            <b style="color: #155724; font-size: 13px;">{code}</b><br>
            <span style="color: #333; font-size: 11px; font-weight: bold;">{name}</span><br>
            <span style="color: #666; font-size: 9px; font-style: italic;">{bu}</span>
        </div>
        """, unsafe_allow_html=True)
        if st.button("Đổi chỗ", key=f"change_{code}", use_container_width=True):
            st.session_state['selected_seat'] = code
            st.rerun()
            
    elif status == "VÀNG":
        st.markdown(f"""
        <div style="border: 2px solid #ffc107; background-color: #fff9c4; padding: 6px; border-radius: 6px; text-align: center; margin-bottom: 4px; min-height: 80px;">
            <b style="color: #856404; font-size: 13px;">{code}</b><br>
            <span style="color: #888; font-size: 11px;">Trống (0)</span><br>
            <span style="color: #666; font-size: 9px; font-style: italic;">{bu}</span>
        </div>
        """, unsafe_allow_html=True)
        if st.button("Chọn", key=f"btn_{code}", use_container_width=True):
            st.session_state['selected_seat'] = code
            st.rerun()
    else:
        st.markdown(f"""
        <div style="border: 2px solid #6c757d; background-color: #e2e3e5; padding: 6px; border-radius: 6px; text-align: center; margin-bottom: 4px; min-height: 80px;">
            <b style="color: #383d41; font-size: 13px;">{code}</b><br>
            <span style="color: #721c24; font-size: 11px; font-weight: bold;">Khóa</span>
        </div>
        """, unsafe_allow_html=True)

# --- KHUNG FORM ĐĂNG KÝ NGAY TRÊN MÀN HÌNH CHÍNH ---
selected_seat = st.session_state.get('selected_seat', '')

if selected_seat:
    st.markdown("---")
    with st.container():
        st.success(f"BẠN ĐANG CHỌN VỊ TRÍ: {selected_seat}")
        
        col_f1, col_f2 = st.columns(2)
        with col_f1:
            user_ldap = st.text_input("Nhập LDAP ID (Mã nhân viên): *", key="input_ldap")
            user_name = st.text_input("Nhập Họ Và Tên: *", key="input_name")
        with col_f2:
            user_bu = st.selectbox("Chọn Phòng Ban (BU):", LIST_BU, key="input_bu")
            st.write("")
            st.write("")
        
        col_btn1, col_btn2 = st.columns([1, 4])
        with col_btn1:
            if st.button("Xác Nhận Đặt", type="primary"):
                if not user_ldap.strip() or not user_name.strip():
                    st.warning("Vui lòng nhập đầy đủ LDAP ID và Họ Tên!")
                else:
                    clean_ldap = user_ldap.strip()
                    existing_booking = df[df['LDAP ID'] == clean_ldap]
                    
                    if not existing_booking.empty:
                        booked_seat = existing_booking.iloc[0]['Mã Chỗ Ngồi']
                        st.error(f"Chặn spam! LDAP {clean_ldap} đã đăng ký ghế {booked_seat} rồi. Mỗi nhân sự chỉ được sở hữu 1 chỗ duy nhất!")
                    else:
                        df.loc[df['Mã Chỗ Ngồi'] == selected_seat, 'Họ Và Tên'] = user_name.strip()
                        df.loc[df['Mã Chỗ Ngồi'] == selected_seat, 'LDAP ID'] = clean_ldap
                        df.loc[df['Mã Chỗ Ngồi'] == selected_seat, 'Phòng Ban (BU)'] = user_bu
                        save_data(df)
                        
                        st.success(f"Đăng ký thành công ghế {selected_seat} cho {user_name} ({user_bu})!")
                        st.session_state['selected_seat'] = ''
                        st.rerun()
        with col_btn2:
            if st.button("Hủy chọn lại"):
                st.session_state['selected_seat'] = ''
                st.rerun()
    st.markdown("---")

st.subheader("SƠ ĐỒ BÀN LÀM VIỆC CHI TIẾT THEO ZONE")

# Các Tab phân khu
tab1, tab2, tab3 = st.tabs(["Zone A (28 Bàn)", "Zone B (26 Bàn)", "Zone C (25 Bàn)"])

# Hàm lấy thông tin ghế theo mã
def get_seat_row(zone_name, seat_code):
    sub = df[(df['Phân Khu'] == zone_name) & (df['Mã Chỗ Ngồi'] == seat_code)]
    if not sub.empty:
        return sub.iloc[0]
    return None

with tab1:
    if os.path.exists(IMG_ZONE_A):
        with st.expander("Xem ảnh sơ đồ gốc Zone A"):
            st.image(IMG_ZONE_A, caption="Sơ đồ chi tiết Zone A", use_container_width=True)
        st.markdown("---")
    
    zone_a_df = df[df['Phân Khu'].astype(str).str.startswith('Zone A')].reset_index(drop=True)
    for i in range(0, len(zone_a_df), 7):
        cols = st.columns(7)
        for j in range(7):
            if i + j < len(zone_a_df):
                with cols[j]:
                    render_seat_box(zone_a_df.iloc[i + j])

with tab2:
    if os.path.exists(IMG_ZONE_B):
        with st.expander("Xem ảnh sơ đồ gốc Zone B"):
            st.image(IMG_ZONE_B, caption="Sơ đồ chi tiết Zone B", use_container_width=True)
        st.markdown("---")
        
    rows_b = [
        [f"B-{i:02d}" for i in range(1, 9)],
        [f"B-{i:02d}" for i in range(9, 17)],
        [f"B-{i:02d}" for i in range(17, 23)] + [None, None],
        [f"B-{i:02d}" for i in range(23, 27)] + [None, None, None, None]
    ]
    
    for row_codes in rows_b:
        cols = st.columns(8)
        for j, code in enumerate(row_codes):
            with cols[j]:
                if code:
                    render_seat_box(get_seat_row('Zone B', code))
                else:
                    st.markdown("")

with tab3:
    if os.path.exists(IMG_ZONE_C):
        with st.expander("Xem ảnh sơ đồ gốc Zone C"):
            st.image(IMG_ZONE_C, caption="Sơ đồ chi tiết Zone C", use_container_width=True)
        st.markdown("---")
        
    # Sơ đồ Zone C cố định chuẩn 10 cột, đảm bảo cột C-06/15/24 và C-07/16/25 thẳng hàng tuyệt đối
    rows_c_exact = [
        [f"C-01", f"C-02", f"C-03", f"C-04", f"C-05", None, None, None, f"C-06", f"C-07"],
        [f"C-08", f"C-09", f"C-10", f"C-11", f"C-12", None, f"C-13", f"C-14", f"C-15", f"C-16"],
        [f"C-17", f"C-18", f"C-19", f"C-20", f"C-21", f"C-22", f"C-23", f"C-24", f"C-25", None]
    ]
    
    for row_codes in rows_c_exact:
        cols = st.columns(10)
        for j, code in enumerate(row_codes):
            with cols[j]:
                if code:
                    render_seat_box(get_seat_row('Zone C', code))
                else:
                    st.markdown("")

# --- TẢI FILE EXCEL (Dành cho Quản trị viên) ---
st.markdown("---")
with st.expander("Dành cho Quản trị viên (Admin)"):
    with open(FILE_PATH, "rb") as f:
        st.download_button("Tải File Excel Danh Sách Chỗ Ngồi", f, file_name="Danh_Sach_Cho_Ngoi_Cap_Nhat.xlsx")
