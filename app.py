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

# Định nghĩa BU cố định cho từng vị trí ở Zone A
ZONE_A_PRESET_BU = {
    'A-01': 'HCVN.GTM Sales Team.Sales',
    'A-02': 'HCVN.GTM Sales Team.Sales',
    'A-03': 'HCVN.IT Service Desk.HD Costs',
    'A-04': 'IOT', 'A-05': 'IOT', 'A-06': 'IOT', 'A-07': 'IOT',
    'A-08': 'HCVN.Japan Desk.Delivery',
    'A-09': 'HCVN.Japan Desk.Delivery',
    'A-10': 'HCVN.Facilities Real Estate Support.HD Costs',
    'A-11': 'IOT', 'A-12': 'IOT', 'A-13': 'IOT', 'A-14': 'IOT',
    'A-15': 'Product Engineering.Delivery', 'A-16': 'Product Engineering.Delivery', 'A-17': 'Product Engineering.Delivery',
    'A-18': 'IOT', 'A-19': 'IOT', 'A-20': 'IOT', 'A-21': 'IOT',
    'A-22': 'Product Engineering.Delivery', 'A-23': 'Product Engineering.Delivery', 'A-24': 'Product Engineering.Delivery',
    'A-25': 'IOT', 'A-26': 'IOT', 'A-27': 'IOT', 'A-28': 'IOT'
}

# Định nghĩa BU cố định cho Zone C theo ảnh mẫu
ZONE_C_PRESET_BU = {}
for code in ['C-01', 'C-02', 'C-08', 'C-09', 'C-17', 'C-18']:
    ZONE_C_PRESET_BU[code] = 'C&D'
for code in ['C-03', 'C-04', 'C-05', 'C-10', 'C-11', 'C-12', 'C-19', 'C-20', 'C-21']:
    ZONE_C_PRESET_BU[code] = 'Workday.Delivery'
for code in ['C-06', 'C-07', 'C-13', 'C-14', 'C-15', 'C-16', 'C-22', 'C-23', 'C-24', 'C-25']:
    ZONE_C_PRESET_BU[code] = 'SAP'

# Hàm đọc và chuẩn hóa dữ liệu từ file Excel
def load_data():
    if not os.path.exists(FILE_PATH):
        seats = []
        for i in range(1, 29):
            code = f'A-{i:02d}'
            default_bu = ZONE_A_PRESET_BU.get(code, 'IOT')
            seats.append({'Mã Chỗ Ngồi': code, 'Phân Khu Mặt Bằng': 'Zone A', 'Họ Và Tên': '', 'LDAP ID': '', 'Phòng Ban (BU)': default_bu, 'Ghi Chú Bố Trí': ''})
        for i in range(1, 27):
            seats.append({'Mã Chỗ Ngồi': f'B-{i:02d}', 'Phân Khu Mặt Bằng': 'Zone B', 'Họ Và Tên': '', 'LDAP ID': '', 'Phòng Ban (BU)': 'C&D', 'Ghi Chú Bố Trí': ''})
        for i in range(1, 26):
            code = f'C-{i:02d}'
            default_bu = ZONE_C_PRESET_BU.get(code, 'C&D')
            seats.append({'Mã Chỗ Ngồi': code, 'Phân Khu Mặt Bằng': 'Zone C', 'Họ Và Tên': '', 'LDAP ID': '', 'Phòng Ban (BU)': default_bu, 'Ghi Chú Bố Trí': ''})
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

    # Đảm bảo Zone A và Zone C luôn cập nhật đúng BU định sẵn
    for idx, row in df.iterrows():
        code = row['Mã Chỗ Ngồi']
        if code.startswith('A-') and code in ZONE_A_PRESET_BU:
            if not row['Phòng Ban (BU)'] or row['Phòng Ban (BU)'] in ['', '0']:
                df.at[idx, 'Phòng Ban (BU)'] = ZONE_A_PRESET_BU[code]
        elif code.startswith('B-'):
            if not row['Phòng Ban (BU)'] or row['Phòng Ban (BU)'] in ['', '0']:
                df.at[idx, 'Phòng Ban (BU)'] = 'C&D'
        elif code.startswith('C-') and code in ZONE_C_PRESET_BU:
            if not row['Phòng Ban (BU)'] or row['Phòng Ban (BU)'] in ['', '0']:
                df.at[idx, 'Phòng Ban (BU)'] = ZONE_C_PRESET_BU[code]

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
st.info("💡 **Hướng dẫn:** Bấm nút **Chọn** ở ghế trống để tiến hành đăng ký thông tin.")

# --- HIỂN THỊ SƠ ĐỒ CHUNG TRỰC QUAN ---
if os.path.exists(IMG_CHUNG):
    with st.expander("🗺️ Xem sơ đồ tổng quan văn phòng", expanded=False):
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
        <div style="border: 2px solid #28a745; background-color: #e8f5e9; padding: 6px; border-radius: 6px; text-align: center; margin-bottom: 2px; min-height: 72px;">
            <b style="color: #155724; font-size: 13px;">{code}</b><br>
            <span style="color: #333; font-size: 11px; font-weight: bold;">{name}</span><br>
            <span style="color: #666; font-size: 9px; font-style: italic;">{bu}</span>
        </div>
        """, unsafe_allow_html=True)
        st.markdown("<p style='text-align: center; font-size: 10px; color: #d9534f; font-weight: bold; margin: 0 0 5px 0;'>Liên hệ admin để đổi chỗ</p>", unsafe_allow_html=True)
            
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
    seat_row = df[df['Mã Chỗ Ngồi'] == selected_seat]
    preset_bu = seat_row.iloc[0]['Phòng Ban (BU)'] if not seat_row.empty else ''
    
    zone_prefix = selected_seat.split('-')[0]
    zone_name = f"Zone {zone_prefix}"

    with st.container():
        st.success(f"BẠN ĐANG CHỌN VỊ TRÍ: {selected_seat}")
        st.info(f"ℹ️ **Lưu ý {zone_name}:** Vị trí **{selected_seat}** được định sẵn phòng ban (BU): **{preset_bu}**. Bạn không thể thay đổi phòng ban khi đặt chỗ tại đây.")
        
        col_f1, col_f2 = st.columns(2)
        with col_f1:
            user_ldap = st.text_input("Nhập LDAP ID (Mã nhân viên): *", key="input_ldap")
            user_name = st.text_input("Nhập Họ Và Tên: *", key="input_name")
        with col_f2:
            st.text_input("Phòng Ban (BU) - Cố định theo vị trí:", value=preset_bu, disabled=True, key="input_bu_fixed")
            user_bu = preset_bu
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

# Các Tab phân khu (Thêm tab Admin trực tuyến)
tab1, tab2, tab3, tab_admin = st.tabs(["Zone A (28 Bàn)", "Zone B (26 Bàn)", "Zone C (25 Bàn)", "🛠️ Quản Trị Viên (Admin)"])

# Hàm lấy thông tin ghế theo mã
def get_seat_row(zone_name, seat_code):
    sub = df[(df['Phân Khu'] == zone_name) & (df['Mã Chỗ Ngồi'] == seat_code)]
    if not sub.empty:
        return sub.iloc[0]
    return None

with tab1:
    if os.path.exists(IMG_ZONE_A):
        with st.expander("📷 Xem ảnh sơ đồ gốc Zone A"):
            st.image(IMG_ZONE_A, caption="Sơ đồ chi tiết Zone A", use_container_width=True)
        st.markdown("---")
    
    st.info("ℹ️ **Quy định Zone A:** Khu vực này chỉ bao gồm các phòng ban (BU): **GTM Sales Team, IT Service Desk, Japan Desk, Facilities Real Estate Support, Product Engineering, IOT**. Các vị trí đã được cố định sẵn BU tương ứng.")
    
    zone_a_df = df[df['Phân Khu'].astype(str).str.startswith('Zone A')].reset_index(drop=True)
    for i in range(0, len(zone_a_df), 7):
        cols = st.columns(7)
        for j in range(7):
            if i + j < len(zone_a_df):
                with cols[j]:
                    render_seat_box(zone_a_df.iloc[i + j])

with tab2:
    if os.path.exists(IMG_ZONE_B):
        with st.expander("📷 Xem ảnh sơ đồ gốc Zone B"):
            st.image(IMG_ZONE_B, caption="Sơ đồ chi tiết Zone B", use_container_width=True)
        st.markdown("---")
        
    st.info("ℹ️ **Quy định Zone B:** Khu vực này chỉ bao gồm phòng ban (BU): **C&D**. Các vị trí đã được cố định sẵn BU tương ứng.")
    
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
        with st.expander("📷 Xem ảnh sơ đồ gốc Zone C"):
            st.image(IMG_ZONE_C, caption="Sơ đồ chi tiết Zone C", use_container_width=True)
        st.markdown("---")
        
    st.info("ℹ️ **Quy định Zone C:** Khu vực này chỉ bao gồm các phòng ban (BU): **C&D, Workday.Delivery, SAP**. Các vị trí đã được cố định sẵn BU tương ứng.")
    
    rows_c_exact = [
        [f"C-01", f"C-02", f"C-03", f"C-04", f"C-05", None, None, f"C-06", f"C-07", None],
        [f"C-08", f"C-09", f"C-10", f"C-11", f"C-12", f"C-13", f"C-14", f"C-15", f"C-16", None],
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

# --- TAB QUẢN TRỊ TRỰC TUYẾN (ADMIN PANEL) ---
with tab_admin:
    st.subheader("🔐 Khu vực Quản Trị Viên (Admin)")
    admin_password = st.text_input("Nhập mật khẩu Admin để chỉnh sửa:", type="password", key="admin_pwd_input")
    
    # Mật khẩu quản trị mặc định bạn có thể đổi ở đây (VD: admin123)
    if admin_password == "admin123@123":
        st.success("🔓 Đăng nhập Admin thành công! Bạn có thể chỉnh sửa trực tiếp thông tin các ghế bên dưới.")
        
        # Hiển thị bảng dữ liệu có thể chỉnh sửa trực tiếp (Data Editor)
        edited_df = st.data_editor(
            df[['Mã Chỗ Ngồi', 'Phân Khu', 'Họ Và Tên', 'LDAP ID', 'Phòng Ban (BU)', 'Ghi Chú']],
            num_rows="fixed",
            use_container_width=True,
            key="admin_seat_editor"
        )
        
        col_save1, col_save2 = st.columns([1, 4])
        with col_save1:
            if st.button("💾 Lưu Thay Đổi", type="primary"):
                # Cập nhật lại dataframe gốc và lưu vào file Excel
                df['Họ Và Tên'] = edited_df['Họ Và Tên'].astype(str).str.strip()
                df['LDAP ID'] = edited_df['LDAP ID'].astype(str).str.strip()
                df['Phòng Ban (BU)'] = edited_df['Phòng Ban (BU)'].astype(str).str.strip()
                df['Ghi Chú'] = edited_df['Ghi Chú'].astype(str).str.strip()
                
                save_data(df)
                st.success("🎉 Đã lưu thay đổi thành công! Sơ đồ sẽ tự động cập nhật.")
                st.rerun()
        with col_save2:
            # Nút xóa nhanh (Reset) người ngồi tại một ghế bất kỳ
            seat_to_clear = st.selectbox("Chọn ghế cần xóa người ngồi (Reset về trống):", ["-- Chọn mã ghế --"] + list(df['Mã Chỗ Ngồi']))
            if seat_to_clear != "-- Chọn mã ghế --":
                if st.button(f"🗑️ Xóa thông tin (Reset) ghế {seat_to_clear}"):
                    df.loc[df['Mã Chỗ Ngồi'] == seat_to_clear, ['Họ Và Tên', 'LDAP ID', 'Ghi Chú']] = ''
                    # Giữ nguyên BU cố định của ghế đó
                    save_data(df)
                    st.success(f"Đã giải phóng ghế {seat_to_clear} thành công!")
                    st.rerun()
                    
        st.markdown("---")
        with st.expander("📥 Tải File Excel Dự Phòng"):
            with open(FILE_PATH, "rb") as f:
                st.download_button("Tải File Excel Hiện Tại", f, file_name="Danh_Sach_Cho_Ngoi_Cap_Nhat.xlsx")
                
    elif admin_password != "":
        st.error("❌ Mật khẩu Admin không chính xác!")
    else:
        st.info("Vui lòng nhập mật khẩu quản trị ở ô trên để mở bảng điều khiển.")
