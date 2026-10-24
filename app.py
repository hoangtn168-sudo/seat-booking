import streamlit as st
import pandas as pd
import os

# Cấu hình giao diện trang web rộng
st.set_page_config(page_title="Hệ Thống Đăng Ký Chỗ Ngồi Online", layout="wide", initial_sidebar_state="expanded")

FILE_PATH = "Danh_Sach_Cho_Ngoi.xlsx"

# Đường dẫn các file ảnh sơ đồ
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
            seats.append({'Mã Chỗ Ngồi': f'A-{i:02d}', 'Phân Khu Mặt Bằng': 'Zone A', 'Họ Và Tên': '', 'LDAP ID': '', 'Phòng Ban (BU)': '', 'Ghi Chú Bố Trí': ''})
        for i in range(1, 27):
            seats.append({'Mã Chỗ Ngồi': f'B-{i:02d}', 'Phân Khu Mặt Bằng': 'Zone B', 'Họ Và Tên': '', 'LDAP ID': '', 'Phòng Ban (BU)': '', 'Ghi Chú Bố Trí': ''})
        for i in range(1, 26):
            seats.append({'Mã Chỗ Ngồi': f'C-{i:02d}', 'Phân Khu Mặt Bằng': 'Zone C', 'Họ Và Tên': '', 'LDAP ID': '', 'Phòng Ban (BU)': '', 'Ghi Chú Bố Trí': ''})
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
            df[col] = df[col].replace(['nan', 'None', 'NaN', 'nat', '<NA>'], '')
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

st.title("📌 HỆ THỐNG ĐĂNG KÝ CHỖ NGỒI VĂN PHÒNG")

# --- HƯỚNG DẪN TRỰC QUAN CHO NGƯỜI DÙNG ĐIỆN THOẠI ---
st.info("💡 **Hướng dẫn:** Xem sơ đồ bên dưới $\rightarrow$ Chọn tab phân khu $\rightarrow$ Bấm nút **(Chọn)** $\rightarrow$ Điền thông tin ở khung đăng ký hiện ra.")

# --- HIỂN THỊ SƠ ĐỒ CHUNG TRỰC QUAN ---
if os.path.exists(IMG_CHUNG):
    with st.expander("🗺️ XEM SƠ ĐỒ TỔNG QUAN VĂN PHÒNG (Bấm để xem ảnh lớn)", expanded=True):
        st.image(IMG_CHUNG, caption="Sơ đồ tổng quan toàn văn phòng", use_container_width=True)

# Xác định trạng thái của ghế dựa trên Ghi Chú và Họ Tên
def get_status(row):
    note = str(row['Ghi Chú']).lower()
    name = str(row['Họ Và Tên'])
    
    if any(k in note for k in ["khóa", "hỏng", "không"]):
        return "XÁM"
    elif name != "" and name.lower() not in ["nan", "none", ""]:
        return "XANH"
    else:
        return "VÀNG"

# --- KHUNG FORM ĐĂNG KÝ NGAY TRÊN MÀN HÌNH CHÍNH ---
selected_seat = st.session_state.get('selected_seat', '')

if selected_seat:
    st.markdown("---")
    with st.container():
        st.success(f"🔥 BẠN ĐANG CHỌN VỊ TRÍ: **{selected_seat}**")
        
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
            if st.button("✅ Xác Nhận Đặt", type="primary"):
                if not user_ldap.strip() or not user_name.strip():
                    st.warning("⚠️ Vui lòng nhập đầy đủ LDAP ID và Họ Tên!")
                else:
                    clean_ldap = user_ldap.strip()
                    existing_booking = df[df['LDAP ID'] == clean_ldap]
                    
                    if not existing_booking.empty:
                        booked_seat = existing_booking.iloc[0]['Mã Chỗ Ngồi']
                        st.error(f"❌ Chặn spam! LDAP **{clean_ldap}** đã đăng ký ghế **{booked_seat}** rồi. Mỗi nhân sự chỉ được sở hữu 1 chỗ duy nhất!")
                    else:
                        df.loc[df['Mã Chỗ Ngồi'] == selected_seat, 'Họ Và Tên'] = user_name.strip()
                        df.loc[df['Mã Chỗ Ngồi'] == selected_seat, 'LDAP ID'] = clean_ldap
                        df.loc[df['Mã Chỗ Ngồi'] == selected_seat, 'Phòng Ban (BU)'] = user_bu
                        save_data(df)
                        
                        st.success(f"🎉 Đăng ký thành công ghế {selected_seat} cho {user_name} ({user_bu})!")
                        st.session_state['selected_seat'] = ''
                        st.rerun()
        with col_btn2:
            if st.button("❌ Hủy chọn lại"):
                st.session_state['selected_seat'] = ''
                st.rerun()
    st.markdown("---")

st.subheader("🪑 DANH SÁCH GHẾ THEO PHÂN KHU")

# Các Tab phân khu
tab1, tab2, tab3 = st.tabs(["Zone A (28 Bàn)", "Zone B (26 Bàn)", "Zone C (25 Bàn)"])

def render_zone(zone_prefix, img_path, cols_per_row=4):
    if os.path.exists(img_path):
        with st.expander(f"📷 Xem sơ đồ chi tiết {zone_prefix}"):
            st.image(img_path, caption=f"Sơ đồ chi tiết {zone_prefix}", use_container_width=True)
        st.markdown("---")
        
    zone_df = df[df['Phân Khu'].astype(str).str.startswith(zone_prefix)].reset_index(drop=True)
    
    if len(zone_df) == 0:
        st.warning(f"Không tìm thấy dữ liệu cho {zone_prefix}")
        return

    for i in range(0, len(zone_df), cols_per_row):
        cols = st.columns(cols_per_row)
        for j in range(cols_per_row):
            if i + j < len(zone_df):
                row = zone_df.iloc[i + j]
                code = row['Mã Chỗ Ngồi']
                name = row['Họ Và Tên'] if pd.notna(row['Họ Và Tên']) and str(row['Họ Và Tên']).strip().lower() not in ["nan", "none", ""] else "[Trống]"
                status = get_status(row)
                
                with cols[j]:
                    if status == "XANH":
                        st.success(f"**{code}**\n\n👤 {name}")
                    elif status == "VÀNG":
                        if st.button(f"🟡 {code}\n(Chọn)", key=f"btn_{code}"):
                            st.session_state['selected_seat'] = code
                            st.rerun()
                    else:
                        st.error(f"**{code}**\n\n🔒 Khóa")

with tab1:
    render_zone("Zone A", IMG_ZONE_A, 4)

with tab2:
    render_zone("Zone B", IMG_ZONE_B, 4)

with tab3:
    render_zone("Zone C", IMG_ZONE_C, 4)

# --- TẢI FILE EXCEL (Dành cho Quản trị viên) ---
st.markdown("---")
with st.expander("🛠️ Dành cho Quản trị viên (Admin)"):
    with open(FILE_PATH, "rb") as f:
        st.download_button("📥 Tải File Excel Danh Sách Chỗ Ngồi", f, file_name="Danh_Sach_Cho_Ngoi_Cap_Nhat.xlsx")
