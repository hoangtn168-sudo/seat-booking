import streamlit as st
import pandas as pd
import os

# Cấu hình giao diện trang web rộng
st.set_page_config(page_title="Hệ Thống Đăng Ký Chỗ Ngồi Online", layout="wide")

FILE_PATH = "Danh_Sach_Cho_Ngoi.xlsx"

# Danh sách các phòng ban (BU) cho phép người dùng tự chọn
LIST_BU = ['IOT', 'C&D', 'WorkDay', 'SAP', 'HelpDesk']

# Hàm đọc và chuẩn hóa dữ liệu từ file Excel
def load_data():
    if not os.path.exists(FILE_PATH):
        # Tạo file mẫu nếu chưa tồn tại
        seats = []
        for i in range(1, 29):
            seats.append({'Mã Chỗ Ngồi': f'A-{i:02d}', 'Phân Khu Mặt Bằng': 'Zone A', 'Họ Và Tên': '', 'LDAP ID': '', 'Phòng Ban (BU)': '', 'Ghi Chú Bố Trí': ''})
        for i in range(1, 27):
            seats.append({'Mã Chỗ Ngồi': f'B-{i:02d}', 'Phân Khu Mặt Bằng': 'Zone B', 'Họ Và Tên': '', 'LDAP ID': '', 'Phòng Ban (BU)': '', 'Ghi Chú Bố Trí': ''})
        for i in range(1, 26):
            seats.append({'Mã Chỗ Ngồi': f'C-{i:02d}', 'Phân Khu Mặt Bằng': 'Zone C', 'Họ Và Tên': '', 'LDAP ID': '', 'Phòng Ban (BU)': '', 'Ghi Chú Bố Trí': ''})
        df_init = pd.DataFrame(seats)
        df_init.to_excel(FILE_PATH, index=False)

    df = pd.read_excel(FILE_PATH)
    
    # Chuẩn hóa tên cột
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
    
    # Ép kiểu dữ liệu sang string an toàn tránh lỗi float64
    for col in ['Họ Và Tên', 'LDAP ID', 'Phòng Ban (BU)', 'Ghi Chú']:
        if col in df.columns:
            df[col] = df[col].astype(str).replace(['nan', 'None', 'NaN', 'nat'], '').str.strip()
        else:
            df[col] = ''

    # Tự động nhận diện Phân Khu nếu chưa có
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

st.title("📌 HỆ THỐNG ĐĂNG KÝ CHỖ NGỒI VĂN PHÒNG ONLINE")
st.markdown("Chọn tab phân khu bên dưới, bấm vào ghế trống màu vàng để tiến hành đăng ký.")

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

# Các Tab phân khu
tab1, tab2, tab3 = st.tabs(["Zone A (28 Bàn)", "Zone B (26 Bàn)", "Zone C (25 Bàn)"])

def render_zone(zone_prefix, cols_per_row=7):
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
                        if st.button(f"🟡 {code}\n(Chọn chỗ)", key=f"btn_{code}"):
                            st.session_state['selected_seat'] = code
                    else:
                        st.error(f"**{code}**\n\n🔒 Khóa/Hỏng")

with tab1:
    render_zone("Zone A", 7)

with tab2:
    render_zone("Zone B", 8)

with tab3:
    render_zone("Zone C", 7)

# --- Sidebar xử lý đặt chỗ ---
st.sidebar.header("📋 FORM ĐĂNG KÝ CHỖ NGỒI")
selected_seat = st.session_state.get('selected_seat', '')

if selected_seat:
    st.sidebar.info(f"Đang chọn vị trí: **{selected_seat}**")
    
    user_ldap = st.sidebar.text_input("Nhập LDAP ID (Mã nhân viên):", key="input_ldap")
    user_name = st.sidebar.text_input("Nhập Họ Và Tên:", key="input_name")
    user_bu = st.sidebar.selectbox("Chọn Phòng Ban (BU):", LIST_BU, key="input_bu")
    
    if st.sidebar.button("Xác Nhận Đặt Chỗ"):
        if not user_ldap.strip() or not user_name.strip():
            st.sidebar.warning("⚠️ Vui lòng nhập đầy đủ LDAP ID và Họ Tên!")
        else:
            # Kiểm tra trùng lặp LDAP ID trên toàn bộ hệ thống file hiện tại
            clean_ldap = user_ldap.strip()
            existing_booking = df[df['LDAP ID'].str.lower() == clean_ldap.lower()]
            
            if not existing_booking.empty:
                booked_seat = existing_booking.iloc[0]['Mã Chỗ Ngồi']
                st.sidebar.error(f"❌ Chặn spam! LDAP **{clean_ldap}** đã đăng ký ghế **{booked_seat}** rồi. Mỗi nhân sự chỉ được sở hữu 1 chỗ duy nhất!")
            else:
                # Ghi nhận thông tin vào ghế được chọn
                df.loc[df['Mã Chỗ Ngồi'] == selected_seat, 'Họ Và Tên'] = user_name.strip()
                df.loc[df['Mã Chỗ Ngồi'] == selected_seat, 'LDAP ID'] = clean_ldap
                df.loc[df['Mã Chỗ Ngồi'] == selected_seat, 'Phòng Ban (BU)'] = user_bu
                save_data(df)
                
                st.sidebar.success(f"✅ Đăng ký thành công ghế {selected_seat} cho {user_name} ({user_bu})!")
                st.session_state['selected_seat'] = ''
                st.rerun()
else:
    st.sidebar.write("👉 Bấm vào nút **(Chọn chỗ)** ở các ô màu vàng trên sơ đồ để tiến hành đăng ký.")

st.sidebar.markdown("---")
with open(FILE_PATH, "rb") as f:
    st.sidebar.download_button("📥 Tải File Excel Cập Nhật", f, file_name="Danh_Sach_Cho_Ngoi_Cap_Nhat.xlsx")