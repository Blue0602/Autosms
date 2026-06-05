# ============================================================
# HỆ THỐNG XỬ LÝ LEAD ZALO TỰ ĐỘNG - GIAI ĐOẠN 1: CHỈ TEXT
# Tác dụng:
# - Upload file Excel khách hàng
# - Lọc và chuẩn hóa SĐT
# - Hiển thị từng khách hàng dạng flashcard
# - Tự sinh nội dung nhắn Zalo bằng f-string
# - Chừa sẵn vị trí nâng cấp tạo ảnh Voucher bằng Pillow
# ============================================================

import re
import pandas as pd
import streamlit as st


# ============================================================
# 1. CẤU HÌNH GIAO DIỆN STREAMLIT
# ============================================================

st.set_page_config(
    page_title="Xử lý Lead Zalo VNPT",
    page_icon="📲",
    layout="wide"
)


# ============================================================
# 2. CSS NHẸ CHO GIAO DIỆN ĐẸP HƠN
# ============================================================

st.markdown(
    """
    <style>
        .main-title {
            font-size: 32px;
            font-weight: 800;
            color: #005baa;
            margin-bottom: 6px;
        }

        .sub-title {
            font-size: 16px;
            color: #555;
            margin-bottom: 24px;
        }

        .customer-card {
            padding: 22px;
            border-radius: 18px;
            background: linear-gradient(135deg, #f5fbff, #ffffff);
            border: 1px solid #d7ecff;
            box-shadow: 0px 4px 18px rgba(0, 91, 170, 0.08);
            margin-bottom: 18px;
        }

        .customer-name {
            font-size: 24px;
            font-weight: 800;
            color: #003b73;
            margin-bottom: 10px;
        }

        .info-line {
            font-size: 16px;
            margin-bottom: 6px;
            color: #333;
        }

        .success-box {
            padding: 12px 16px;
            border-radius: 12px;
            background-color: #e9f8ef;
            border: 1px solid #b7ebc6;
            color: #176b36;
            font-weight: 600;
        }

        .warning-box {
            padding: 12px 16px;
            border-radius: 12px;
            background-color: #fff8e6;
            border: 1px solid #ffe0a3;
            color: #8a5a00;
            font-weight: 600;
        }
    </style>
    """,
    unsafe_allow_html=True
)


# ============================================================
# 3. HÀM CHUẨN HÓA SỐ ĐIỆN THOẠI
# ============================================================

def chuan_hoa_sdt(value):
    """
    Chuẩn hóa số điện thoại từ Excel.

    Mục tiêu:
    - Chuyển số điện thoại về chuỗi.
    - Bỏ phần .0 nếu Excel đọc thành dạng float.
    - Chỉ giữ lại chữ số.
    - Nếu thiếu số 0 ở đầu thì tự động thêm 0.
    """

    if pd.isna(value):
        return ""

    sdt = str(value).strip()

    # Trường hợp Excel đọc số thành 933599234.0
    if sdt.endswith(".0"):
        sdt = sdt[:-2]

    # Chỉ giữ lại chữ số, bỏ khoảng trắng, dấu chấm, dấu gạch...
    sdt = re.sub(r"\D", "", sdt)

    # Nếu số bắt đầu bằng mã quốc gia 84 thì đổi về 0
    # Ví dụ: 84933599234 -> 0933599234
    if sdt.startswith("84") and len(sdt) >= 10:
        sdt = "0" + sdt[2:]

    # Nếu thiếu số 0 đầu thì thêm vào
    # Ví dụ: 933599234 -> 0933599234
    if sdt and not sdt.startswith("0"):
        sdt = "0" + sdt

    return sdt


# ============================================================
# 4. HÀM TRÍCH XUẤT KHU VỰC TỪ ĐỊA CHỈ
# ============================================================

def trich_xuat_khu_vuc(dia_chi):
    """
    Trích xuất khu vực từ địa chỉ.

    Quy tắc:
    - Lấy phần chữ nằm giữa dấu phẩy đầu tiên và dấu phẩy thứ hai.
    - Nếu lỗi hoặc không đủ dấu phẩy thì mặc định là "Long Thành".

    Ví dụ:
    "Ấp 1, Xã Phước Thái, Huyện Long Thành, Đồng Nai"
    => "Xã Phước Thái"
    """

    try:
        if pd.isna(dia_chi):
            return "Long Thành"

        parts = str(dia_chi).split(",")

        if len(parts) >= 3:
            khu_vuc = parts[1].strip()
            return khu_vuc if khu_vuc else "Long Thành"

        return "Long Thành"

    except Exception:
        return "Long Thành"


# ============================================================
# 5. HÀM ĐỌC VÀ LÀM SẠCH FILE EXCEL
# ============================================================

def doc_va_lam_sach_excel(uploaded_file):
    """
    Đọc file Excel và trả về DataFrame đã làm sạch.

    Các cột bắt buộc:
    - Tên KH (*)
    - Địa chỉ (*)
    - Điện thoại (*)
    """

    required_columns = ["Tên KH (*)", "Địa chỉ (*)", "Điện thoại (*)"]

    df = pd.read_excel(uploaded_file)

    # Kiểm tra thiếu cột
    missing_columns = [col for col in required_columns if col not in df.columns]

    if missing_columns:
        raise ValueError(
            "File Excel đang thiếu các cột bắt buộc: "
            + ", ".join(missing_columns)
        )

    # Chỉ lấy đúng các cột cần dùng
    df = df[required_columns].copy()

    # Bỏ dòng trống số điện thoại
    df = df.dropna(subset=["Điện thoại (*)"])

    # Chuẩn hóa dữ liệu
    df["Tên KH (*)"] = df["Tên KH (*)"].fillna("Khách hàng").astype(str).str.strip()
    df["Địa chỉ (*)"] = df["Địa chỉ (*)"].fillna("").astype(str).str.strip()
    df["Điện thoại (*)"] = df["Điện thoại (*)"].apply(chuan_hoa_sdt)

    # Bỏ những dòng sau khi chuẩn hóa mà SĐT vẫn rỗng
    df = df[df["Điện thoại (*)"] != ""]

    # Reset index để chạy flashcard từ 0, 1, 2...
    df = df.reset_index(drop=True)

    return df


# ============================================================
# 6. HÀM TẠO TIN NHẮN ZALO BẰNG F-STRING
# ============================================================

def tao_tin_nhan_zalo(ten_kh, khu_vuc):
    """
    Tạo nội dung tin nhắn Zalo tự động.

    BẮT BUỘC dùng f-string:
    - {ten_kh} thay đổi theo từng khách hàng.
    - {khu_vuc} thay đổi theo địa chỉ từng khách hàng.
    """

    tin_nhan = f"""🏢 [VNPT ĐỒNG NAI] HỖ TRỢ CHUYỂN ĐỔI SỐ HỘ KINH DOANH

Dạ em chào anh/chị chủ cơ sở {ten_kh},
Em là Thuận – Chuyên viên VNPT phụ trách trực tiếp hỗ trợ các Hộ kinh doanh tại địa bàn {khu_vuc}.

Nhằm giúp cơ sở mình tuân thủ kịp thời quy định của Cơ quan Thuế, em gửi anh/chị hệ sinh thái Dịch vụ Số niêm yết của VNPT:
🔹 Hóa đơn máy tính tiền (VNPT Invoice)
🔹 Chữ ký số (VNPT SmartCA/CA)
🔹 Phần mềm Kế toán (SME Accounting)
🔹 Internet Cáp quang (FiberVNN)

🎁 Đặc biệt, VNPT đang có chính sách ưu đãi cước phí riêng cho các cơ sở đăng ký Combo tại khu vực nhà mình. Anh/chị xem qua, nếu cần hỗ trợ cứ nhắn hoặc nháy máy, em Thuận sẽ gọi lại ngay ạ!"""

    return tin_nhan


# ============================================================
# 7. 🌟 [TƯƠNG LAI - TÍNH NĂNG TẠO ẢNH VOUCHER SẼ ĐẶT Ở ĐÂY]
# ============================================================

def tao_anh_voucher_tuong_lai(ten_kh, khu_vuc, sdt):
    """
    🌟 HÀM DỰ PHÒNG CHO GIAI ĐOẠN 2

    Sau này khi bạn muốn sinh ảnh Voucher bằng Pillow,
    hãy viết code tạo ảnh trong hàm này.

    Gợi ý nâng cấp sau:
    - from PIL import Image, ImageDraw, ImageFont
    - Mở template voucher PNG/JPG
    - Ghi tên khách hàng lên ảnh
    - Ghi khu vực
    - Ghi mã ưu đãi
    - Xuất ảnh thành file PNG để tải xuống

    Hiện tại Giai đoạn 1 chỉ xử lý TEXT nên hàm này chưa chạy.
    """

    # 🌟 [TƯƠNG LAI - IMPORT PILLOW Ở ĐÂY]
    # from PIL import Image, ImageDraw, ImageFont

    # 🌟 [TƯƠNG LAI - MỞ FILE TEMPLATE VOUCHER Ở ĐÂY]
    # image = Image.open("voucher_template.png")

    # 🌟 [TƯƠNG LAI - VẼ TÊN KHÁCH HÀNG / KHU VỰC / SĐT LÊN ẢNH Ở ĐÂY]
    # draw = ImageDraw.Draw(image)

    # 🌟 [TƯƠNG LAI - LƯU ẢNH VOUCHER Ở ĐÂY]
    # image.save("voucher_output.png")

    pass


# ============================================================
# 8. KHỞI TẠO SESSION STATE
# ============================================================

if "current_index" not in st.session_state:
    st.session_state.current_index = 0

if "uploaded_file_name" not in st.session_state:
    st.session_state.uploaded_file_name = None


# ============================================================
# 9. HÀM ĐIỀU HƯỚNG FLASHCARD
# ============================================================

def di_lui():
    """Chuyển về khách hàng trước đó."""
    if st.session_state.current_index > 0:
        st.session_state.current_index -= 1


def di_toi(tong_so_khach):
    """Chuyển sang khách hàng tiếp theo."""
    if st.session_state.current_index < tong_so_khach - 1:
        st.session_state.current_index += 1


# ============================================================
# 10. GIAO DIỆN CHÍNH
# ============================================================

st.markdown('<div class="main-title">📲 Hệ thống Xử lý Lead Zalo Tự động</div>', unsafe_allow_html=True)
st.markdown(
    '<div class="sub-title">Giai đoạn 1: Upload Excel → Chuẩn hóa SĐT → Sinh tin nhắn Zalo tự động cho từng khách hàng.</div>',
    unsafe_allow_html=True
)

uploaded_file = st.file_uploader(
    "📁 Tải file Excel danh sách khách hàng",
    type=["xls", "xlsx"],
    help="File cần có 3 cột: Tên KH (*), Địa chỉ (*), Điện thoại (*)"
)

if uploaded_file is None:
    st.info("Bạn hãy tải file Excel lên để bắt đầu xử lý lead Zalo.")
    st.stop()


# Reset current_index nếu người dùng upload file mới
if st.session_state.uploaded_file_name != uploaded_file.name:
    st.session_state.uploaded_file_name = uploaded_file.name
    st.session_state.current_index = 0


# ============================================================
# 11. ĐỌC FILE VÀ XỬ LÝ LỖI
# ============================================================

try:
    df = doc_va_lam_sach_excel(uploaded_file)

except Exception as e:
    st.error("Không thể đọc hoặc xử lý file Excel.")
    st.warning(str(e))
    st.stop()


if df.empty:
    st.error("File Excel không có khách hàng hợp lệ sau khi lọc số điện thoại.")
    st.stop()


tong_so_khach = len(df)

# Đảm bảo current_index không vượt quá số dòng hiện có
if st.session_state.current_index >= tong_so_khach:
    st.session_state.current_index = tong_so_khach - 1

if st.session_state.current_index < 0:
    st.session_state.current_index = 0


# ============================================================
# 12. LẤY KHÁCH HÀNG HIỆN TẠI
# ============================================================

current_index = st.session_state.current_index
khach_hang = df.iloc[current_index]

ten_kh = khach_hang["Tên KH (*)"]
dia_chi = khach_hang["Địa chỉ (*)"]
sdt = khach_hang["Điện thoại (*)"]
khu_vuc = trich_xuat_khu_vuc(dia_chi)

# Tạo tin nhắn bằng f-string
tin_nhan = tao_tin_nhan_zalo(ten_kh=ten_kh, khu_vuc=khu_vuc)


# ============================================================
# 13. HIỂN THỊ THẺ THÔNG TIN KHÁCH HÀNG
# ============================================================

st.markdown(
    f"""
    <div class="customer-card">
        <div class="customer-name">👤 Khách hàng {current_index + 1} / {tong_so_khach}: {ten_kh}</div>
        <div class="info-line"><b>📍 Khu vực tự trích xuất:</b> {khu_vuc}</div>
        <div class="info-line"><b>🏠 Địa chỉ gốc:</b> {dia_chi}</div>
        <div class="info-line"><b>☎️ Số điện thoại đã chuẩn hóa:</b> {sdt}</div>
    </div>
    """,
    unsafe_allow_html=True
)

st.progress((current_index + 1) / tong_so_khach)

st.markdown(
    f"""
    <div class="success-box">
        Đã xử lý {tong_so_khach} khách hàng hợp lệ. 
        Đang xem khách hàng số {current_index + 1}.
    </div>
    """,
    unsafe_allow_html=True
)

st.write("")


# ============================================================
# 14. KHU VỰC COPY 1 CHẠM
# ============================================================

col1, col2 = st.columns([1, 2])

with col1:
    st.subheader("📞 Số điện thoại")
    st.code(sdt, language="text")

with col2:
    st.subheader("💬 Nội dung tin nhắn Zalo")
    st.code(tin_nhan, language="text")


# ============================================================
# 15. 🌟 [TƯƠNG LAI - KHU VỰC HIỂN THỊ ẢNH VOUCHER SẼ ĐẶT Ở ĐÂY]
# ============================================================

st.markdown("---")
st.markdown("### 🎁 Khu vực Voucher")
st.caption("Giai đoạn 1 chưa sinh ảnh. Khu vực này đã được chừa sẵn để nâng cấp bằng Pillow sau này.")

# 🌟 [TƯƠNG LAI - GỌI HÀM TẠO ẢNH VOUCHER Ở ĐÂY]
# voucher_path = tao_anh_voucher_tuong_lai(ten_kh=ten_kh, khu_vuc=khu_vuc, sdt=sdt)

# 🌟 [TƯƠNG LAI - HIỂN THỊ ẢNH VOUCHER Ở ĐÂY]
# st.image(voucher_path, caption="Voucher ưu đãi dành cho khách hàng")

# 🌟 [TƯƠNG LAI - NÚT TẢI ẢNH VOUCHER Ở ĐÂY]
# with open(voucher_path, "rb") as file:
#     st.download_button(
#         label="⬇️ Tải ảnh Voucher",
#         data=file,
#         file_name=f"voucher_{sdt}.png",
#         mime="image/png"
#     )


# ============================================================
# 16. NÚT ĐIỀU HƯỚNG
# ============================================================

st.markdown("---")

nav_col1, nav_col2, nav_col3 = st.columns([1, 1, 1])

with nav_col1:
    st.button(
        "⬅️ Quay lại",
        on_click=di_lui,
        disabled=current_index == 0,
        use_container_width=True
    )

with nav_col2:
    st.markdown(
        f"""
        <div class="warning-box" style="text-align:center;">
            {current_index + 1} / {tong_so_khach}
        </div>
        """,
        unsafe_allow_html=True
    )

with nav_col3:
    st.button(
        "Bỏ qua / Đã nhắn Zalo ➡️",
        on_click=di_toi,
        args=(tong_so_khach,),
        disabled=current_index == tong_so_khach - 1,
        use_container_width=True
    )


# ============================================================
# 17. XEM NHANH DỮ LIỆU ĐÃ LÀM SẠCH
# ============================================================

with st.expander("📊 Xem bảng dữ liệu đã làm sạch"):
    st.dataframe(df, use_container_width=True)
