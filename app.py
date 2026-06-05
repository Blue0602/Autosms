# ============================================================
# HỆ THỐNG QUẢN LÝ VÀ XỬ LÝ LEAD ZALO B2A CHO VNPT
# Giai đoạn 1: Xử lý Text + Quản lý trạng thái + Xuất báo cáo Excel
# ============================================================

import re
from io import BytesIO

import pandas as pd
import streamlit as st


# ============================================================
# 1. CẤU HÌNH TRANG
# ============================================================

st.set_page_config(
    page_title="Lead Zalo B2A VNPT",
    page_icon="📲",
    layout="wide"
)


# ============================================================
# 2. CSS GIAO DIỆN
# ============================================================

st.markdown(
    """
    <style>
        .main-title {
            font-size: 32px;
            font-weight: 800;
            color: #005baa;
            margin-bottom: 4px;
        }

        .sub-title {
            font-size: 16px;
            color: #555;
            margin-bottom: 24px;
        }

        .customer-card {
            padding: 24px;
            border-radius: 18px;
            background: linear-gradient(135deg, #f3f9ff, #ffffff);
            border: 1px solid #d8ecff;
            box-shadow: 0px 4px 18px rgba(0, 91, 170, 0.08);
            margin-bottom: 18px;
        }

        .customer-name {
            font-size: 24px;
            font-weight: 800;
            color: #003b73;
            margin-bottom: 12px;
        }

        .info-line {
            font-size: 16px;
            color: #333;
            margin-bottom: 8px;
        }

        .status-done {
            display: inline-block;
            padding: 6px 12px;
            border-radius: 999px;
            background-color: #e9f8ef;
            color: #176b36;
            border: 1px solid #b7ebc6;
            font-weight: 700;
        }

        .status-pending {
            display: inline-block;
            padding: 6px 12px;
            border-radius: 999px;
            background-color: #fff8e6;
            color: #8a5a00;
            border: 1px solid #ffe0a3;
            font-weight: 700;
        }

        .metric-box {
            padding: 16px;
            border-radius: 14px;
            background-color: #f7fbff;
            border: 1px solid #dceeff;
            text-align: center;
        }

        .metric-number {
            font-size: 26px;
            font-weight: 800;
            color: #005baa;
        }

        .metric-label {
            font-size: 14px;
            color: #555;
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
    Chuẩn hóa số điện thoại.

    Yêu cầu:
    - Không xóa dòng nếu trống SĐT.
    - Nếu trống thì trả về "Chưa có SĐT".
    - Chuyển về chuỗi.
    - Bỏ lỗi .0 do Excel đọc dạng số.
    - Chỉ giữ lại chữ số.
    - Nếu thiếu số 0 đầu thì tự thêm.
    """

    if pd.isna(value):
        return "Chưa có SĐT"

    sdt = str(value).strip()

    if sdt == "" or sdt.lower() in ["nan", "none", "null"]:
        return "Chưa có SĐT"

    # Trường hợp Excel đọc thành 933599234.0
    if sdt.endswith(".0"):
        sdt = sdt[:-2]

    # Chỉ giữ lại chữ số
    sdt = re.sub(r"\D", "", sdt)

    if sdt == "":
        return "Chưa có SĐT"

    # Nếu bắt đầu bằng 84 thì chuyển về 0
    # Ví dụ: 84933599234 -> 0933599234
    if sdt.startswith("84") and len(sdt) >= 10:
        sdt = "0" + sdt[2:]

    # Nếu thiếu số 0 ở đầu thì thêm vào
    if not sdt.startswith("0"):
        sdt = "0" + sdt

    return sdt


# ============================================================
# 4. HÀM TRÍCH XUẤT KHU VỰC TỪ ĐỊA CHỈ
# ============================================================

def trich_xuat_khu_vuc(dia_chi):
    """
    Trích xuất khu vực từ địa chỉ.

    Quy tắc:
    - Lấy phần nằm giữa dấu phẩy thứ nhất và dấu phẩy thứ hai.
    - Nếu lỗi thì mặc định là "Long Thành".

    Ví dụ:
    "Ấp 1, Xã Phước Thái, Huyện Long Thành, Đồng Nai"
    => "Xã Phước Thái"
    """

    try:
        if pd.isna(dia_chi):
            return "Long Thành"

        dia_chi = str(dia_chi).strip()

        if dia_chi == "":
            return "Long Thành"

        parts = dia_chi.split(",")

        if len(parts) >= 3:
            khu_vuc = parts[1].strip()
            return khu_vuc if khu_vuc else "Long Thành"

        return "Long Thành"

    except Exception:
        return "Long Thành"


# ============================================================
# 5. HÀM ĐỌC VÀ XỬ LÝ FILE EXCEL
# ============================================================

def doc_file_excel(file_input):
    """
    Đọc file Excel và xử lý dữ liệu.

    Cột yêu cầu:
    - Tên KH (*)
    - Địa chỉ (*)
    - Điện thoại (*)

    Lưu ý quan trọng:
    - Không xóa bất kỳ dòng nào.
    - Dòng trống SĐT sẽ được giữ lại và ghi là "Chưa có SĐT".
    - Nếu chưa có cột "Trạng thái", hệ thống tự tạo.
    """

    required_columns = ["Tên KH (*)", "Địa chỉ (*)", "Điện thoại (*)"]

    df = pd.read_excel(file_input)

    # Kiểm tra cột bắt buộc
    missing_columns = [col for col in required_columns if col not in df.columns]

    if missing_columns:
        raise ValueError(
            "File Excel đang thiếu cột bắt buộc: "
            + ", ".join(missing_columns)
        )

    # Đảm bảo các cột quan trọng không bị NaN
    df["Tên KH (*)"] = df["Tên KH (*)"].fillna("Khách hàng chưa có tên").astype(str).str.strip()
    df["Địa chỉ (*)"] = df["Địa chỉ (*)"].fillna("").astype(str).str.strip()

    # Chuẩn hóa SĐT nhưng KHÔNG xóa dòng
    df["Điện thoại (*)"] = df["Điện thoại (*)"].apply(chuan_hoa_sdt)

    # Tạo cột trạng thái nếu chưa có
    if "Trạng thái" not in df.columns:
        df["Trạng thái"] = "Chưa gửi"

    # Nếu trạng thái bị trống thì điền mặc định
    df["Trạng thái"] = df["Trạng thái"].fillna("Chưa gửi").astype(str).str.strip()
    df.loc[df["Trạng thái"] == "", "Trạng thái"] = "Chưa gửi"

    # Tạo cột khu vực để tiện báo cáo
    df["Khu vực tự trích xuất"] = df["Địa chỉ (*)"].apply(trich_xuat_khu_vuc)

    # Giữ nguyên số dòng, chỉ reset lại index cho dễ chạy flashcard
    df = df.reset_index(drop=True)

    return df


# ============================================================
# 6. HÀM TẠO TIN NHẮN ZALO BẰNG F-STRING
# ============================================================

def tao_tin_nhan_zalo(ten_kh, khu_vuc):
    """
    Tạo nội dung tin nhắn Zalo tự động.

    Bắt buộc dùng f-string để:
    - {ten_kh} thay đổi theo từng khách hàng.
    - {khu_vuc} thay đổi theo từng địa chỉ.
    """

    tin_nhan = f"""🏢 [VNPT ĐỒNG NAI] HỖ TRỢ CHUYỂN ĐỔI SỐ HỘ KINH DOANH

Dạ em chào anh/chị chủ cơ sở {ten_kh},
Em là Thuận – Chuyên viên VNPT phụ trách trực tiếp hỗ trợ các Hộ kinh doanh tại địa bàn {khu_vuc}.

Nhằm giúp cơ sở mình tuân thủ kịp thời quy định của Cơ quan Thuế, em gửi anh/chị hệ sinh thái Dịch vụ Số niêm yết của VNPT:
🔹 Hóa đơn máy tính tiền (VNPT Invoice)
🔹 Chữ ký số (VNPT SmartCA/CA)
🔹 Phần mềm Kế toán (SME Accounting)
🔹 Internet Cáp quang (FiberVNN)

🎁 Đặc biệt, VNPT đang có chính sách ưu đãi cước phí riêng cho các cơ sở đăng ký Combo tại khu vực nhà mình. 
Số điện thoại hỗ trợ: 0837892579. Anh/chị xem qua, nếu cần tư vấn cứ nhắn hoặc nháy máy, em Thuận sẽ gọi lại ngay ạ! Trân trọng!"""

    return tin_nhan


# ============================================================
# 7. HÀM GHI NHẬN TRẠNG THÁI
# ============================================================

def ghi_nhan_da_gui_zalo():
    """
    Khi bấm nút [Đã nhắn Zalo ➡️]:
    - Cập nhật cột "Trạng thái" của khách hiện tại thành "Đã gửi Zalo".
    - Sau đó tự chuyển sang khách hàng tiếp theo nếu còn.
    """

    current_index = st.session_state.current_index

    if "df_leads" in st.session_state:
        st.session_state.df_leads.loc[current_index, "Trạng thái"] = "Đã gửi Zalo"

        tong_so_khach = len(st.session_state.df_leads)

        if current_index < tong_so_khach - 1:
            st.session_state.current_index += 1


def quay_lai_khach_truoc():
    """
    Chuyển về khách hàng trước đó.
    """

    if st.session_state.current_index > 0:
        st.session_state.current_index -= 1


# ============================================================
# 8. HÀM XUẤT EXCEL
# ============================================================

def tao_file_excel_download(df):
    """
    Tạo file Excel trong bộ nhớ để Streamlit tải xuống.
    """

    output = BytesIO()

    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name="Bao_cao_Lead_Zalo")

    output.seek(0)

    return output


# ============================================================
# 9. 🌟 [TƯƠNG LAI - TÍNH NĂNG TẠO ẢNH VOUCHER SẼ ĐẶT Ở ĐÂY]
# ============================================================

def tao_anh_voucher_tuong_lai(ten_kh, khu_vuc, sdt):
    """
    🌟 HÀM DỰ PHÒNG CHO GIAI ĐOẠN 2

    Sau này khi cần tự động sinh ảnh Voucher bằng Pillow,
    hãy viết code vào hàm này.

    Gợi ý nâng cấp:
    - from PIL import Image, ImageDraw, ImageFont
    - Mở ảnh nền Voucher.
    - Ghi tên khách hàng.
    - Ghi khu vực.
    - Ghi số điện thoại.
    - Sinh mã ưu đãi riêng.
    - Xuất file PNG/JPG.
    """

    # 🌟 [TƯƠNG LAI - IMPORT PILLOW Ở ĐÂY]
    # from PIL import Image, ImageDraw, ImageFont

    # 🌟 [TƯƠNG LAI - MỞ TEMPLATE ẢNH VOUCHER Ở ĐÂY]
    # image = Image.open("voucher_template.png")

    # 🌟 [TƯƠNG LAI - VẼ TEXT LÊN ẢNH Ở ĐÂY]
    # draw = ImageDraw.Draw(image)

    # 🌟 [TƯƠNG LAI - LƯU ẢNH VOUCHER Ở ĐÂY]
    # image.save("voucher_output.png")

    pass


# ============================================================
# 10. KHỞI TẠO SESSION STATE
# ============================================================

if "current_index" not in st.session_state:
    st.session_state.current_index = 0

if "df_leads" not in st.session_state:
    st.session_state.df_leads = None

if "file_name" not in st.session_state:
    st.session_state.file_name = None


# ============================================================
# 11. GIAO DIỆN CHÍNH
# ============================================================

st.markdown(
    '<div class="main-title">📲 Hệ thống Quản lý và Xử lý Lead Zalo B2A cho VNPT</div>',
    unsafe_allow_html=True
)

st.markdown(
    '<div class="sub-title">Quản lý danh sách Hộ Kinh Doanh, nhắn Zalo nhanh, cập nhật trạng thái và xuất báo cáo cuối ngày.</div>',
    unsafe_allow_html=True
)


# ============================================================
# 12. KHU VỰC TẢI FILE
# ============================================================

st.sidebar.header("📁 Dữ liệu đầu vào")

uploaded_file = st.sidebar.file_uploader(
    "Tải file Excel",
    type=["xls", "xlsx"],
    help="File mẫu: report-1780634516013.xls"
)

st.sidebar.caption("File cần có các cột: Tên KH (*), Địa chỉ (*), Điện thoại (*)")


# ============================================================
# 13. ĐỌC FILE TỪ UPLOAD HOẶC FILE CÓ SẴN
# ============================================================

if uploaded_file is not None:
    try:
        # Nếu upload file mới thì đọc lại và reset về khách đầu tiên
        if st.session_state.file_name != uploaded_file.name:
            st.session_state.df_leads = doc_file_excel(uploaded_file)
            st.session_state.current_index = 0
            st.session_state.file_name = uploaded_file.name

    except Exception as e:
        st.error("Không thể đọc file Excel.")
        st.warning(str(e))
        st.stop()

else:
    st.info("Bạn hãy tải file `report-1780634516013.xls` lên để bắt đầu.")
    st.stop()


df = st.session_state.df_leads

if df is None or df.empty:
    st.error("Không có dữ liệu khách hàng để xử lý.")
    st.stop()


# ============================================================
# 14. ĐẢM BẢO INDEX KHÔNG BỊ VƯỢT GIỚI HẠN
# ============================================================

tong_so_khach = len(df)

if st.session_state.current_index < 0:
    st.session_state.current_index = 0

if st.session_state.current_index >= tong_so_khach:
    st.session_state.current_index = tong_so_khach - 1


current_index = st.session_state.current_index
khach_hang = df.iloc[current_index]


# ============================================================
# 15. LẤY THÔNG TIN KHÁCH HÀNG HIỆN TẠI
# ============================================================

ten_kh = khach_hang["Tên KH (*)"]
dia_chi = khach_hang["Địa chỉ (*)"]
sdt = khach_hang["Điện thoại (*)"]
khu_vuc = khach_hang["Khu vực tự trích xuất"]
trang_thai = khach_hang["Trạng thái"]

tin_nhan = tao_tin_nhan_zalo(
    ten_kh=ten_kh,
    khu_vuc=khu_vuc
)


# ============================================================
# 16. THỐNG KÊ NHANH
# ============================================================

da_gui = int((df["Trạng thái"] == "Đã gửi Zalo").sum())
chua_gui = tong_so_khach - da_gui

metric_col1, metric_col2, metric_col3 = st.columns(3)

with metric_col1:
    st.markdown(
        f"""
        <div class="metric-box">
            <div class="metric-number">{tong_so_khach}</div>
            <div class="metric-label">Tổng số lead</div>
        </div>
        """,
        unsafe_allow_html=True
    )

with metric_col2:
    st.markdown(
        f"""
        <div class="metric-box">
            <div class="metric-number">{da_gui}</div>
            <div class="metric-label">Đã gửi Zalo</div>
        </div>
        """,
        unsafe_allow_html=True
    )

with metric_col3:
    st.markdown(
        f"""
        <div class="metric-box">
            <div class="metric-number">{chua_gui}</div>
            <div class="metric-label">Chưa gửi</div>
        </div>
        """,
        unsafe_allow_html=True
    )

st.write("")


# ============================================================
# 17. THANH TIẾN TRÌNH
# ============================================================

st.progress((current_index + 1) / tong_so_khach)
st.caption(f"Đang xử lý khách hàng {current_index + 1} / {tong_so_khach}")


# ============================================================
# 18. FLASHCARD THÔNG TIN KHÁCH HÀNG
# ============================================================

status_class = "status-done" if trang_thai == "Đã gửi Zalo" else "status-pending"

st.markdown(
    f"""
    <div class="customer-card">
        <div class="customer-name">👤 Khách hàng {current_index + 1} / {tong_so_khach}: {ten_kh}</div>
        <div class="info-line"><b>📞 Số điện thoại:</b> {sdt}</div>
        <div class="info-line"><b>📍 Khu vực:</b> {khu_vuc}</div>
        <div class="info-line"><b>🏠 Địa chỉ:</b> {dia_chi}</div>
        <div class="info-line"><b>📌 Trạng thái:</b> <span class="{status_class}">{trang_thai}</span></div>
    </div>
    """,
    unsafe_allow_html=True
)


# ============================================================
# 19. KHU VỰC COPY 1 CHẠM
# ============================================================

copy_col1, copy_col2 = st.columns([1, 2])

with copy_col1:
    st.subheader("📞 Copy số điện thoại")
    st.code(sdt, language="text")

with copy_col2:
    st.subheader("💬 Copy nội dung tin nhắn")
    st.code(tin_nhan, language="text")


# ============================================================
# 20. 🌟 [TƯƠNG LAI - KHU VỰC HIỂN THỊ ẢNH VOUCHER SẼ ĐẶT Ở ĐÂY]
# ============================================================

st.markdown("---")
st.markdown("### 🎁 Khu vực Voucher tương lai")
st.caption("Giai đoạn 1 chỉ xử lý tin nhắn text. Khu vực này đã chừa sẵn để sau này gắn tính năng tạo ảnh bằng Pillow.")

# 🌟 [TƯƠNG LAI - GỌI HÀM TẠO ẢNH VOUCHER Ở ĐÂY]
# voucher_path = tao_anh_voucher_tuong_lai(
#     ten_kh=ten_kh,
#     khu_vuc=khu_vuc,
#     sdt=sdt
# )

# 🌟 [TƯƠNG LAI - HIỂN THỊ ẢNH VOUCHER Ở ĐÂY]
# st.image(voucher_path, caption="Voucher ưu đãi dành riêng cho khách hàng")

# 🌟 [TƯƠNG LAI - NÚT TẢI ẢNH VOUCHER Ở ĐÂY]
# with open(voucher_path, "rb") as file:
#     st.download_button(
#         label="⬇️ Tải ảnh Voucher",
#         data=file,
#         file_name=f"voucher_{sdt}.png",
#         mime="image/png"
#     )


# ============================================================
# 21. NÚT ĐIỀU HƯỚNG
# ============================================================

st.markdown("---")

nav_col1, nav_col2, nav_col3 = st.columns([1, 1, 1])

with nav_col1:
    st.button(
        "⬅️ Quay lại",
        on_click=quay_lai_khach_truoc,
        disabled=current_index == 0,
        use_container_width=True
    )

with nav_col2:
    st.markdown(
        f"<h4 style='text-align:center;'>Lead {current_index + 1} / {tong_so_khach}</h4>",
        unsafe_allow_html=True
    )

with nav_col3:
    st.button(
        "Đã nhắn Zalo ➡️",
        on_click=ghi_nhan_da_gui_zalo,
        disabled=current_index == tong_so_khach - 1 and trang_thai == "Đã gửi Zalo",
        use_container_width=True
    )


# ============================================================
# 22. TẢI BÁO CÁO EXCEL CUỐI NGÀY
# ============================================================

st.markdown("---")
st.subheader("⬇️ Xuất báo cáo cuối ngày")

excel_file = tao_file_excel_download(st.session_state.df_leads)

st.download_button(
    label="⬇️ Tải báo cáo .xlsx",
    data=excel_file,
    file_name="bao_cao_lead_zalo_b2a_vnpt.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    use_container_width=True
)


# ============================================================
# 23. XEM BẢNG DỮ LIỆU ĐÃ CẬP NHẬT
# ============================================================

with st.expander("📊 Xem bảng dữ liệu hiện tại"):
    st.dataframe(st.session_state.df_leads, use_container_width=True)
