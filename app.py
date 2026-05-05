import streamlit as st
import io
import zipfile
from pathlib import Path
from core import process_export

TEMPLATES_DIR = Path(__file__).parent / "templates"

st.set_page_config(page_title="Transfer file Export", layout="wide")

st.title("📚 Transfer file Export")

# ── Section 1: Upload file data ──────────────────────────────────────────────
st.header("1. Tải lên file dữ liệu")
uploaded_files = st.file_uploader(
    "Chọn 1 hoặc nhiều file data (.xlsx)",
    type=["xlsx"],
    accept_multiple_files=True,
    help="File data cần có sheet 'GIAO-VIEN' và/hoặc 'Danh sách HS toàn trường'",
)

# ── Section 2: Loại export ───────────────────────────────────────────────────
st.header("2. Chọn kiểu xuất file Export")
export_type = st.radio(
    # "Chọn kiểu xuất file:",
    options=["File Export gộp chung", "File Export tách lẻ", "File Export Tiểu học"],
    index=0,
    horizontal=True,
)

# ── Section 3: Thông tin trường & admin ──────────────────────────────────────
st.header("3. Thông tin trường & Admin")

col_label, col_tk, col_mk = st.columns([3, 3, 3])
col_label.markdown("**Chức danh**")
col_tk.markdown("**Tài khoản**")
col_mk.markdown("**Mật khẩu**")

with col_label:
    ten_truong = st.text_input("Tên trường", key="ten_truong")
    ten_hieu_truong = st.text_input("Tên Hiệu trưởng", key="ten_ht")
    ten_hp1 = st.text_input("Tên Hiệu phó 1", key="ten_hp1")
    ten_hp2 = st.text_input("Tên Hiệu phó 2", key="ten_hp2")
    ten_hp3 = st.text_input("Tên Hiệu phó 3", key="ten_hp3")

with col_tk:
    tk_admin = st.text_input("Tài khoản Admin", key="tk_admin")
    tk_ht = st.text_input("Tài khoản Hiệu trưởng", key="tk_ht")
    tk_hp1 = st.text_input("Tài khoản Hiệu phó 1", key="tk_hp1")
    tk_hp2 = st.text_input("Tài khoản Hiệu phó 2", key="tk_hp2")
    tk_hp3 = st.text_input("Tài khoản Hiệu phó 3", key="tk_hp3")

with col_mk:
    mk_admin = st.text_input("Mật khẩu Admin", key="mk_admin")
    mk_ht = st.text_input("Mật khẩu Hiệu trưởng", key="mk_ht")
    mk_hp1 = st.text_input("Mật khẩu Hiệu phó 1", key="mk_hp1")
    mk_hp2 = st.text_input("Mật khẩu Hiệu phó 2", key="mk_hp2")
    mk_hp3 = st.text_input("Mật khẩu Hiệu phó 3", key="mk_hp3")

admin_info = {
    "ten_truong": ten_truong.strip(),
    "tk_admin": tk_admin.strip(),
    "mk_admin": mk_admin.strip(),
    "rows": [
        {"ten": ten_hieu_truong.strip(), "tk": tk_ht.strip(), "mk": mk_ht.strip()},
        {"ten": ten_hp1.strip(), "tk": tk_hp1.strip(), "mk": mk_hp1.strip()},
        {"ten": ten_hp2.strip(), "tk": tk_hp2.strip(), "mk": mk_hp2.strip()},
        {"ten": ten_hp3.strip(), "tk": tk_hp3.strip(), "mk": mk_hp3.strip()},
    ],
}

# ── Section 4: Xuất file ──────────────────────────────────────────────────────
st.header("4. Xuất file")

if st.button("🚀 Tạo file export", type="primary", use_container_width=True):
    if not uploaded_files:
        st.error("⚠️ Vui lòng tải lên ít nhất 1 file dữ liệu.")
    else:
        with st.spinner("Đang xử lý..."):
            try:
                result = process_export(
                    uploaded_files=uploaded_files,
                    export_type=export_type,
                    admin_info=admin_info,
                    templates_dir=TEMPLATES_DIR,
                    ten_truong=ten_truong.strip(),
                )
                st.success(f"✅ Tạo thành công {len(result)} file!")

                if len(result) == 1:
                    fname, fdata = result[0]
                    st.download_button(
                        label=f"⬇️ Tải về: {fname}",
                        data=fdata,
                        file_name=fname,
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        use_container_width=True,
                    )
                else:
                    # Zip tất cả files lại
                    zip_buf = io.BytesIO()
                    with zipfile.ZipFile(zip_buf, "w", zipfile.ZIP_DEFLATED) as zf:
                        for fname, fdata in result:
                            zf.writestr(fname, fdata)
                    zip_buf.seek(0)
                    zip_name = f"TK Onluyen - {ten_truong.strip() or 'Export'}.zip"
                    st.download_button(
                        label=f"⬇️ Tải về tất cả ({len(result)} files): {zip_name}",
                        data=zip_buf.getvalue(),
                        file_name=zip_name,
                        mime="application/zip",
                        use_container_width=True,
                    )
                    st.markdown("**Danh sách file trong zip:**")
                    for fname, _ in result:
                        st.write(f"  • {fname}")

            except Exception as e:
                st.error(f"❌ Lỗi: {e}")
                import traceback
                st.code(traceback.format_exc())
