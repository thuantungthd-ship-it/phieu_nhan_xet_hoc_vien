import base64
import os
import re
import shutil
import tempfile
import time

import streamlit as st
import streamlit.components.v1 as components

import engine

st.set_page_config(
    page_title="Phiếu Nhận Xét Học Viên | CIE VIETNAM",
    page_icon="🎓",
    layout="centered",
    initial_sidebar_state="collapsed",
)

# =========================================================================
# CSS - giao dien song dong
# =========================================================================
st.markdown(
    """
    <style>
    @import url('https://fonts.googleapis.com/css2?family=Be+Vietnam+Pro:wght@400;500;600;700;800&display=swap');

    html, body, [class*="css"] {
        font-family: 'Be Vietnam Pro', sans-serif;
    }

    #MainMenu, footer, header {visibility: hidden;}

    .stApp {
        background: linear-gradient(160deg, #f4f7ff 0%, #eef2ff 40%, #fdf2f8 100%);
    }

    .hero {
        background: linear-gradient(120deg, #1f4e79 0%, #2f6fb0 55%, #c0392b 130%);
        border-radius: 22px;
        padding: 34px 30px 30px 30px;
        margin-bottom: 26px;
        box-shadow: 0 14px 40px rgba(31, 78, 121, 0.25);
        text-align: center;
        color: white;
        position: relative;
        overflow: hidden;
    }
    .hero::before {
        content: "";
        position: absolute;
        top: -60px; right: -60px;
        width: 180px; height: 180px;
        background: rgba(255,255,255,0.12);
        border-radius: 50%;
    }
    .hero::after {
        content: "";
        position: absolute;
        bottom: -50px; left: -40px;
        width: 140px; height: 140px;
        background: rgba(255,255,255,0.08);
        border-radius: 50%;
    }
    .hero h1 {
        font-size: 1.85rem;
        font-weight: 800;
        margin: 6px 0 4px 0;
        letter-spacing: 0.2px;
    }
    .hero p {
        font-size: 0.98rem;
        opacity: 0.92;
        margin: 0;
        max-width: 560px;
        margin-left: auto;
        margin-right: auto;
    }
    .hero .badge {
        display: inline-block;
        background: rgba(255,255,255,0.18);
        border: 1px solid rgba(255,255,255,0.35);
        border-radius: 999px;
        padding: 4px 14px;
        font-size: 0.78rem;
        font-weight: 600;
        letter-spacing: 0.5px;
        margin-bottom: 10px;
    }

    .stat-row {display: flex; gap: 14px; margin-bottom: 22px;}
    .stat-card {
        flex: 1;
        background: white;
        border-radius: 16px;
        padding: 16px 14px;
        text-align: center;
        box-shadow: 0 6px 18px rgba(31, 78, 121, 0.08);
        border: 1px solid #eef1f8;
    }
    .stat-card .num {
        font-size: 1.6rem;
        font-weight: 800;
        color: #1f4e79;
        line-height: 1.1;
    }
    .stat-card .lbl {
        font-size: 0.78rem;
        color: #6b7280;
        margin-top: 3px;
        font-weight: 500;
    }

    .upload-card {
        background: white;
        border-radius: 18px;
        padding: 22px 22px 10px 22px;
        box-shadow: 0 8px 26px rgba(31, 78, 121, 0.10);
        border: 1px solid #eef1f8;
        margin-bottom: 18px;
    }
    .upload-card h3 {
        margin-top: 0;
        color: #1f4e79;
        font-size: 1.08rem;
    }

    div[data-testid="stFileUploader"] section {
        border: 2px dashed #94b3d6 !important;
        border-radius: 14px !important;
        background: #f7faff !important;
    }
    div[data-testid="stFileUploader"] section:hover {
        border-color: #1f4e79 !important;
        background: #eef4fc !important;
    }

    .stButton > button, .stDownloadButton > button {
        border-radius: 10px !important;
        font-weight: 600 !important;
        transition: transform 0.15s ease;
    }
    .stButton > button:hover, .stDownloadButton > button:hover {
        transform: translateY(-1px);
    }
    .stButton > button[kind="primary"] {
        background: linear-gradient(120deg, #1f4e79, #2f6fb0) !important;
        border: none !important;
    }

    .history-item {
        display: flex;
        justify-content: space-between;
        padding: 8px 4px;
        border-bottom: 1px solid #f0f2f6;
        font-size: 0.92rem;
    }
    .history-item:last-child {border-bottom: none;}
    .history-item .cnt {
        background: #e8f0fe;
        color: #1f4e79;
        border-radius: 999px;
        padding: 1px 10px;
        font-weight: 700;
        font-size: 0.8rem;
    }

    footer-note {
        text-align: center;
        color: #9aa3b2;
        font-size: 0.8rem;
        margin-top: 24px;
    }
    </style>
    """,
    unsafe_allow_html=True,
)

# =========================================================================
# Anh nen he thong (CO DINH cho tat ca moi nguoi, khong the tuy chinh)
# Khi co file anh, dan base64 vao engine.DEFAULT_BG_B64 (giong cach lam voi logo).
# =========================================================================
if getattr(engine, "DEFAULT_BG_B64", ""):
    st.markdown(
        f"""
        <style>
        .stApp {{
            background-image:
                linear-gradient(160deg, rgba(244,247,255,0.92) 0%, rgba(238,242,255,0.90) 40%, rgba(253,242,248,0.92) 100%),
                url("data:image/png;base64,{engine.DEFAULT_BG_B64}");
            background-size: cover;
            background-position: center;
            background-attachment: fixed;
        }}
        </style>
        """,
        unsafe_allow_html=True,
    )

# =========================================================================
# Logo he thong - CO DINH DUY NHAT, khong ai chinh sua duoc qua giao dien
# =========================================================================
if "logo_path" not in st.session_state:
    tmp_dir = tempfile.mkdtemp()
    logo_path = os.path.join(tmp_dir, "logo.png")
    with open(logo_path, "wb") as f:
        f.write(base64.b64decode(engine.DEFAULT_LOGO_B64))
    st.session_state["logo_path"] = logo_path

engine.COMPANY_LOGO_PATH = st.session_state["logo_path"]

# =========================================================================
# Man hinh cho (splash) - hien logo truoc khi vao giao dien chinh.
# Chi hien 1 lan moi phien lam viec (session_state luu lai, khong lap lai
# moi khi nguoi dung bam nut trong app).
# =========================================================================
if "splash_shown" not in st.session_state:
    splash = st.empty()
    with splash.container():
        st.markdown(
            f"""
            <div style="
                position: fixed; inset: 0; z-index: 99999;
                background: linear-gradient(135deg, #1f4e79 0%, #2f6fb0 55%, #c0392b 130%);
                display: flex; flex-direction: column; align-items: center; justify-content: center;">
                <img src="data:image/png;base64,{engine.DEFAULT_LOGO_B64}"
                     style="width: 130px; margin-bottom: 18px; animation: pulse 1.4s ease-in-out infinite;">
                <div style="color: white; font-family: 'Be Vietnam Pro', sans-serif; font-weight: 600; font-size: 0.95rem; letter-spacing: 0.5px;">
                    Đang tải hệ thống...
                </div>
            </div>
            <style>
            @keyframes pulse {{
                0%, 100% {{ transform: scale(1); opacity: 1; }}
                50% {{ transform: scale(1.08); opacity: 0.85; }}
            }}
            </style>
            """,
            unsafe_allow_html=True,
        )
    time.sleep(1.3)
    splash.empty()
    st.session_state["splash_shown"] = True

if "uploader_key" not in st.session_state:
    st.session_state["uploader_key"] = 0
if "batch_counter" not in st.session_state:
    st.session_state["batch_counter"] = 0
if "history" not in st.session_state:
    st.session_state["history"] = []
if "total_students" not in st.session_state:
    st.session_state["total_students"] = 0


def safe_name(s):
    return re.sub(r'[\\/:*?"<>|]', "_", str(s)).strip()


def build_zip_for_students(students):
    tmp_dir = tempfile.mkdtemp()
    used_names = set()
    for s in students:
        doc = engine.build_phieu(s)
        lop = safe_name(s["class_info"].get("lop") or s["class_info"].get("sheet") or "Lop")
        ten = safe_name(s["name"])
        base = f"{lop}__{ten}"
        fname = base
        i = 2
        while fname in used_names:
            fname = f"{base}_{i}"
            i += 1
        used_names.add(fname)
        doc.save(os.path.join(tmp_dir, fname + ".docx"))
    zip_base = os.path.join(tempfile.mkdtemp(), "PhieuNhanXet")
    zip_path = shutil.make_archive(zip_base, "zip", tmp_dir)
    with open(zip_path, "rb") as f:
        return f.read()


def trigger_browser_download(filename, data_bytes):
    b64 = base64.b64encode(data_bytes).decode()
    html = f"""
    <html><body>
    <script>
    const b64data = "{b64}";
    const byteChars = atob(b64data);
    const byteNumbers = new Array(byteChars.length);
    for (let i = 0; i < byteChars.length; i++) {{
        byteNumbers[i] = byteChars.charCodeAt(i);
    }}
    const byteArray = new Uint8Array(byteNumbers);
    const blob = new Blob([byteArray], {{type: "application/zip"}});
    const url = window.URL.createObjectURL(blob);
    const a = document.createElement("a");
    a.href = url;
    a.download = "{filename}";
    document.body.appendChild(a);
    a.click();
    setTimeout(() => {{ window.URL.revokeObjectURL(url); }}, 3000);
    </script>
    </body></html>
    """
    components.html(html, height=0, width=0)


# =========================================================================
# Hero header
# =========================================================================
st.markdown(
    """
    <div class="hero">
        <div class="badge">🎓 CIE VIETNAM · AUTO REPORT TOOL</div>
        <h1>Tạo Phiếu Nhận Xét & Đánh Giá Học Viên</h1>
        <p>Upload file điểm — nhận ngay bộ phiếu Word chuẩn mẫu công ty cho từng học viên,
        kèm nhận xét, kiến nghị và lộ trình ôn tập tự động.</p>
    </div>
    """,
    unsafe_allow_html=True,
)

# Nhac nen (tuy chon) - widget nho, dat ngay duoi hero, co nut bat/tat
# (Luu y ky thuat: components.html chay trong 1 iframe rieng cua Streamlit,
#  nen KHONG dung position:fixed o day - iframe fixed=0x0 se lam nut bi an mat.
#  Thay vao do dat widget noi lien trong dong chay cua trang, voi kich thuoc du hien nut.)
# =========================================================================
if st.session_state.get("bg_music_b64"):
    music_b64 = st.session_state["bg_music_b64"]
    components.html(
        f"""
        <div style="display:flex; align-items:center; gap:10px; font-family: 'Be Vietnam Pro', sans-serif;">
          <button id="music-toggle-btn" style="
              width: 42px; height: 42px; border-radius: 50%; border: none;
              background: linear-gradient(120deg, #1f4e79, #2f6fb0); color: white;
              font-size: 18px; cursor: pointer; box-shadow: 0 4px 12px rgba(31,78,121,0.30);
              flex-shrink: 0;">
            🔇
          </button>
          <span style="font-size: 0.85rem; color: #4b5563;">Nhạc nền (bấm để bật/tắt)</span>
        </div>
        <audio id="bg-audio" loop>
          <source src="data:audio/mp3;base64,{music_b64}" type="audio/mpeg">
        </audio>
        <script>
        const audio = document.getElementById('bg-audio');
        const btn = document.getElementById('music-toggle-btn');
        audio.volume = 0.4;
        let playing = false;
        audio.muted = true;
        audio.play().catch(() => {{}});
        btn.addEventListener('click', () => {{
            if (!playing) {{
                audio.muted = false;
                audio.play();
                btn.textContent = '🔊';
                playing = true;
            }} else {{
                audio.pause();
                btn.textContent = '🔇';
                playing = false;
            }}
        }});
        </script>
        """,
        height=50,
    )


st.markdown(
    f"""
    <div class="stat-row">
        <div class="stat-card"><div class="num">{len(st.session_state['history'])}</div><div class="lbl">Lớp đã xử lý</div></div>
        <div class="stat-card"><div class="num">{st.session_state['total_students']}</div><div class="lbl">Học viên đã tạo phiếu</div></div>
        <div class="stat-card"><div class="num">∞</div><div class="lbl">Số lần upload tiếp theo</div></div>
    </div>
    """,
    unsafe_allow_html=True,
)

# =========================================================================
# Upload card
# =========================================================================
st.markdown('<div class="upload-card">', unsafe_allow_html=True)
st.markdown("### 📁 Upload file điểm để bắt đầu")
st.caption("Hỗ trợ file 1 lớp hoặc cả workbook nhiều lớp (.xlsx). Xử lý xong sẽ tự động tải file .zip về máy.")

uploaded_file = st.file_uploader(
    "Chọn file Excel",
    type=["xlsx"],
    key=f"uploader_{st.session_state['uploader_key']}",
    label_visibility="collapsed",
)
st.markdown("</div>", unsafe_allow_html=True)

if uploaded_file is not None:
    progress = st.progress(0, text="Đang đọc file...")
    tmp_path = os.path.join(tempfile.mkdtemp(), uploaded_file.name)
    with open(tmp_path, "wb") as f:
        f.write(uploaded_file.getbuffer())

    try:
        progress.progress(30, text="Đang tính điểm, xếp loại, chọn nhận xét phù hợp...")
        students = engine.process_workbook(tmp_path)
    except Exception as e:
        progress.empty()
        st.error(f"❌ Không đọc được file này: {e}")
        students = []

    if students:
        progress.progress(70, text="Đang tạo phiếu Word cho từng học viên...")
        zip_data = build_zip_for_students(students)
        progress.progress(100, text="Hoàn tất!")
        progress.empty()

        st.session_state["batch_counter"] += 1
        st.session_state["total_students"] += len(students)
        zip_filename = f"PhieuNhanXet_{st.session_state['batch_counter']:02d}_{safe_name(uploaded_file.name).replace('.xlsx','')}.zip"
        st.session_state["history"].append((uploaded_file.name, len(students)))

        st.success(f"✔ Đã tạo **{len(students)} phiếu** từ file **{uploaded_file.name}**. Đang tự động tải về...")
        st.balloons()
        trigger_browser_download(zip_filename, zip_data)

        st.download_button(
            "⬇️ Nếu trình duyệt chặn tự động tải, bấm vào đây để tải thủ công",
            data=zip_data,
            file_name=zip_filename,
            mime="application/zip",
        )

        st.session_state["uploader_key"] += 1
        st.button("➕ Upload lớp tiếp theo", type="primary", use_container_width=True)
    elif uploaded_file is not None and not students:
        progress.empty()
        st.warning("⚠️ Không tìm thấy học viên nào trong file này. Kiểm tra lại cấu trúc file.")

# =========================================================================
# Lich su + tuy chinh
# =========================================================================
if st.session_state["history"]:
    with st.expander(f"📊 Lịch sử đã xử lý ({len(st.session_state['history'])} lớp · {st.session_state['total_students']} học viên)"):
        for fn, n in st.session_state["history"]:
            st.markdown(
                f'<div class="history-item"><span>📄 {fn}</span><span class="cnt">{n} học viên</span></div>',
                unsafe_allow_html=True,
            )

with st.expander("🎵 Nhạc nền cho trang web (tuỳ chọn)"):
    st.caption(
        "Nhạc chỉ áp dụng cho **giao diện web này**, không ảnh hưởng gì tới phiếu Word được tạo ra."
    )
    new_music = st.file_uploader("Nhạc nền (.mp3)", type=["mp3"], key="music_uploader")
    if new_music is not None:
        st.session_state["bg_music_b64"] = base64.b64encode(new_music.getbuffer()).decode()
        st.success("Đã đặt nhạc nền mới — nút bật/tắt nhạc sẽ xuất hiện phía trên trang.")
        st.rerun()

    if st.session_state.get("bg_music_b64") and st.button("🗑️ Bỏ nhạc nền"):
        del st.session_state["bg_music_b64"]
        st.rerun()

st.markdown('<div class="footer-note">Made for CIE VIETNAM · Powered by Streamlit</div>', unsafe_allow_html=True)
