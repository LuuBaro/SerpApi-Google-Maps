import os
import time
import streamlit as st
import pandas as pd

from main import (
    load_key_or_raise,
    build_queries,
    crawl_cities,
    postprocess_dataframe,
    export_excel_bytes
)

st.set_page_config(page_title="SerpApi Google Maps Crawler", layout="wide")

st.title("🗺️ SerpApi Google Maps – Crawler (Coworking & Cafe làm việc)")
st.caption("Thu thập HN/HCM, lọc – chống trùng – xuất Excel đẹp theo mapping khách hàng.")

# --- Sidebar config
with st.sidebar:
    st.header("Cấu hình")

    # API key from env/.env handled inside main.py, but allow override
    key_override = st.text_input("SERPAPI_KEY (tùy chọn, nếu không dùng .env)", type="password")

    cities = st.multiselect(
    "Chọn thành phố",
    ["TP.HCM", "Hà Nội", "Hải Phòng", "Quảng Ninh"],
    default=["TP.HCM", "Hà Nội"]
)

    st.subheader("Query pack")
    use_core = st.checkbox("Core (coworking + cafe làm việc)", value=True)
    use_brands = st.checkbox("Brands/chuỗi lớn", value=True)
    use_districts = st.checkbox("Theo quận (tăng đủ 2.000)", value=True)

    st.subheader("Crawler")
    max_pages = st.slider("Số trang mỗi query (mỗi trang ~20)", min_value=1, max_value=20, value=10, step=1)
    sleep_sec = st.slider("Delay giữa request (giây)", min_value=0.0, max_value=2.0, value=0.3, step=0.1)

    st.subheader("Lọc")
    min_rating = st.slider("Rating tối thiểu (0 = không lọc)", 0.0, 5.0, 0.0, 0.1)
    min_reviews = st.number_input("Số review tối thiểu (0 = không lọc)", min_value=0, value=0, step=10)

    st.subheader("Xuất file")
    output_mode = st.radio(
    "Chế độ xuất",
    [
        "Theo tỉnh/TP (mỗi sheet 1 tỉnh/TP)",
        "Theo khu vực (mỗi sheet 1 khu vực)",
        "1 file (2 sheet: HN + HCM)",
        "1 file (sheet duy nhất)",
    ],
    index=0
)

    include_raw_cols = st.checkbox("Giữ thêm cột raw (rating, phone, ...)", value=False)

    st.divider()
    st.info("Gợi ý: Nếu mục tiêu 1.500–2.000, bật **Theo quận** và đặt pages 8–15.")

# --- Validate
try:
    load_key_or_raise(key_override if key_override else None)
except Exception as e:
    st.error(str(e))
    st.stop()

# Build queries preview
queries_preview = build_queries(
    cities=cities,
    use_core=use_core,
    use_brands=use_brands,
    use_districts=use_districts
)

st.subheader("📌 Preview query (tự động)")
colA, colB = st.columns([2, 1])
with colA:
    st.write("Tổng số query:", sum(len(v) for v in queries_preview.values()))
with colB:
    st.write({k: len(v) for k, v in queries_preview.items()})

with st.expander("Xem danh sách query"):
    for city, qs in queries_preview.items():
        st.markdown(f"**{city}** ({len(qs)} queries)")
        st.code("\n".join(qs[:120]) + ("\n...\n" if len(qs) > 120 else ""))

# --- Run crawler
st.subheader("🚀 Chạy crawl")

if "running" not in st.session_state:
    st.session_state.running = False
if "df" not in st.session_state:
    st.session_state.df = None

run_col1, run_col2 = st.columns([1, 3])

with run_col1:
    start_btn = st.button("▶ Start", use_container_width=True, disabled=st.session_state.running)
with run_col2:
    stop_btn = st.button("⏹ Stop", use_container_width=True, disabled=not st.session_state.running)

progress = st.progress(0)
status = st.empty()

if stop_btn:
    st.session_state.running = False
    status.warning("Đã gửi yêu cầu dừng. Sẽ dừng ở vòng lặp gần nhất.")

if start_btn:
    st.session_state.running = True
    status.info("Đang crawl...")

    # Crawl with live progress
    all_rows = []
    total_queries = sum(len(v) for v in queries_preview.values())
    done_queries = 0

    def on_progress(msg, done, total):
        progress.progress(min(1.0, done / max(1, total)))
        status.info(msg)

    rows = crawl_cities(
        cities=cities,
        queries_by_city=queries_preview,
        max_pages_per_query=max_pages,
        sleep_sec=sleep_sec,
        min_rating=min_rating,
        min_reviews=min_reviews,
        on_progress=on_progress,
        should_stop=lambda: (not st.session_state.running),
    )
    all_rows.extend(rows)

    st.session_state.running = False

    df = postprocess_dataframe(pd.DataFrame(all_rows))
    st.session_state.df = df
    status.success(f"Hoàn tất! Tổng địa điểm sau dedupe + lọc: {len(df):,}")

# --- Show results
st.subheader("📊 Kết quả")
df = st.session_state.df
if df is None or df.empty:
    st.info("Chưa có dữ liệu. Hãy bấm Start.")
else:
    st.write("Số dòng:", len(df))
    st.dataframe(df.head(200), use_container_width=True)

    # Export
    st.subheader("⬇️ Xuất Excel")
    excel_bytes = export_excel_bytes(
        df=df,
        mode=output_mode,
        include_raw_cols=include_raw_cols,
        filename="coworking_cafe_HN_HCM.xlsx",
    )
    st.download_button(
        "Tải file Excel",
        data=excel_bytes,
        file_name="coworking_cafe_HN_HCM.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
