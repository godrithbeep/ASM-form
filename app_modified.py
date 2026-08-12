# -*- coding: utf-8 -*-
"""
Form Kế hoạch bán hàng - Salesman
==================================
Luồng:
  1) Đăng nhập bằng Mã NV -> tự hiện Tên + Vùng bán hàng
  2) Bước 1: nhập Khách hàng (C1) x Sản phẩm x Tháng (số lượng theo THÙNG).
     - Khách hàng, Tỉnh & Sản phẩm bắt buộc.
     - Sản phẩm chỉ gợi ý theo đúng Vùng của salesman.
     - Khách hàng: chọn từ danh sách đã từng nhập (dropdown gợi ý) hoặc thêm KH mới qua ô riêng.
     - Tỉnh chọn theo Vùng, tự đồng bộ theo từng Khách hàng đã có Tỉnh — KHÔNG BAO GIỜ ghi đè
       một Tỉnh mà dòng đó đã tự có (fix lỗi "nhập xong mất data").
     - Bảng "Tổng quan theo Khách hàng x Tháng" tự tính, đặt ở đầu trang, cao bằng bảng chi tiết.
  3) Bước 2: phân bổ tỷ lệ % theo cây trồng cho từng (Tỉnh, Sản phẩm, Tháng) đã gom nhóm.
     - Kg/L = Số lượng (thùng) x Carton weight (kg/thùng)
     - Target VNĐ = Số lượng (thùng) x Đơn giá
     - Mỗi dòng bắt buộc tổng % = 100 mới hợp lệ. Cột "Tổng %" tự tô xanh/đỏ.
  4) Xuất file Excel (nhiều sheet) để salesman tải về và gửi lại.

File này đọc dữ liệu tham chiếu (danh sách salesman / tỉnh / sản phẩm theo vùng)
từ "reference_data.json" nằm cùng thư mục.
"""

import json
import io
import traceback
from datetime import datetime, date
from pathlib import Path

import pandas as pd
import streamlit as st

# ------------------------------------------------------------------
# CONFIG & REFERENCE DATA
# ------------------------------------------------------------------
st.set_page_config(page_title="Kế hoạch bán hàng - Salesman", layout="wide")

DATA_PATH = Path(__file__).parent / "reference_data.json"

@st.cache_data
def load_reference_data():
    with open(DATA_PATH, "r", encoding="utf-8") as f:
        raw = json.load(f)
    salesman_df = pd.DataFrame(raw["salesman"])          # code, name, region
    tinh_df = pd.DataFrame(raw["tinh"])                   # region, province
    products_df = pd.DataFrame(raw["products"])           # region, name, short, carton, price
    months = raw["months"]                                 # ['2026-10', ...]
    return salesman_df, tinh_df, products_df, months

salesman_df, tinh_df, products_df, MONTHS = load_reference_data()

def month_label(m: str) -> str:
    """'2026-10' -> '10/2026' (ngắn gọn, không tiền tố)."""
    y, mo = m.split("-")
    return f"{int(mo):02d}/{y}"

MONTH_LABELS = [month_label(m) for m in MONTHS]
QTY_COLS = MONTH_LABELS[:]   # cột số lượng dùng thẳng nhãn tháng làm tên cột, vd "10/2026"

CROPS = ["Durian", "Coffee", "Rice", "Dragon fruit", "Mango", "Vegetable & others"]
CROP_LABELS = {
    "Durian": "🌾 Sầu riêng (%)",
    "Coffee": "🌾 Cà phê (%)",
    "Rice": "🌾 Lúa (%)",
    "Dragon fruit": "🌾 Thanh long (%)",
    "Mango": "🌾 Xoài (%)",
    "Vegetable & others": "🌾 Rau màu & khác (%)",
}

DETAIL_COLS = ["Khách hàng (C1)", "Tỉnh", "Sản phẩm", "Đơn giá"] + QTY_COLS + ["Thành tiền"]

# ------------------------------------------------------------------
# SESSION STATE INIT
# ------------------------------------------------------------------
if "user" not in st.session_state:
    st.session_state.user = None
if "step" not in st.session_state:
    st.session_state.step = 1
if "detail_df" not in st.session_state:
    st.session_state.detail_df = pd.DataFrame(
        [{"Khách hàng (C1)": "", "Tỉnh": "", "Sản phẩm": "", "Đơn giá": 0,
          **{c: 0 for c in QTY_COLS}, "Thành tiền": 0}]
    )
if "crop_df" not in st.session_state:
    st.session_state.crop_df = None  # built lazily entering step 2
if "known_customers" not in st.session_state:
    st.session_state.known_customers = []  # danh sách KH đã dùng, dùng làm gợi ý dropdown


# ------------------------------------------------------------------
# HELPERS
# ------------------------------------------------------------------
def region_products(region: str) -> pd.DataFrame:
    return products_df[products_df["region"] == region]

def region_provinces(region: str) -> list:
    return sorted(tinh_df[tinh_df["region"] == region]["province"].tolist())

def product_lookup(region: str, name: str):
    """Return dict(price, carton, short) for a product name within region, or None."""
    rp = region_products(region)
    row = rp[rp["name"] == name]
    if row.empty:
        return None
    r = row.iloc[0]
    return {"price": float(r["price"]), "carton": float(r["carton"]), "short": r["short"]}

def sync_known_customers(df: pd.DataFrame):
    """Đưa mọi tên KH đã xuất hiện trong bảng vào danh sách gợi ý (không xoá cái cũ)."""
    names = sorted(set(df["Khách hàng (C1)"].fillna("").astype(str).str.strip()) - {""})
    changed = False
    for n in names:
        if n not in st.session_state.known_customers:
            st.session_state.known_customers.append(n)
            changed = True
    if changed:
        st.session_state.known_customers.sort()

def recompute_detail(df: pd.DataFrame, region: str) -> pd.DataFrame:
    """Recompute Đơn giá / Thành tiền, and sync Tỉnh across rows sharing the same customer.

    QUAN TRỌNG: chỉ điền Tỉnh cho các dòng đang TRỐNG Tỉnh, dựa theo Tỉnh đã có sẵn của
    cùng khách hàng đó ở dòng khác. KHÔNG BAO GIỜ ghi đè lên một Tỉnh mà dòng đó đã có sẵn
    (kể cả khi khác với dòng khác) — đây là nguyên nhân gây mất dữ liệu vừa nhập trước đây.
    """
    df = df.copy()
    for col in ["Khách hàng (C1)", "Tỉnh", "Sản phẩm"]:
        if col not in df.columns:
            df[col] = ""
    df["Khách hàng (C1)"] = df["Khách hàng (C1)"].fillna("").astype(str).str.strip()
    df["Sản phẩm"] = df["Sản phẩm"].fillna("").astype(str).str.strip()
    df["Tỉnh"] = df["Tỉnh"].fillna("").astype(str).str.strip()

    # auto price / carton weight from product (deterministic, an toàn để ghi đè mỗi lần)
    prices, cartons = [], []
    for _, row in df.iterrows():
        info = product_lookup(region, row["Sản phẩm"]) if row["Sản phẩm"] else None
        prices.append(info["price"] if info else 0)
        cartons.append(info["carton"] if info else 0)
    df["Đơn giá"] = prices
    df["_carton_weight"] = cartons

    # sync province: chỉ điền cho dòng ĐANG TRỐNG, ưu tiên Tỉnh đã có sẵn của cùng KH đó
    prov_map = {}
    for _, row in df.iterrows():
        cust, prov = row["Khách hàng (C1)"], row["Tỉnh"]
        if cust and prov and cust not in prov_map:
            prov_map[cust] = prov

    def fill_province(r):
        if not r["Khách hàng (C1)"]:
            return r["Tỉnh"]
        if r["Tỉnh"]:              # dòng đã có Tỉnh riêng -> giữ nguyên, KHÔNG ghi đè
            return r["Tỉnh"]
        return prov_map.get(r["Khách hàng (C1)"], "")   # dòng trống -> mượn Tỉnh của KH đó nếu có

    df["Tỉnh"] = df.apply(fill_province, axis=1)

    for c in QTY_COLS:
        if c not in df.columns:
            df[c] = 0
        df[c] = pd.to_numeric(df[c], errors="coerce").fillna(0)

    df["Thành tiền"] = df[QTY_COLS].sum(axis=1) * df["Đơn giá"]
    return df

def invalid_mask(df: pd.DataFrame):
    """Bắt buộc: Khách hàng, Tỉnh, Sản phẩm."""
    return (df["Khách hàng (C1)"] == "") | (df["Sản phẩm"] == "") | (df["Tỉnh"] == "")

def build_overview(df: pd.DataFrame) -> pd.DataFrame:
    valid = df[~invalid_mask(df) & (df["Khách hàng (C1)"] != "")]
    if valid.empty:
        return pd.DataFrame(columns=["Khách hàng"] + MONTH_LABELS).set_index("Khách hàng")
    rows = []
    for cust, g in valid.groupby("Khách hàng (C1)"):
        row = {"Khách hàng": cust}
        for lbl in MONTH_LABELS:
            row[lbl] = float((g[lbl] * g["Đơn giá"]).sum())
        rows.append(row)
    out = pd.DataFrame(rows).set_index("Khách hàng")
    out.loc["TỔNG"] = out.sum(numeric_only=True)
    return out

def build_crop_summary(df: pd.DataFrame) -> pd.DataFrame:
    """Group valid rows by (Tỉnh, Sản phẩm, Tháng). Kg/L = SL(thùng) x carton weight.

    Dùng .agg() với named aggregation (ổn định qua mọi phiên bản pandas) thay vì
    .groupby().apply() trả Series — cách cũ dễ vỡ tuỳ phiên bản pandas trên server
    và là nghi phạm chính gây lỗi khi xuất Excel.
    """
    valid = df[~invalid_mask(df) & (df["Tỉnh"] != "")]
    recs = []
    region = st.session_state.user["region"]
    for _, row in valid.iterrows():
        for lbl, m in zip(MONTH_LABELS, MONTHS):
            qty = row[lbl]
            if not qty:
                continue
            info = product_lookup(region, row["Sản phẩm"])
            short = info["short"] if info else row["Sản phẩm"]
            carton_w = info["carton"] if info else 0
            price = info["price"] if info else 0
            recs.append({
                "Tỉnh": row["Tỉnh"],
                "Short name": short,
                "MonthLabel": lbl,
                "_month_sort": m,
                "_kgl": qty * carton_w,
                "_money": qty * price,
            })
    if not recs:
        return pd.DataFrame(columns=["Tỉnh", "Short name", "MonthLabel", "_month_sort", "Kg/L", "Target VNĐ"] + CROPS + ["Tổng %"])

    raw = pd.DataFrame(recs)
    grouped = raw.groupby(["Tỉnh", "Short name", "MonthLabel", "_month_sort"], as_index=False).agg(
        **{"Kg/L": ("_kgl", "sum"), "Target VNĐ": ("_money", "sum")}
    )
    grouped = grouped.sort_values(["Tỉnh", "Short name", "_month_sort"]).reset_index(drop=True)
    for c in CROPS:
        grouped[c] = 0.0
    grouped["Tổng %"] = 0.0
    return grouped

def merge_crop_alloc(new_summary: pd.DataFrame, old_alloc: pd.DataFrame) -> pd.DataFrame:
    """Preserve previously entered % values when the summary is rebuilt."""
    if old_alloc is None or old_alloc.empty:
        return new_summary
    key_cols = ["Tỉnh", "Short name", "MonthLabel"]
    merged = new_summary.merge(
        old_alloc[key_cols + CROPS], on=key_cols, how="left", suffixes=("", "_old")
    )
    for c in CROPS:
        old_c = f"{c}_old"
        if old_c in merged.columns:
            merged[c] = merged[old_c].fillna(merged[c])
            merged.drop(columns=[old_c], inplace=True)
    merged["Tổng %"] = merged[CROPS].sum(axis=1)
    return merged

def fmt_vnd(x):
    try:
        return f"{x:,.0f}"
    except Exception:
        return str(x)


# ------------------------------------------------------------------
# EXPORT TO EXCEL
# ------------------------------------------------------------------
def export_excel(user: dict, detail_df: pd.DataFrame, overview_df: pd.DataFrame, crop_df: pd.DataFrame) -> bytes:
    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as writer:
        # Sheet 0: thông tin
        info_df = pd.DataFrame([{
            "Mã NV": user["code"], "Tên": user["name"], "Vùng": user["region"],
            "Ngày xuất kế hoạch": datetime.now().strftime("%d/%m/%Y %H:%M"),
        }])
        info_df.to_excel(writer, sheet_name="Thong tin", index=False)

        # Sheet 1: chi tiết KH x SP x Tháng
        d = detail_df[~invalid_mask(detail_df)].copy()
        d = d.drop(columns=[c for c in d.columns if c.startswith("_")], errors="ignore")
        d.to_excel(writer, sheet_name="KH x San pham", index=False)

        # Sheet 2: tổng quan KH x Tháng
        overview_df.reset_index().to_excel(writer, sheet_name="Tong quan KH", index=False)

        # Sheet 3: Cây trồng — dùng thẳng nhãn tháng dạng chuỗi (an toàn, không parse datetime)
        cdf = crop_df.copy()
        cdf = cdf.rename(columns={"MonthLabel": "Tháng"})
        out_cols = ["Tỉnh", "Short name", "Tháng", "Kg/L", "Target VNĐ"] + CROPS
        cdf[out_cols].to_excel(writer, sheet_name="Cay trong", index=False)

    return buf.getvalue()


# ------------------------------------------------------------------
# LOGIN SCREEN
# ------------------------------------------------------------------
def login_screen():
    st.markdown("## 🌱 Kế hoạch bán hàng 2026–2027")
    st.caption("Nhập mã nhân viên để tải vùng bán hàng và bắt đầu điền kế hoạch.")
    with st.form("login_form", clear_on_submit=False):
        code = st.text_input("Mã nhân viên", placeholder="VD: E1234")
        submitted = st.form_submit_button("Đăng nhập →", type="primary")
    if submitted:
        match = salesman_df[salesman_df["code"].str.upper() == code.strip().upper()]
        if match.empty:
            st.error(f'Không tìm thấy mã nhân viên "{code}". Vui lòng kiểm tra lại hoặc liên hệ quản lý vùng.')
        else:
            row = match.iloc[0]
            st.session_state.user = {"code": row["code"], "name": row["name"], "region": row["region"]}
            st.rerun()


# ------------------------------------------------------------------
# STEP 1: DETAIL ENTRY
# ------------------------------------------------------------------
def step1_screen():
    user = st.session_state.user
    region = user["region"]

    sync_known_customers(st.session_state.detail_df)

    # ---- thêm khách hàng mới vào danh sách gợi ý ----
    with st.expander("➕ Thêm khách hàng mới (nếu chưa có trong danh sách gợi ý bên dưới)"):
        colA, colB = st.columns([4, 1])
        new_cust = colA.text_input(
            "Tên khách hàng mới", key="new_cust_input",
            label_visibility="collapsed", placeholder="Nhập tên khách hàng mới rồi bấm Thêm",
        )
        if colB.button("+ Thêm KH", width="stretch"):
            name = (new_cust or "").strip()
            if name and name not in st.session_state.known_customers:
                st.session_state.known_customers.append(name)
                st.session_state.known_customers.sort()
                st.success(f'Đã thêm "{name}" — chọn ở cột Khách hàng (C1) trong bảng bên dưới.')
            elif not name:
                st.warning("Nhập tên khách hàng trước khi bấm Thêm.")

    # ---- Overview panel (đầu trang, cao/kéo bằng bảng chi tiết) ----
    st.markdown("##### 📌 Tổng quan theo Khách hàng × Tháng &nbsp;·&nbsp; *tự tính, không cần nhập tay*")
    overview_df = build_overview(st.session_state.detail_df)
    if overview_df.empty:
        st.info("Chưa có khách hàng nào — nhập ở bảng chi tiết bên dưới.")
    else:
        styled = overview_df.style.format(fmt_vnd)
        st.dataframe(styled, height=420, width="stretch")

    st.divider()

    # ---- Detail editor ----
    st.markdown("##### 📝 Chi tiết: Khách hàng × Sản phẩm × Tháng")
    st.caption(
        "Bắt buộc: **Khách hàng**, **Tỉnh** và **Sản phẩm**. Đơn giá tự lấy theo sản phẩm (vùng "
        f"**{region}**). Số lượng nhập theo **thùng**. Khách hàng đã nhập trước đó sẽ hiện trong "
        "dropdown để chọn nhanh — Tỉnh của khách hàng đó cũng tự điền theo, không cần gõ lại."
    )

    product_options = [""] + sorted(region_products(region)["name"].unique().tolist())
    province_options = [""] + region_provinces(region)
    customer_options = [""] + st.session_state.known_customers

    column_config = {
        "Khách hàng (C1)": st.column_config.SelectboxColumn(
            "Khách hàng (C1)", options=customer_options, required=True, width="medium",
            help="Chọn KH đã có trong danh sách, hoặc dùng ô '➕ Thêm khách hàng mới' phía trên nếu là KH mới.",
        ),
        "Tỉnh": st.column_config.SelectboxColumn("Tỉnh", options=province_options, required=True, width="small"),
        "Sản phẩm": st.column_config.SelectboxColumn(
            "Sản phẩm", options=product_options, required=True, width="large",
            help="Gõ để lọc nhanh trong danh sách sản phẩm của vùng bạn.",
        ),
        "Đơn giá": st.column_config.NumberColumn("Đơn giá", format="%,d", disabled=True),
        "Thành tiền": st.column_config.NumberColumn("Thành tiền", format="%,d", disabled=True),
    }
    for c in QTY_COLS:
        column_config[c] = st.column_config.NumberColumn(c, min_value=0, step=1, format="%,d")

    edited = st.data_editor(
        st.session_state.detail_df[DETAIL_COLS],
        column_config=column_config,
        num_rows="dynamic",
        height=420,
        key="detail_editor",
    )

    recomputed = recompute_detail(edited, region)
    st.session_state.detail_df = recomputed
    sync_known_customers(recomputed)

    bad = invalid_mask(recomputed)
    n_bad = int(bad.sum())
    grand_total = recomputed.loc[~bad, "Thành tiền"].sum()

    c1, c2, c3 = st.columns([2, 2, 3])
    c1.metric("Dòng hợp lệ", f"{(~bad).sum()}/{len(recomputed)}")
    c2.metric("Tổng kế hoạch", f"{fmt_vnd(grand_total)} đ")
    with c3:
        st.write("")
        if n_bad > 0:
            st.warning(f"⚠ Còn {n_bad} dòng thiếu Khách hàng, Tỉnh hoặc Sản phẩm.")
        else:
            st.success("✓ Tất cả dòng hợp lệ.")

    disabled_next = n_bad > 0 or (~bad).sum() == 0
    if st.button("Tiếp theo: Phân bổ cây trồng →", type="primary", disabled=disabled_next):
        new_summary = build_crop_summary(st.session_state.detail_df)
        st.session_state.crop_df = merge_crop_alloc(new_summary, st.session_state.crop_df)
        st.session_state.step = 2
        st.rerun()


# ------------------------------------------------------------------
# STEP 2: CROP ALLOCATION
# ------------------------------------------------------------------
def step2_screen():
    st.markdown("##### 🌾 Phân bổ tỷ lệ % theo cây trồng")
    st.caption(
        "Tỉnh / Sản phẩm / Tháng / Kg-L / Target VNĐ được tổng hợp tự động từ kế hoạch ở Bước 1. "
        "**Kg/L = Số lượng (thùng) × Carton weight**. Điền % vào 6 cột có tiền tố 🌾 bên dưới — "
        "**mỗi dòng bắt buộc tổng % = 100** (cột 'Tổng %' tự cộng và tô màu, không cần tính tay)."
    )

    if st.session_state.crop_df is None or st.session_state.crop_df.empty:
        st.info("Chưa có dữ liệu — quay lại Bước 1 và đảm bảo các dòng có Tỉnh và số lượng > 0.")
        if st.button("← Quay lại Bước 1"):
            st.session_state.step = 1
            st.rerun()
        return

    df = st.session_state.crop_df.copy()

    column_config = {
        "Tỉnh": st.column_config.TextColumn("Tỉnh", disabled=True, width="small"),
        "Short name": st.column_config.TextColumn("Sản phẩm", disabled=True, width="medium"),
        "MonthLabel": st.column_config.TextColumn("Tháng", disabled=True, width="small"),
        "Kg/L": st.column_config.NumberColumn("Kg/L", format="%,.1f", disabled=True, width="small"),
        "Target VNĐ": st.column_config.NumberColumn("Target VNĐ", format="%,d", disabled=True, width="small"),
        "Tổng %": st.column_config.NumberColumn("Tổng %", format="%.0f%%", disabled=True, width="small"),
    }
    for c in CROPS:
        column_config[c] = st.column_config.NumberColumn(
            CROP_LABELS[c], min_value=0, max_value=100, step=1, format="%d", width="small"
        )

    show_cols = ["Tỉnh", "Short name", "MonthLabel", "Kg/L", "Target VNĐ"] + CROPS + ["Tổng %"]

    # Tô màu Tổng % (xanh nếu = 100, đỏ nếu khác 100). GIỚI HẠN KỸ THUẬT của Streamlit:
    # chỉ cột đã KHOÁ (disabled) mới tô nền theo giá trị được — 6 cột % cây trồng đang cho
    # sửa nên KHÔNG tô nền theo giá trị được (Streamlit chưa hỗ trợ), nên chỉ tô "Tổng %".
    def highlight_total_pct(s):
        colors = []
        for v in s:
            if 99.99 <= v <= 100.01:
                colors.append("background-color:#d9ead3; color:#274e13; font-weight:700;")
            else:
                colors.append("background-color:#f4cccc; color:#990000; font-weight:700;")
        return colors

    styled_source = df[show_cols].style.apply(highlight_total_pct, subset=["Tổng %"])

    n_rows = len(df)
    fit_height = min(560, 46 + 36 * n_rows)  # scale theo số dòng, tránh phải cuộn khi ít dòng

    edited = st.data_editor(
        styled_source,
        column_config=column_config,
        num_rows="fixed",
        height=fit_height,
        key="crop_editor",
    )
    edited["Tổng %"] = edited[CROPS].sum(axis=1)

    # gắn lại cột sắp xếp tháng thật (ẩn) để lần merge/xuất sau vẫn đúng thứ tự
    full = edited.copy()
    full["_month_sort"] = df["_month_sort"].values
    st.session_state.crop_df = full

    complete_mask = (full["Tổng %"] - 100).abs() < 1e-6
    n_complete = int(complete_mask.sum())
    n_total = len(full)

    c1, c2 = st.columns([2, 5])
    c1.metric("Dòng đạt 100%", f"{n_complete}/{n_total}")
    with c2:
        st.write("")
        if n_complete < n_total:
            st.warning("Còn dòng chưa đủ 100% — xem cột 'Tổng %' tô đỏ ở bảng trên (không cần tự cộng tay).")
        else:
            st.success("✓ Toàn bộ dòng đạt 100%.")

    b1, b2 = st.columns([1, 1])
    with b1:
        if st.button("← Quay lại Bước 1"):
            st.session_state.step = 1
            st.rerun()
    with b2:
        ready = n_total > 0 and n_complete == n_total
        if st.button("Hoàn tất kế hoạch & Xuất Excel ✓", type="primary", disabled=not ready):
            try:
                overview_df = build_overview(st.session_state.detail_df)
                xls_bytes = export_excel(st.session_state.user, st.session_state.detail_df, overview_df, full)
                st.session_state["export_bytes"] = xls_bytes
                st.session_state["export_ready"] = True
            except Exception as e:
                st.session_state["export_ready"] = False
                st.error("Xuất Excel thất bại. Chi tiết lỗi thật (không bị ẩn) hiện bên dưới để báo lại:")
                st.exception(e)
                st.code(traceback.format_exc())

    if st.session_state.get("export_ready"):
        st.success("✓ Kế hoạch hợp lệ! Tải file Excel bên dưới và gửi lại cho quản lý.")
        fname = f"KeHoachBanHang_{st.session_state.user['code']}_{date.today().isoformat()}.xlsx"
        st.download_button(
            "⬇ Tải file Excel kế hoạch",
            data=st.session_state["export_bytes"],
            file_name=fname,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            type="primary",
        )


# ------------------------------------------------------------------
# MAIN
# ------------------------------------------------------------------
def main():
    if st.session_state.user is None:
        login_screen()
        return

    user = st.session_state.user
    top_l, top_r = st.columns([5, 1])
    with top_l:
        st.markdown(
            f"**NV:** `{user['code']}` &nbsp;|&nbsp; **Người phụ trách:** {user['name']} "
            f"&nbsp;|&nbsp; **Vùng:** `{user['region']}` &nbsp;|&nbsp; "
            f"**Bước:** {'① Nhập kế hoạch' if st.session_state.step == 1 else '② Phân bổ cây trồng'}"
        )
    with top_r:
        if st.button("Đăng xuất"):
            for k in ["user", "detail_df", "crop_df", "step", "export_ready", "export_bytes", "known_customers"]:
                st.session_state.pop(k, None)
            st.rerun()
    st.divider()

    if st.session_state.step == 1:
        step1_screen()
    else:
        step2_screen()


if __name__ == "__main__":
    main()
