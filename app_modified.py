# -*- coding: utf-8 -*-
"""
Form Kế hoạch bán hàng - Salesman
==================================
Luồng:
  1) Đăng nhập bằng Mã NV -> tự hiện Tên + Vùng bán hàng
  2) Bước 1: nhập Khách hàng (C1) x Sản phẩm x Tháng (số lượng theo THÙNG).
     - Khách hàng & Sản phẩm bắt buộc.
     - Sản phẩm chỉ gợi ý theo đúng Vùng của salesman.
     - Tỉnh chọn theo Vùng, tự đồng bộ theo từng Khách hàng.
     - Bảng "Tổng quan theo Khách hàng x Tháng" tự tính, đặt ở đầu trang, thu nhỏ ~1/3 màn hình.
  3) Bước 2: phân bổ tỷ lệ % theo cây trồng cho từng (Tỉnh, Sản phẩm, Tháng) đã gom nhóm.
     - Kg/L = Số lượng (thùng) x Carton weight (kg/thùng)   <-- CÔNG THỨC BẮT BUỘC
     - Target VNĐ = Số lượng (thùng) x Đơn giá
     - Mỗi dòng bắt buộc tổng % = 100 mới hợp lệ.
  4) Xuất file Excel (nhiều sheet) để salesman tải về và gửi lại.

File này đọc dữ liệu tham chiếu (danh sách salesman / tỉnh / sản phẩm theo vùng)
từ "reference_data.json" nằm cùng thư mục.
"""

import json
import io
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
    y, mo = m.split("-")
    return f"{int(mo):02d}/{y}"

MONTH_LABELS = [month_label(m) for m in MONTHS]
QTY_COLS = [f"SL {lbl}" for lbl in MONTH_LABELS]

CROPS = ["Durian", "Coffee", "Rice", "Dragon fruit", "Mango", "Vegetable & others"]
CROP_LABELS = {
    "Durian": "Sầu riêng (%)",
    "Coffee": "Cà phê (%)",
    "Rice": "Lúa (%)",
    "Dragon fruit": "Thanh long (%)",
    "Mango": "Xoài (%)",
    "Vegetable & others": "Rau màu & khác (%)",
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

def recompute_detail(df: pd.DataFrame, region: str) -> pd.DataFrame:
    """Recompute Đơn giá / Thành tiền, and sync Tỉnh across rows sharing the same customer."""
    df = df.copy()
    for col in ["Khách hàng (C1)", "Tỉnh", "Sản phẩm"]:
        if col not in df.columns:
            df[col] = ""
    df["Khách hàng (C1)"] = df["Khách hàng (C1)"].fillna("").astype(str).str.strip()
    df["Sản phẩm"] = df["Sản phẩm"].fillna("").astype(str).str.strip()
    df["Tỉnh"] = df["Tỉnh"].fillna("").astype(str).str.strip()

    # auto price / carton weight from product
    prices, cartons = [], []
    for _, row in df.iterrows():
        info = product_lookup(region, row["Sản phẩm"]) if row["Sản phẩm"] else None
        prices.append(info["price"] if info else 0)
        cartons.append(info["carton"] if info else 0)
    df["Đơn giá"] = prices
    df["_carton_weight"] = cartons

    # sync province across rows of the same customer (first non-empty wins)
    prov_map = {}
    for _, row in df.iterrows():
        cust, prov = row["Khách hàng (C1)"], row["Tỉnh"]
        if cust and prov and cust not in prov_map:
            prov_map[cust] = prov
    df["Tỉnh"] = df.apply(
        lambda r: prov_map.get(r["Khách hàng (C1)"], r["Tỉnh"]) if r["Khách hàng (C1)"] else r["Tỉnh"],
        axis=1,
    )

    for c in QTY_COLS:
        if c not in df.columns:
            df[c] = 0
        df[c] = pd.to_numeric(df[c], errors="coerce").fillna(0)

    df["Thành tiền"] = df[QTY_COLS].sum(axis=1) * df["Đơn giá"]
    return df

def invalid_mask(df: pd.DataFrame):
    return (
        (df["Khách hàng (C1)"] == "")
        | (df["Tỉnh"] == "")
        | (df["Sản phẩm"] == "")
    )

def build_overview(df: pd.DataFrame) -> pd.DataFrame:
    valid = df[~invalid_mask(df) & (df["Khách hàng (C1)"] != "")]
    if valid.empty:
        return pd.DataFrame(columns=["Khách hàng"] + MONTH_LABELS)
    rows = []
    for cust, g in valid.groupby("Khách hàng (C1)"):
        row = {"Khách hàng": cust}
        for lbl, qcol in zip(MONTH_LABELS, QTY_COLS):
            row[lbl] = float((g[qcol] * g["Đơn giá"]).sum())
        rows.append(row)
    out = pd.DataFrame(rows).set_index("Khách hàng")
    out.loc["TỔNG"] = out.sum(numeric_only=True)
    return out

def build_crop_summary(df: pd.DataFrame) -> pd.DataFrame:
    """Group valid rows by (Tỉnh, Sản phẩm, Tháng). Kg/L = SL(thùng) x carton weight."""
    valid = df[~invalid_mask(df) & (df["Tỉnh"] != "")]
    recs = []
    for _, row in valid.iterrows():
        for lbl, qcol, m in zip(MONTH_LABELS, QTY_COLS, MONTHS):
            qty = row[qcol]
            if not qty:
                continue
            info = product_lookup(st.session_state.user["region"], row["Sản phẩm"])
            short = info["short"] if info else row["Sản phẩm"]
            carton_w = info["carton"] if info else 0
            price = info["price"] if info else 0
            recs.append({
                "Tỉnh": row["Tỉnh"],
                "Short name": short,
                "Month": m,
                "MonthLabel": lbl,
                "_qty": qty,
                "_carton_w": carton_w,
                "_price": price,
            })
    if not recs:
        return pd.DataFrame(columns=["Tỉnh", "Short name", "Month", "Kg/L", "Target VNĐ"] + CROPS + ["Tổng %"])

    raw = pd.DataFrame(recs)
    grouped = raw.groupby(["Tỉnh", "Short name", "Month", "MonthLabel"], as_index=False).apply(
        lambda g: pd.Series({
            "Kg/L": (g["_qty"] * g["_carton_w"]).sum(),          # <-- công thức bắt buộc: qty x carton weight
            "Target VNĐ": (g["_qty"] * g["_price"]).sum(),
        }),
        include_groups=False,
    ) if hasattr(pd.DataFrame, "groupby") else None

    # pandas >=2.2 requires include_groups kw only for apply; fallback for older pandas:
    if grouped is None or "Tỉnh" not in grouped.columns:
        grouped = raw.groupby(["Tỉnh", "Short name", "Month", "MonthLabel"]).apply(
            lambda g: pd.Series({
                "Kg/L": (g["_qty"] * g["_carton_w"]).sum(),
                "Target VNĐ": (g["_qty"] * g["_price"]).sum(),
            })
        ).reset_index()

    grouped = grouped.sort_values(["Tỉnh", "Short name", "Month"]).reset_index(drop=True)
    for c in CROPS:
        grouped[c] = 0.0
    grouped["Tổng %"] = 0.0
    return grouped

def merge_crop_alloc(new_summary: pd.DataFrame, old_alloc: pd.DataFrame) -> pd.DataFrame:
    """Preserve previously entered % values when the summary is rebuilt (e.g. user goes back and edits step 1)."""
    if old_alloc is None or old_alloc.empty:
        return new_summary
    key_cols = ["Tỉnh", "Short name", "Month"]
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
    """Xuất Excel ổn định, có fallback engine và định dạng số rõ ràng."""
    buf = io.BytesIO()

    # openpyxl là lựa chọn chính; nếu môi trường thiếu thì thử xlsxwriter.
    engine = None
    for candidate in ("openpyxl", "xlsxwriter"):
        try:
            __import__(candidate)
            engine = candidate
            break
        except ImportError:
            continue
    if engine is None:
        raise RuntimeError("Môi trường chưa cài openpyxl hoặc xlsxwriter. Hãy thêm openpyxl vào requirements.txt.")

    try:
        with pd.ExcelWriter(buf, engine=engine) as writer:
            info_df = pd.DataFrame([{
                "Mã NV": user["code"],
                "Tên": user["name"],
                "Vùng": user["region"],
                "Ngày xuất kế hoạch": datetime.now().strftime("%d/%m/%Y %H:%M"),
            }])
            info_df.to_excel(writer, sheet_name="Thông tin", index=False)

            d = detail_df[~invalid_mask(detail_df)].copy()
            d = d.drop(columns=[c for c in d.columns if c.startswith("_")], errors="ignore")
            d.to_excel(writer, sheet_name="KH x Sản phẩm", index=False)

            overview_df.reset_index().to_excel(writer, sheet_name="Tổng quan KH", index=False)

            cdf = crop_df.copy()
            if not cdf.empty and "Month" in cdf.columns:
                cdf["Month"] = pd.to_datetime(cdf["Month"], errors="coerce")
            out_cols = ["Tỉnh", "Short name", "Month", "Kg/L", "Target VNĐ"] + CROPS
            cdf[out_cols].to_excel(writer, sheet_name="Cây trồng", index=False)

            # Định dạng Excel.
            if engine == "openpyxl":
                for ws in writer.sheets.values():
                    ws.freeze_panes = "A2"
                    ws.auto_filter.ref = ws.dimensions
                    for column_cells in ws.columns:
                        max_len = max(len(str(cell.value or "")) for cell in column_cells)
                        ws.column_dimensions[column_cells[0].column_letter].width = min(max(max_len + 2, 10), 32)

                for sheet_name in ("KH x Sản phẩm", "Tổng quan KH"):
                    ws = writer.sheets[sheet_name]
                    for row in ws.iter_rows():
                        for cell in row:
                            if isinstance(cell.value, (int, float)) and cell.row > 1:
                                cell.number_format = '#,##0'

                ws = writer.sheets["Cây trồng"]
                for row in ws.iter_rows():
                    for cell in row:
                        if cell.column == 5 and cell.row > 1:
                            cell.number_format = '#,##0'
            else:
                # XlsxWriter dùng API riêng.
                money_fmt = writer.book.add_format({"num_format": "#,##0"})
                for sheet_name, df_out in [
                    ("KH x Sản phẩm", d),
                    ("Tổng quan KH", overview_df.reset_index()),
                    ("Cây trồng", cdf[out_cols]),
                ]:
                    ws = writer.sheets[sheet_name]
                    ws.freeze_panes(1, 0)
                    ws.autofilter(0, 0, len(df_out), max(len(df_out.columns) - 1, 0))
                    for idx, col in enumerate(df_out.columns):
                        width = min(max(len(str(col)) + 2, 10), 32)
                        ws.set_column(idx, idx, width)
                    if sheet_name in ("KH x Sản phẩm", "Tổng quan KH"):
                        ws.set_column(0, max(len(df_out.columns) - 1, 0), None, None)
                    if sheet_name == "Cây trồng" and "Target VNĐ" in df_out.columns:
                        idx = list(df_out.columns).index("Target VNĐ")
                        ws.set_column(idx, idx, 16, money_fmt)

        return buf.getvalue()
    except Exception as exc:
        raise RuntimeError(f"Không thể tạo file Excel: {type(exc).__name__}: {exc}") from exc


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
    st.markdown("---")
    st.caption("Mã ví dụ để demo: `E1234` (chỉ là mã minh họa, không phải mã nhân viên thực tế).")


# ------------------------------------------------------------------
# STEP 1: DETAIL ENTRY
# ------------------------------------------------------------------
def customer_suggestions(df: pd.DataFrame):
    """Các khách hàng đã nhập ở các dòng phía trên, dùng làm danh sách gợi ý."""
    if df is None or df.empty or "Khách hàng (C1)" not in df.columns:
        return [""]
    vals = (
        df["Khách hàng (C1)"]
        .fillna("")
        .astype(str)
        .str.strip()
    )
    return [""] + list(dict.fromkeys([v for v in vals if v]))


def commit_detail_editor():
    """Áp dụng delta của data_editor vào DataFrame hiện tại sau mỗi lần đổi ô."""
    try:
        state = st.session_state.get("detail_editor", {})
        base = st.session_state.detail_df.copy()
        edited_rows = state.get("edited_rows", {})
        added_rows = state.get("added_rows", [])
        deleted_rows = state.get("deleted_rows", [])

        for row_idx, changes in edited_rows.items():
            idx = int(row_idx)
            for col, value in changes.items():
                if col in base.columns and idx < len(base):
                    base.at[idx, col] = value

        if deleted_rows:
            base = base.drop(index=[int(i) for i in deleted_rows], errors="ignore").reset_index(drop=True)

        for row in added_rows:
            new_row = {col: 0 for col in base.columns}
            for col, value in row.items():
                if col in new_row:
                    new_row[col] = value
            base = pd.concat([base, pd.DataFrame([new_row])], ignore_index=True)

        st.session_state.detail_df = recompute_detail(
            base.reset_index(drop=True), st.session_state.user["region"]
        )
    except Exception:
        # Không làm hỏng form nếu Streamlit trả về trạng thái trung gian.
        pass

def step1_screen():
    user = st.session_state.user
    region = user["region"]

    # ---- Overview panel ----
    st.markdown("##### 📌 Tổng quan theo Khách hàng × Tháng &nbsp;·&nbsp; *tự tính từ toàn bộ bảng chi tiết*")
    overview_df = build_overview(st.session_state.detail_df)
    if overview_df.empty:
        st.info("Chưa có khách hàng nào — nhập ở bảng chi tiết bên dưới.")
    else:
        styled = overview_df.style.format(fmt_vnd)
        # Không giới hạn 1/3 màn hình; chiều cao bám theo số dòng thực tế.
        overview_height = max(120, 38 * (len(overview_df) + 1) + 10)
        st.dataframe(styled, height=overview_height, use_container_width=True)

    st.divider()

    # ---- Detail editor ----
    st.markdown("##### 📝 Chi tiết: Khách hàng × Sản phẩm × Tháng")
    st.caption(
        "Bắt buộc: **Khách hàng, Tỉnh, Sản phẩm**. Đơn giá tự lấy theo sản phẩm (vùng "
        f"**{region}**). Số lượng nhập theo **thùng**. Khi đổi ô, dữ liệu được chốt ngay; "
        "không cần bấm Enter."
    )

    product_options = [""] + sorted(region_products(region)["name"].unique().tolist())
    province_options = [""] + region_provinces(region)
    customer_options = customer_suggestions(st.session_state.detail_df)

    # Streamlit data_editor không hỗ trợ autocomplete/datalist trực tiếp cho TextColumn.
    # Vì vậy giữ ô khách hàng là TextColumn để vẫn nhập được khách hàng mới, đồng thời
    # có ô gợi ý phía trên để chọn nhanh khách hàng đã nhập trước đó.
    if len(customer_options) > 1:
        st.caption("💡 Gợi ý khách hàng đã nhập trước đó:")
        suggestion = st.selectbox(
            "Chọn khách hàng để tham khảo",
            customer_options[1:],
            index=None,
            placeholder="Gõ vài ký tự để tìm khách hàng...",
            label_visibility="collapsed",
            key="customer_suggestion",
        )
        if suggestion:
            st.info(f"Khách hàng đã có: **{suggestion}** — có thể copy tên này vào dòng mới; tỉnh sẽ tự đồng bộ.")

    column_config = {
        "Khách hàng (C1)": st.column_config.TextColumn(
            "Khách hàng (C1)", required=True, width="medium",
            help="Nhập tự do. Các khách hàng đã dùng trước đó được gợi ý ở ô phía trên."
        ),
        "Tỉnh": st.column_config.SelectboxColumn("Tỉnh", options=province_options, required=True, width="small"),
        "Sản phẩm": st.column_config.SelectboxColumn(
            "Sản phẩm", options=product_options, required=True, width="large",
            help="Gõ để lọc nhanh trong danh sách sản phẩm của vùng bạn.",
        ),
        "Đơn giá": st.column_config.NumberColumn("Đơn giá", format="%d", disabled=True),
        "Thành tiền": st.column_config.NumberColumn("Thành tiền", format="%d", disabled=True),
    }
    for c in QTY_COLS:
        column_config[c] = st.column_config.NumberColumn(c, min_value=0, step=1, format="%d")

    display_cols = DETAIL_COLS
    source_for_editor = st.session_state.detail_df[display_cols].copy()
    edited = st.data_editor(
        source_for_editor,
        column_config=column_config,
        num_rows="dynamic",
        use_container_width=True,
        height=max(320, min(720, 38 * (len(source_for_editor) + 3))),
        key="detail_editor",
    )

    # Commit ngay khi Streamlit trả về dữ liệu sau thao tác đổi ô (không cần Enter).
    if not edited.equals(source_for_editor):
        st.session_state.detail_df = recompute_detail(edited, region)

    recomputed = st.session_state.detail_df

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
        "**Kg/L = Số lượng (thùng) × Carton weight**. Mỗi dòng bắt buộc tổng % = 100."
    )

    if st.session_state.crop_df is None or st.session_state.crop_df.empty:
        st.info("Chưa có dữ liệu — quay lại Bước 1 và đảm bảo các dòng có Tỉnh và số lượng > 0.")
        if st.button("← Quay lại Bước 1"):
            st.session_state.step = 1
            st.rerun()
        return

    df = st.session_state.crop_df.copy()

    column_config = {
        "Tỉnh": st.column_config.TextColumn("Tỉnh", disabled=True),
        "Short name": st.column_config.TextColumn("Sản phẩm", disabled=True),
        "MonthLabel": st.column_config.TextColumn("Tháng", disabled=True),
        "Kg/L": st.column_config.NumberColumn("Kg/L", format="%.1f", disabled=True),
        "Target VNĐ": st.column_config.NumberColumn("Target VNĐ", format="#,##0", disabled=True),
        "Tổng %": st.column_config.NumberColumn("Tổng %", format="%.1f", disabled=True),
    }
    for c in CROPS:
        column_config[c] = st.column_config.NumberColumn(
            ("⭐ " if c == "Durian" else "☕ " if c == "Coffee" else "") + CROP_LABELS[c],
            min_value=0, max_value=100, step=0.5, format="%.1f",
            help="Ưu tiên nhập cột này." if c in ("Durian", "Coffee") else None,
        )

    show_cols = ["Tỉnh", "Short name", "MonthLabel", "Kg/L", "Target VNĐ"] + CROPS + ["Tổng %"]

    # Tô nổi bật Sầu riêng + Cà phê để salesman biết hai cột cần chú ý.
    st.markdown(
        '<div style="padding:8px 12px;border-radius:8px;background:#fff7d6;border:1px solid #f0c36d;">'
        '👉 <b>Ưu tiên nhập tỷ lệ vào Sầu riêng (%) và Cà phê (%)</b>. Các cây trồng khác vẫn có thể nhập khi cần.'
        '</div>',
        unsafe_allow_html=True,
    )

    edited = st.data_editor(
        df[show_cols],
        column_config=column_config,
        num_rows="fixed",
        use_container_width=True,
        # Chiều cao bám số dòng, tránh scrollbar riêng trong bảng.
        height=max(180, 40 * (len(df) + 2)),
        key="crop_editor",
    )
    edited["Tổng %"] = edited[CROPS].sum(axis=1)

    # merge back the hidden 'Month' (real date) column for export
    full = edited.copy()
    full["Month"] = df["Month"].values
    st.session_state.crop_df = full

    complete_mask = (full["Tổng %"] - 100).abs() < 1e-6
    n_complete = int(complete_mask.sum())
    n_total = len(full)

    # Không tạo thêm bảng preview. Chỉ hiển thị trạng thái tổng % ngay dưới bảng.
    if n_complete == n_total:
        st.success(f"✓ Tổng tỷ lệ: tất cả {n_total} dòng đều = 100%.")
    else:
        st.error(f"✗ Tổng tỷ lệ: còn {n_total - n_complete}/{n_total} dòng chưa = 100%.")


    c1, c2, c3 = st.columns([2, 3, 3])
    c1.metric("Dòng đạt 100%", f"{n_complete}/{n_total}")
    with c2:
        st.write("")
        if n_complete < n_total:
            st.warning("Còn dòng chưa đủ 100% — sửa % ở bảng trên (không cần tự cộng tay).")
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
            overview_df = build_overview(st.session_state.detail_df)
            try:
                xls_bytes = export_excel(
                    st.session_state.user,
                    st.session_state.detail_df,
                    overview_df,
                    full,
                )
                st.session_state["export_bytes"] = xls_bytes
                st.session_state["export_ready"] = True
                st.rerun()
            except Exception as exc:
                st.error(f"Không thể tạo file Excel: {exc}")

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
            for k in ["user", "detail_df", "crop_df", "step", "export_ready", "export_bytes"]:
                st.session_state.pop(k, None)
            st.rerun()
    st.divider()

    if st.session_state.step == 1:
        step1_screen()
    else:
        step2_screen()


if __name__ == "__main__":
    main()
