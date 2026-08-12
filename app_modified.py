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
import os
import urllib.request
import urllib.error
import urllib.parse

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
    return f"{int(mo)}/{y}"

MONTH_LABELS = [month_label(m) for m in MONTHS]
QTY_COLS = MONTH_LABELS.copy()

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
# SUPABASE DRAFT STORAGE (server-side REST; no extra Python package needed)
# ------------------------------------------------------------------
SUPABASE_URL = st.secrets.get("SUPABASE_URL", os.getenv("SUPABASE_URL", "")).rstrip("/")
SUPABASE_SERVICE_ROLE_KEY = st.secrets.get(
    "SUPABASE_SERVICE_ROLE_KEY", os.getenv("SUPABASE_SERVICE_ROLE_KEY", "")
)
DRAFT_TABLE = "salesman_drafts"


def supabase_enabled() -> bool:
    return bool(SUPABASE_URL and SUPABASE_SERVICE_ROLE_KEY)


def _supabase_request(method: str, path: str, payload=None, query: str = ""):
    if not supabase_enabled():
        return None
    url = f"{SUPABASE_URL}/rest/v1/{path}"
    if query:
        url += f"?{query}"
    data = None
    headers = {
        "apikey": SUPABASE_SERVICE_ROLE_KEY,
        "Authorization": f"Bearer {SUPABASE_SERVICE_ROLE_KEY}",
        "Content-Type": "application/json",
        "Accept": "application/json",
    }
    if method.upper() == "POST":
        headers["Prefer"] = "resolution=merge-duplicates,return=representation"
    if payload is not None:
        data = json.dumps(payload, ensure_ascii=False, default=str).encode("utf-8")
    req = urllib.request.Request(url, data=data, headers=headers, method=method.upper())
    try:
        with urllib.request.urlopen(req, timeout=12) as resp:
            raw = resp.read().decode("utf-8")
            return json.loads(raw) if raw else []
    except urllib.error.HTTPError as exc:
        body = exc.read().decode("utf-8", errors="replace")
        raise RuntimeError(f"Supabase HTTP {exc.code}: {body[:500]}") from exc
    except Exception as exc:
        raise RuntimeError(f"Không kết nối được Supabase: {exc}") from exc


def _json_safe_df(df: pd.DataFrame) -> list:
    if df is None or df.empty:
        return []
    out = df.copy()
    out = out.drop(columns=[c for c in out.columns if str(c).startswith("_")], errors="ignore")
    out = out.where(pd.notna(out), None)
    return out.to_dict(orient="records")


def save_draft_to_supabase(user: dict, detail_df: pd.DataFrame, crop_df: pd.DataFrame, step: int):
    if not supabase_enabled():
        return False
    payload = {
        "employee_code": user["code"],
        "detail_data": _json_safe_df(detail_df),
        "crop_data": _json_safe_df(crop_df),
        "step": int(step),
        "updated_at": datetime.now().astimezone().isoformat(),
    }
    _supabase_request("POST", DRAFT_TABLE, payload)
    return True


def load_draft_from_supabase(user: dict):
    if not supabase_enabled():
        return None
    rows = _supabase_request(
        "GET", DRAFT_TABLE,
        query=f"employee_code=eq.{urllib.parse.quote(user['code'])}&select=employee_code,detail_data,crop_data,step,updated_at&limit=1",
    )
    if not rows:
        return None
    return rows[0]


def clear_draft_from_supabase(user: dict):
    if not supabase_enabled():
        return False
    _supabase_request("DELETE", DRAFT_TABLE, query=f"employee_code=eq.{urllib.parse.quote(user['code'])}")
    return True


def dataframe_from_draft(records, columns):
    if not records:
        return None
    df = pd.DataFrame(records)
    for c in columns:
        if c not in df.columns:
            df[c] = 0 if c in QTY_COLS or c in ["Đơn giá", "Thành tiền"] else ""
    return df[columns].copy()


def _draft_fingerprint(detail_df, crop_df, step):
    """Stable-enough fingerprint used to detect unsaved changes in this session."""
    payload = {
        "step": int(step),
        "detail": _json_safe_df(detail_df),
        "crop": _json_safe_df(crop_df),
    }
    return json.dumps(payload, ensure_ascii=False, sort_keys=True, default=str)


def mark_draft_dirty(detail_df=None, crop_df=None, step=None):
    """Mark current in-memory draft as changed without writing immediately."""
    detail_df = st.session_state.detail_df if detail_df is None else detail_df
    crop_df = st.session_state.crop_df if crop_df is None else crop_df
    step = st.session_state.step if step is None else step
    st.session_state.draft_dirty = True
    st.session_state.draft_fingerprint = _draft_fingerprint(detail_df, crop_df, step)


def _perform_autosave():
    """Save at most once per debounce window when the draft is dirty."""
    if not st.session_state.get("user") or not st.session_state.get("draft_dirty"):
        return
    if not supabase_enabled():
        return

    now = datetime.now().timestamp()
    last = float(st.session_state.get("draft_last_saved_ts", 0) or 0)
    debounce_seconds = 7
    if now - last < debounce_seconds:
        return

    try:
        save_draft_to_supabase(
            st.session_state.user,
            st.session_state.detail_df,
            st.session_state.crop_df,
            st.session_state.step,
        )
        st.session_state.draft_last_saved_ts = now
        st.session_state.draft_last_saved_at = datetime.now().astimezone().strftime("%d/%m/%Y %H:%M:%S")
        st.session_state.draft_dirty = False
        st.session_state.draft_saved_fingerprint = st.session_state.get("draft_fingerprint", "")
        st.session_state.draft_autosave_error = None
    except Exception as exc:
        # Do not interrupt typing if the network/database is temporarily unavailable.
        st.session_state.draft_autosave_error = str(exc)


@st.fragment(run_every="5s")
def autosave_fragment():
    """Small periodic fragment: gives true debounce behavior without rerunning the whole page."""
    if not st.session_state.get("user"):
        return
    _perform_autosave()
    if st.session_state.get("draft_last_saved_at"):
        st.caption(f"☁️ Tự động lưu gần nhất: {st.session_state['draft_last_saved_at']}")
    elif supabase_enabled():
        st.caption("☁️ Tự động lưu: đang chờ thay đổi…")
    if st.session_state.get("draft_autosave_error"):
        st.caption("⚠️ Autosave tạm thời không kết nối được; dữ liệu vẫn được giữ trên phiên hiện tại và sẽ thử lại.")

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
if "draft_dirty" not in st.session_state:
    st.session_state.draft_dirty = False
if "draft_last_saved_ts" not in st.session_state:
    st.session_state.draft_last_saved_ts = 0.0
if "draft_last_saved_at" not in st.session_state:
    st.session_state.draft_last_saved_at = None
if "draft_fingerprint" not in st.session_state:
    st.session_state.draft_fingerprint = ""
if "draft_saved_fingerprint" not in st.session_state:
    st.session_state.draft_saved_fingerprint = ""
if "draft_autosave_error" not in st.session_state:
    st.session_state.draft_autosave_error = None


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
    """Rows with no customer and no quantities are blank placeholders.
    Incomplete rows are allowed while drafting; they are simply excluded from summaries/export.
    """
    qty = df[QTY_COLS].apply(pd.to_numeric, errors="coerce").fillna(0).sum(axis=1)
    customer = df["Khách hàng (C1)"].fillna("").astype(str).str.strip()
    product = df["Sản phẩm"].fillna("").astype(str).str.strip()
    province = df["Tỉnh"].fillna("").astype(str).str.strip()
    return (customer == "") & (product == "") & (province == "") & (qty == 0)


def draft_has_plannable_rows(df: pd.DataFrame) -> bool:
    if df is None or df.empty:
        return False
    qty = df[QTY_COLS].apply(pd.to_numeric, errors="coerce").fillna(0).sum(axis=1)
    product = df["Sản phẩm"].fillna("").astype(str).str.strip()
    return bool(((product != "") & (qty > 0)).any())

def build_overview(df: pd.DataFrame) -> pd.DataFrame:
    valid = df[(df["Khách hàng (C1)"].fillna("").astype(str).str.strip() != "")]
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
    valid = df[df["Sản phẩm"].fillna("").astype(str).str.strip() != ""].copy()
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
        return f"{float(x):,.0f}"
    except Exception:
        return str(x)


# ------------------------------------------------------------------
# EXPORT TO EXCEL
# ------------------------------------------------------------------
def export_excel(user: dict, detail_df: pd.DataFrame, overview_df: pd.DataFrame, crop_df: pd.DataFrame) -> bytes:
    """Create the Excel workbook in memory with a robust engine fallback."""
    last_error = None
    for engine in ("openpyxl", "xlsxwriter"):
        try:
            buf = io.BytesIO()
            with pd.ExcelWriter(buf, engine=engine) as writer:
                pd.DataFrame([{
                    "Mã NV": user["code"], "Tên": user["name"], "Vùng": user["region"],
                    "Ngày xuất kế hoạch": datetime.now().strftime("%d/%m/%Y %H:%M"),
                }]).to_excel(writer, sheet_name="Thông tin", index=False)
                d = detail_df[~((detail_df["Khách hàng (C1)"].fillna("").astype(str).str.strip() == "") & (detail_df["Tỉnh"].fillna("").astype(str).str.strip() == "") & (detail_df["Sản phẩm"].fillna("").astype(str).str.strip() == "") & (detail_df[QTY_COLS].sum(axis=1) == 0))].copy()
                d = d.drop(columns=[c for c in d.columns if c.startswith("_")], errors="ignore")
                d.to_excel(writer, sheet_name="KH x Sản phẩm", index=False)
                overview_df.reset_index().to_excel(writer, sheet_name="Tổng quan KH", index=False)
                cdf = crop_df.copy()
                if not cdf.empty:
                    cdf["Month"] = pd.to_datetime(cdf["Month"])
                cdf[["Tỉnh", "Short name", "Month", "Kg/L", "Target VNĐ"] + CROPS].to_excel(writer, sheet_name="Cây trồng", index=False)
                if engine == "openpyxl":
                    for ws in writer.book.worksheets:
                        ws.freeze_panes = "A2"
                        for row in ws.iter_rows():
                            for cell in row:
                                if isinstance(cell.value, (int, float)):
                                    cell.number_format = "#,##0"
                        for col in ws.columns:
                            letter = col[0].column_letter
                            width = min(max(len(str(c.value or "")) for c in col) + 2, 45)
                            ws.column_dimensions[letter].width = width
                else:
                    money_fmt = writer.book.add_format({"num_format": "#,##0"})
                    for ws in writer.sheets.values():
                        ws.freeze_panes(1, 0); ws.set_column("A:Z", 14); ws.set_column("D:Z", 14, money_fmt)
            return buf.getvalue()
        except Exception as exc:
            last_error = exc
    raise RuntimeError(f"Không thể tạo file Excel: {last_error}")


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
            # Load only the latest draft for this employee.
            try:
                draft = load_draft_from_supabase(st.session_state.user)
                if draft:
                    detail = dataframe_from_draft(draft.get("detail_data"), DETAIL_COLS)
                    crop = dataframe_from_draft(draft.get("crop_data"), ["Tỉnh", "Short name", "Month", "MonthLabel", "Kg/L", "Target VNĐ"] + CROPS + ["Tổng %"])
                    if detail is not None and not detail.empty:
                        st.session_state.detail_df = recompute_detail(detail, row["region"])
                    if crop is not None:
                        st.session_state.crop_df = crop
                    st.session_state.step = int(draft.get("step") or 1)
                    st.session_state.draft_loaded_at = draft.get("updated_at")
                    st.session_state.draft_loaded = True
                    st.session_state.draft_dirty = False
                    st.session_state.draft_last_saved_at = draft.get("updated_at")
                    st.session_state.draft_last_saved_ts = datetime.now().timestamp()
                    st.session_state.draft_fingerprint = _draft_fingerprint(st.session_state.detail_df, st.session_state.crop_df, st.session_state.step)
                    st.session_state.draft_saved_fingerprint = st.session_state.draft_fingerprint
                    st.session_state.draft_autosave_error = None
            except Exception as exc:
                st.warning(f"Không tải được bản nháp từ Supabase: {exc}")
            st.rerun()
    st.markdown("---")
    st.caption("Ví dụ định dạng mã nhân viên: `E1234`.")


# ------------------------------------------------------------------
# STEP 1: DETAIL ENTRY
# ------------------------------------------------------------------
def step1_screen():
    user = st.session_state.user
    region = user["region"]

    if st.session_state.get("draft_loaded") and st.session_state.get("draft_loaded_at"):
        st.info(f"↩ Đã khôi phục bản nháp gần nhất: {st.session_state['draft_loaded_at']}")
        st.session_state.draft_loaded = False

    st.markdown("##### 📌 Tổng quan theo Khách hàng × Tháng")
    computed = recompute_detail(st.session_state.detail_df, region)
    overview_df = build_overview(computed)
    col_ov, _spacer = st.columns([1, 2])
    with col_ov:
        if overview_df.empty:
            st.info("Chưa có dữ liệu kế hoạch.")
        else:
            st.dataframe(overview_df.style.format(fmt_vnd), height=260, use_container_width=True)

    st.divider()
    st.markdown("##### 📝 Chi tiết: Khách hàng × Sản phẩm × Tháng")
    st.caption("Có thể nhập dần, chưa cần điền đủ thông tin ngay. Khách hàng và Tỉnh là ô nhập tự do; Sản phẩm có thể chọn từ danh sách theo vùng. Dữ liệu nhập sẽ được giữ lại qua các lần Streamlit rerun.")

    product_options = [""] + sorted(region_products(region)["name"].unique().tolist())
    province_options = [""] + region_provinces(region)

    # IMPORTANT: Customer is intentionally a free-text column now.
    # We do not rebuild a customer dropdown from the current dataframe, which was one
    # of the causes of editing friction/resetting when rows changed.
    column_config = {
        "Khách hàng (C1)": st.column_config.TextColumn("Khách hàng (C1)", width="medium"),
        "Tỉnh": st.column_config.SelectboxColumn("Tỉnh", options=province_options, width="small"),
        "Sản phẩm": st.column_config.SelectboxColumn("Sản phẩm", options=product_options, width="large"),
        "Đơn giá": st.column_config.NumberColumn("Đơn giá", format="%,.0f", disabled=True),
        "Thành tiền": st.column_config.NumberColumn("Thành tiền", format="%,.0f", disabled=True),
    }
    for c in QTY_COLS:
        column_config[c] = st.column_config.NumberColumn(c, min_value=0, step=1, format="%,d")

    # Use a stable editable dataframe as the widget source. Calculated columns are
    # refreshed only for display, never used to overwrite the user's editable state.
    base = st.session_state.detail_df.copy()
    if base.empty:
        base = pd.DataFrame([{"Khách hàng (C1)": "", "Tỉnh": "", "Sản phẩm": "", "Đơn giá": 0, **{c: 0 for c in QTY_COLS}, "Thành tiền": 0}])
    base = recompute_detail(base, region)

    edited = st.data_editor(
        base[DETAIL_COLS],
        column_config=column_config,
        num_rows="dynamic",
        use_container_width=True,
        height=420,
        key="detail_editor",
    )

    # Persist ONLY what the user actually edited. Do not write a recomputed dataframe
    # back into the widget source; that avoids the common 'typed value disappears' effect.
    st.session_state.detail_df = edited.copy()
    mark_draft_dirty(st.session_state.detail_df, st.session_state.crop_df, 1)
    computed = recompute_detail(st.session_state.detail_df, region)

    # Show computed values immediately below without mutating the editor state.
    bad_blank = invalid_mask(computed)
    grand_total = computed.loc[~bad_blank, "Thành tiền"].sum()
    n_plannable = int(draft_has_plannable_rows(computed))

    c1, c2, c3 = st.columns([2, 2, 3])
    c1.metric("Dòng đang nhập", f"{len(computed)}")
    c2.metric("Tổng kế hoạch", f"{fmt_vnd(grand_total)} đ")
    with c3:
        st.caption("Bạn có thể để trống một số ô và quay lại nhập tiếp sau.")

    b1, b2, b3 = st.columns([1, 1, 2])
    with b1:
        if st.button("💾 Lưu bản nháp", type="secondary", key="save_detail_draft"):
            try:
                # Save current editable data, plus any existing crop allocation.
                if save_draft_to_supabase(user, st.session_state.detail_df, st.session_state.crop_df, 1):
                    st.session_state.draft_dirty = False
                    st.session_state.draft_last_saved_ts = datetime.now().timestamp()
                    st.session_state.draft_last_saved_at = datetime.now().astimezone().strftime("%d/%m/%Y %H:%M:%S")
                    st.session_state.draft_saved_fingerprint = st.session_state.get("draft_fingerprint", "")
                    st.success("✓ Đã lưu bản nháp gần nhất.")
                else:
                    st.warning("Supabase chưa được cấu hình. Hãy thêm SUPABASE_URL và SUPABASE_SERVICE_ROLE_KEY vào Secrets.")
            except Exception as exc:
                st.error(f"Lưu bản nháp thất bại: {exc}")
    with b2:
        if st.button("🧹 Xóa bản nháp online", key="clear_draft"):
            try:
                if clear_draft_from_supabase(user):
                    st.session_state.detail_df = pd.DataFrame([{"Khách hàng (C1)": "", "Tỉnh": "", "Sản phẩm": "", "Đơn giá": 0, **{c: 0 for c in QTY_COLS}, "Thành tiền": 0}])
                    st.session_state.crop_df = None
                    st.session_state.step = 1
                    st.session_state.pop("detail_editor", None)
                    st.success("Đã xóa bản nháp online.")
                    st.rerun()
            except Exception as exc:
                st.error(f"Không thể xóa bản nháp: {exc}")
    with b3:
        st.caption("Bản nháp chỉ giữ bản mới nhất của từng mã nhân viên.")

    # Continue only requires at least one usable product+quantity row; unfinished rows
    # are allowed and remain as draft data.
    if st.button("Tiếp theo: Phân bổ cây trồng →", type="primary", disabled=not bool(n_plannable), key="next_to_crop"):
        new_summary = build_crop_summary(computed)
        st.session_state.crop_df = merge_crop_alloc(new_summary, st.session_state.crop_df)
        st.session_state.step = 2
        try:
            save_draft_to_supabase(user, st.session_state.detail_df, st.session_state.crop_df, 2)
        except Exception:
            pass
        st.rerun()


# ------------------------------------------------------------------
# STEP 2: CROP ALLOCATION
# ------------------------------------------------------------------
def step2_screen():
    st.markdown("##### 🌾 Phân bổ tỷ lệ % theo cây trồng")
    st.caption("Tỉnh / Sản phẩm / Tháng / Kg-L / Target VNĐ được tổng hợp tự động từ kế hoạch ở Bước 1. **Kg/L = Số lượng (thùng) × Carton weight**. Mỗi dòng bắt buộc tổng % = 100.")
    if st.session_state.crop_df is None or st.session_state.crop_df.empty:
        st.info("Chưa có dữ liệu — quay lại Bước 1 và đảm bảo các dòng có Tỉnh và số lượng > 0.")
        if st.button("← Quay lại Bước 1"):
            st.session_state.step = 1; st.rerun()
        return
    df = st.session_state.crop_df.copy()
    column_config = {
        "Tỉnh": st.column_config.TextColumn("Tỉnh", disabled=True),
        "Short name": st.column_config.TextColumn("Sản phẩm", disabled=True),
        "MonthLabel": st.column_config.TextColumn("Tháng", disabled=True),
        "Kg/L": st.column_config.NumberColumn("Kg/L", format="%,.1f", disabled=True),
        "Target VNĐ": st.column_config.NumberColumn("Target VNĐ", format="%,.0f", disabled=True),
        "Tổng %": st.column_config.NumberColumn("Tổng %", format="%.1f", disabled=True),
    }
    for c in CROPS:
        column_config[c] = st.column_config.NumberColumn(CROP_LABELS[c], min_value=0, max_value=100, step=0.5, format="%.1f")
    show_cols = ["Tỉnh", "Short name", "MonthLabel", "Kg/L", "Target VNĐ"] + CROPS + ["Tổng %"]
    edited = st.data_editor(df[show_cols], column_config=column_config, num_rows="fixed", use_container_width=True, height=min(520, max(220, 52 + len(df) * 35)), key="crop_editor")
    edited["Tổng %"] = edited[CROPS].sum(axis=1)
    full = edited.copy(); full["Month"] = df["Month"].values
    st.session_state.crop_df = full
    mark_draft_dirty(st.session_state.detail_df, st.session_state.crop_df, 2)
    # Keep the latest crop allocation in the in-memory draft; autosave persists it after the debounce window.
    complete_mask = (full["Tổng %"] - 100).abs() < 1e-6
    n_complete = int(complete_mask.sum()); n_total = len(full)

    # Highlight crop-entry headers and make the allocation table compact.
    # Columns: 1 Tỉnh, 2 Sản phẩm, 3 Tháng, 4 Kg/L, 5 Target,
    # 6-11 are the six crop columns, 12 is Tổng %.
    st.markdown("""<style>
    div[data-testid="stDataEditor"] [role="columnheader"]:nth-child(6),
    div[data-testid="stDataEditor"] [role="columnheader"]:nth-child(7),
    div[data-testid="stDataEditor"] [role="columnheader"]:nth-child(8),
    div[data-testid="stDataEditor"] [role="columnheader"]:nth-child(9),
    div[data-testid="stDataEditor"] [role="columnheader"]:nth-child(10),
    div[data-testid="stDataEditor"] [role="columnheader"]:nth-child(11) {
        background: #e8f3e8 !important;
        font-weight: 700 !important;
    }
    div[data-testid="stDataEditor"] [role="columnheader"]:nth-child(12) {
        background: #eef0f2 !important;
        font-weight: 700 !important;
    }
    div[data-testid="stDataEditor"] [role="gridcell"] {
        font-size: 12px !important;
    }
    </style>""", unsafe_allow_html=True)
    c1, c2, c3 = st.columns([2, 2, 3])
    c1.metric("Dòng đạt 100%", f"{n_complete}/{n_total}")
    c2.metric("Tổng Target VNĐ", fmt_vnd(full["Target VNĐ"].sum()))
    with c3:
        if n_complete < n_total: st.warning("⚠ Còn dòng chưa đủ 100%. Tổng % phải đạt 100% để xuất Excel.")
        else: st.success("✓ Toàn bộ dòng đạt 100%.")
    st.caption("🟩 Các cột cây trồng là nơi salesman nhập tỷ lệ. Ô Tổng % phải đạt 100%.")
    b1, b2 = st.columns([1, 1])
    with b1:
        if st.button("← Quay lại Bước 1"):
            st.session_state.step = 1; st.rerun()
    with b2:
        ready = n_total > 0 and n_complete == n_total
        if st.button("Hoàn tất kế hoạch & Xuất Excel ✓", type="primary", disabled=not ready):
            overview_df = build_overview(st.session_state.detail_df)
            try:
                try:
                    save_draft_to_supabase(st.session_state.user, st.session_state.detail_df, full, 2)
                except Exception:
                    pass
                st.session_state["export_bytes"] = export_excel(st.session_state.user, st.session_state.detail_df, overview_df, full)
                st.session_state["export_ready"] = True
            except Exception as exc:
                st.session_state["export_ready"] = False; st.error(f"Không thể tạo Excel: {exc}")
    if st.session_state.get("export_ready"):
        st.success("✓ Kế hoạch hợp lệ! Tải file Excel bên dưới và gửi lại cho quản lý.")
        fname = f"KeHoachBanHang_{st.session_state.user['code']}_{date.today().isoformat()}.xlsx"
        st.download_button("⬇ Tải file Excel kế hoạch", data=st.session_state["export_bytes"], file_name=fname, mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", type="primary")


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
            try:
                save_draft_to_supabase(user, st.session_state.detail_df, st.session_state.crop_df, st.session_state.step)
            except Exception:
                pass
            for k in ["user", "detail_df", "crop_df", "step", "export_ready", "export_bytes", "detail_editor", "crop_editor", "draft_loaded", "draft_loaded_at", "draft_dirty", "draft_last_saved_ts", "draft_last_saved_at", "draft_fingerprint", "draft_saved_fingerprint", "draft_autosave_error"]:
                st.session_state.pop(k, None)
            st.rerun()
    st.divider()
    if not supabase_enabled():
        st.caption("💾 Lưu bản nháp online: chưa cấu hình Supabase")
    else:
        st.caption("☁️ Supabase đã kết nối · Autosave sau khoảng 7 giây kể từ thay đổi cuối cùng")

    # Periodic 5-second fragment performs a 7-second debounce autosave without rerunning the editor.
    autosave_fragment()

    if st.session_state.step == 1:
        step1_screen()
    else:
        step2_screen()


if __name__ == "__main__":
    main()
