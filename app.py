import os
import re
import uuid
from datetime import date, timedelta
from io import BytesIO
import pandas as pd
import psycopg2
from psycopg2.extras import RealDictCursor
import streamlit as st

# --- PAGE CONFIG ---
st.set_page_config(
    page_title="KK & Company | UV Print & Stock Manager",
    page_icon="🖨️",
    layout="wide"
)

# --- AUTHENTICATION ---
def check_password():
    if "authenticated" not in st.session_state:
        st.session_state.authenticated = False
    if st.session_state.authenticated:
        return True

    st.markdown("### 🖨️ KK & Company — UV Print & Stock Account Manager")
    pwd = st.text_input("सॉफ्टवेयर पासवर्ड डालें (Enter Password):", type="password")
    if st.button("Login", type="primary"):
        expected = st.secrets.get("APP_PASSWORD", os.environ.get("APP_PASSWORD", "kk1234"))
        if pwd == expected:
            st.session_state.authenticated = True
            st.rerun()
        else:
            st.error("गलत पासवर्ड! कृपया सही पासवर्ड डालें।")
    return False

if not check_password():
    st.stop()

from urllib.parse import urlsplit, unquote

def get_conn():
    return psycopg2.connect(
        host=st.secrets["DB_HOST"],
        port=st.secrets["DB_PORT"],
        user=st.secrets["DB_USER"],
        password=st.secrets["DB_PASSWORD"],
        dbname=st.secrets["DB_NAME"],
        sslmode="require",
        cursor_factory=RealDictCursor,
        connect_timeout=15
    )

def normalize_model(name: str) -> str:
    if not name or pd.isna(name):
        return ""
    # Removes spaces, hyphens, underscores and converts to lowercase for 100% matching
    return re.sub(r'[^a-zA-Z0-9]', '', str(name)).lower()

def extract_sku_model_and_prefix(sku: str):
    s = str(sku).strip()
    if "-" in s:
        parts = s.split("-", 1)
        model_part = parts[0].strip()
        suffix = parts[1].strip()
        prefix_match = re.match(r'^([A-Za-z]+)', suffix)
        prefix = prefix_match.group(1).upper() if prefix_match else ""
        return model_part, prefix
    return s, ""

def to_excel_bytes(sheets_dict: dict) -> bytes:
    output = BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        for sheet_name, df in sheets_dict.items():
            df.to_excel(writer, index=False, sheet_name=sheet_name[:31])
    return output.getvalue()

# --- LOAD MASTER DATA ---
def fetch_sellers():
    with get_conn() as conn, conn.cursor() as cur:
        cur.execute("SELECT * FROM sellers ORDER BY name;")
        return cur.fetchall()

def fetch_stock():
    with get_conn() as conn, conn.cursor() as cur:
        cur.execute("SELECT * FROM stock_master ORDER BY display_model;")
        return pd.DataFrame(cur.fetchall())

# --- SIDEBAR NAVIGATION ---
st.sidebar.title("🖨️ KK Print Manager")
menu = st.sidebar.radio(
    "मेन्यू चुनें:",
    [
        "📊 Dashboard & True Costing",
        "📥 Stock Add (Inward Excel)",
        "📤 Stock Reduce (Dispatch Excel)",
        "⚠️ Daily Loss & Wastage",
        "💸 Ink & Fixed Expenses",
        "🚨 10-Day Auto Order Sheet",
        "📒 Seller Ledger & Undo",
        "⚙️ Seller & 3-Category Rates"
    ]
)

# ==============================================================================
# TAB 1: DASHBOARD & TRUE COSTING
# ==============================================================================
if menu == "📊 Dashboard & True Costing":
    st.header("📊 लाइव डैशबोर्ड और असली लागत (True Costing)")

    col_d1, col_d2 = st.columns(2)
    start_date = col_d1.date_input("शुरुआती तारीख (From)", value=date.today().replace(day=1))
    end_date = col_d2.date_input("अंतिम तारीख (To)", value=date.today())

    stock_df = fetch_stock()
    with get_conn() as conn, conn.cursor() as cur:
        cur.execute("""
            SELECT tx_type, COALESCE(SUM(qty),0) as total_qty, COALESCE(SUM(total_amount),0) as total_amt
            FROM transactions
            WHERE tx_date BETWEEN %s AND %s
            GROUP BY tx_type;
        """, (start_date, end_date))
        tx_summary = {r["tx_type"]: r for r in cur.fetchall()}

        cur.execute("""
            SELECT category, COALESCE(SUM(amount),0) as total_exp
            FROM expenses
            WHERE exp_date BETWEEN %s AND %s
            GROUP BY category;
        """, (start_date, end_date))
        exp_rows = cur.fetchall()

    total_blank_qty = int(stock_df["current_qty"].sum()) if not stock_df.empty else 0
    total_blank_val = float((stock_df["current_qty"] * stock_df["avg_rate"]).sum()) if not stock_df.empty else 0.0

    ok_printed = int(tx_summary.get("DISPATCH", {}).get("total_qty", 0))
    ok_cover_cost = float(tx_summary.get("DISPATCH", {}).get("total_amt", 0))

    wrong_qty = int(tx_summary.get("WRONG_PRINT", {}).get("total_qty", 0))
    wrong_cover_loss = float(tx_summary.get("WRONG_PRINT", {}).get("total_amt", 0))

    wastage_qty = int(tx_summary.get("WASTAGE", {}).get("total_qty", 0))
    wastage_cover_loss = float(tx_summary.get("WASTAGE", {}).get("total_amt", 0))

    pending_empty_prints = int(tx_summary.get("PENDING_PRINT", {}).get("total_qty", 0))

    total_expenses = sum(float(r["total_exp"]) for r in exp_rows)
    ink_exp = sum(float(r["total_exp"]) for r in exp_rows if "Ink" in r["category"])
    elec_exp = sum(float(r["total_exp"]) for r in exp_rows if "Electricity" in r["category"])

    # Total machine impressions (OK + Wrong + Wastage + Pending Empty Bed)
    total_impressions = ok_printed + wrong_qty + wastage_qty + pending_empty_prints
    ink_elec_per_impression = ((ink_exp + elec_exp) / total_impressions) if total_impressions > 0 else 0.0

    # Pending empty bed loss = only ink + electricity wasted on empty slots
    pending_loss_rs = pending_empty_prints * ink_elec_per_impression
    total_loss_rs = wrong_cover_loss + wastage_cover_loss + pending_loss_rs

    # True landing cost per OK cover sold/dispatched
    true_cost_per_ok = (
        (ok_cover_cost + total_expenses + wrong_cover_loss + wastage_cover_loss) / ok_printed
        if ok_printed > 0 else 0.0
    )

    m1, m2, m3, m4 = st.columns(4)
    m1.metric("📦 गोदाम में कुल स्टॉक", f"{total_blank_qty:,} पीस", f"₹{total_blank_val:,.0f} वैल्यू")
    m2.metric("✅ सही प्रिंट (OK Covers)", f"{ok_printed:,} पीस", f"कवर लागत: ₹{ok_cover_cost:,.0f}")
    m3.metric("⚠️ कुल नुकसान (Loss)", f"₹{total_loss_rs:,.0f}", f"Wastage:{wastage_qty} | Wrong:{wrong_qty} | Pending:{pending_empty_prints}")
    m4.metric("🎯 1 सही कवर की असली लागत", f"₹{true_cost_per_ok:,.2f}", f"कुल खर्चे: ₹{total_expenses:,.0f}")

    st.divider()
    st.subheader("📦 लाइव स्टॉक लिस्ट (Low Stock लाल रंग में)")
    if not stock_df.empty:
        search_q = st.text_input("मॉडल सर्च करें (Search Model):")
        view_df = stock_df.copy()
        if search_q:
            view_df = view_df[view_df["display_model"].str.contains(search_q, case=False, na=False)]
        view_df["Total Value (₹)"] = (view_df["current_qty"] * view_df["avg_rate"]).round(2)
        view_df["Status"] = view_df.apply(
            lambda r: "🔴 LOW STOCK" if r["current_qty"] <= r["low_stock_limit"] else "🟢 OK", axis=1
        )
        st.dataframe(
            view_df[["Status", "display_model", "current_qty", "avg_rate", "Total Value (₹)", "default_seller", "low_stock_limit"]],
            use_container_width=True,
            hide_index=True
        )
    else:
        st.info("अभी स्टॉक खाली है। 'Stock Add' टैब से पहली Excel अपलोड करें।")

# ==============================================================================
# TAB 2: STOCK ADD (INWARD EXCEL WITH 3-CATEGORY RATES)
# ==============================================================================
elif menu == "📥 Stock Add (Inward Excel)":
    st.header("📥 नया स्टॉक जोड़ें (Purchase / Inward Excel)")
    sellers = fetch_sellers()
    if not sellers:
        st.warning("पहले 'Seller & 3-Category Rates' टैब में जाकर कम से कम एक सेलर और उसकी 3 कैटेगरी के रेट सेव करें।")
        st.stop()

    seller_map = {s["name"]: s for s in sellers}
    col1, col2 = st.columns(2)
    sel_seller_name = col1.selectbox("सेलर चुनें (Select Seller):", list(seller_map.keys()))
    inward_date = col2.date_input("खरीद की तारीख (Purchase Date):", value=date.today())

    s_info = seller_map[sel_seller_name]
    cat_options = {
        f"Cat 1: {s_info['cat1_name']} (₹{s_info['cat1_rate']})": (1, float(s_info["cat1_rate"])),
        f"Cat 2: {s_info['cat2_name']} (₹{s_info['cat2_rate']})": (2, float(s_info["cat2_rate"])),
        f"Cat 3: {s_info['cat3_name']} (₹{s_info['cat3_rate']})": (3, float(s_info["cat3_rate"])),
    }
    default_sheet_cat = st.selectbox(
        "इस पूरी शीट की डिफ़ॉल्ट रेट कैटेगरी चुनें (पुराने मॉडल्स पर उनका सेव किया हुआ रेट अपने आप लगेगा):",
        list(cat_options.keys())
    )

    uploaded_file = st.file_uploader("सेलर की Stock Add Excel अपलोड करें (Col A = Model Name, Col B = Quantity)", type=["xlsx", "xls"])
    if uploaded_file:
        xl = pd.ExcelFile(uploaded_file)
        sheet_name = st.selectbox("Excel शीट चुनें:", xl.sheet_names)
        raw_df = xl.parse(sheet_name)

        # Read Col A (index 0) and Col B (index 1)
        df_clean = pd.DataFrame({
            "Model Name": raw_df.iloc[:, 0].astype(str).str.strip(),
            "Quantity": pd.to_numeric(raw_df.iloc[:, 1], errors="coerce").fillna(0).astype(int)
        })
        df_clean = df_clean[(df_clean["Model Name"] != "") & (df_clean["Model Name"].str.lower() != "nan") & (df_clean["Quantity"] > 0)]

        # Match existing models to auto-fill saved category/rate
        stock_df = fetch_stock()
        existing_map = {}
        if not stock_df.empty:
            for _, r in stock_df.iterrows():
                existing_map[r["norm_model"]] = r

        def_cat_num, def_rate = cat_options[default_sheet_cat]
        preview_rows = []
        for _, row in df_clean.iterrows():
            m_name = row["Model Name"]
            qty = int(row["Quantity"])
            norm = normalize_model(m_name)
            if norm in existing_map:
                saved = existing_map[norm]
                c_num = int(saved["default_cat"] or def_cat_num)
                rate = float(s_info[f"cat{c_num}_rate"]) if float(s_info[f"cat{c_num}_rate"]) > 0 else float(saved["avg_rate"])
                status = "Saved Model"
            else:
                c_num = def_cat_num
                rate = def_rate
                status = "New Model"
            preview_rows.append({
                "Model Name": m_name,
                "Quantity": qty,
                "Category (1/2/3)": c_num,
                "Rate (₹)": rate,
                "Status": status
            })

        edit_df = st.data_editor(
            pd.DataFrame(preview_rows),
            use_container_width=True,
            num_rows="fixed",
            column_config={
                "Category (1/2/3)": st.column_config.SelectboxColumn("Category (1/2/3)", options=[1, 2, 3], required=True),
                "Rate (₹)": st.column_config.NumberColumn("Rate (₹)", min_value=0.0, step=0.5)
            }
        )

        # Recalculate rate if user changed category number in editor
        for idx, r in edit_df.iterrows():
            c_n = int(r["Category (1/2/3)"])
            cat_r = float(s_info[f"cat{c_n}_rate"])
            if cat_r > 0 and r["Rate (₹)"] == def_rate and c_n != def_cat_num:
                edit_df.at[idx, "Rate (₹)"] = cat_r

        total_pieces = int(edit_df["Quantity"].sum())
        total_bill = float((edit_df["Quantity"] * edit_df["Rate (₹)"]).sum())
        st.info(f"*कुल पीस:* {total_pieces:,} | *कुल खरीद बिल:* ₹{total_bill:,.2f}")

        if st.button("✅ स्टॉक और सेलर बिल सेव करें", type="primary"):
            batch_id = f"IN-{uuid.uuid4().hex[:8].upper()}"
            with get_conn() as conn, conn.cursor() as cur:
                for _, r in edit_df.iterrows():
                    m_name = str(r["Model Name"]).strip()
                    norm = normalize_model(m_name)
                    qty = int(r["Quantity"])
                    rate = float(r["Rate (₹)"])
                    c_n = int(r["Category (1/2/3)"])
                    amt = qty * rate

                    cur.execute("SELECT current_qty, avg_rate FROM stock_master WHERE norm_model = %s;", (norm,))
                    ex = cur.fetchone()
                    if ex:
                        old_q = int(ex["current_qty"])
                        old_r = float(ex["avg_rate"])
                        new_q = old_q + qty
                        new_avg = ((old_q * old_r) + (qty * rate)) / new_q if new_q > 0 else rate
                        cur.execute("""
                            UPDATE stock_master
                            SET current_qty = %s, avg_rate = %s, default_seller = %s, default_cat = %s
                            WHERE norm_model = %s;
                        """, (new_q, round(new_avg, 2), sel_seller_name, c_n, norm))
                    else:
                        cur.execute("""
                            INSERT INTO stock_master (norm_model, display_model, current_qty, avg_rate, default_seller, default_cat)
                            VALUES (%s, %s, %s, %s, %s, %s);
                        """, (norm, m_name, qty, rate, sel_seller_name, c_n))

                    cur.execute("""
                        INSERT INTO transactions (batch_id, tx_date, tx_type, seller_name, norm_model, display_model, qty, unit_rate, total_amount)
                        VALUES (%s, %s, 'INWARD', %s, %s, %s, %s, %s, %s);
                    """, (batch_id, inward_date, sel_seller_name, norm, m_name, qty, rate, amt))
                conn.commit()
            st.success(f"सफलतापूर्वक {total_pieces} कवर स्टॉक में जुड़ गए! (Batch ID: {batch_id})")

# ==============================================================================
# TAB 3: STOCK REDUCE (SELLER DISPATCH / PRINT EXCEL)
# ==============================================================================
elif menu == "📤 Stock Reduce (Dispatch Excel)":
    st.header("📤 प्रिंट / डिस्पैच स्टॉक घटाएं (Seller-Wise SKU Excel)")
    sellers = fetch_sellers()
    seller_names = [s["name"] for s in sellers] if sellers else ["General Seller"]

    col1, col2 = st.columns(2)
    dispatch_date = col1.date_input("प्रिंट / डिस्पैच तारीख:", value=date.today())
    selected_seller = col2.selectbox("किस सेलर / अकाउंट को माल गया? (Seller Name):", seller_names + ["+ नया सेलर नाम लिखें"])
    if selected_seller == "+ नया सेलर नाम लिखें":
        selected_seller = st.text_input("नया सेलर / अकाउंट नाम लिखें (जैसे: VK, KK & Company, Radhika):")

    uploaded_files = st.file_uploader(
        "प्रिंट Excel फाइलें अपलोड करें (Col A = SKU जैसे 'Vivov30Pro5G-VK449', हर Row = 1 Qty)",
        type=["xlsx", "xls"],
        accept_multiple_files=True
    )

    if uploaded_files and selected_seller:
        all_items = []
        for f in uploaded_files:
            df = pd.read_excel(f)
            sku_series = df.iloc[:, 0].dropna().astype(str)
            for sku in sku_series:
                if sku.strip().lower() in ["sku", "model", "model name", "nan", ""]:
                    continue
                m_name, _ = extract_sku_model_and_prefix(sku)
                all_items.append({"Raw Model": m_name, "Norm": normalize_model(m_name)})

        if all_items:
            grp_df = pd.DataFrame(all_items).groupby(["Norm", "Raw Model"]).size().reset_index(name="Print Qty")
            # Merge duplicate norms with different casing
            grp_df = grp_df.groupby("Norm", as_index=False).agg({"Raw Model": "first", "Print Qty": "sum"})

            stock_df = fetch_stock()
            stock_map = {r["norm_model"]: r for _, r in stock_df.iterrows()} if not stock_df.empty else {}

            summary_rows = []
            for _, r in grp_df.iterrows():
                norm = r["Norm"]
                p_qty = int(r["Print Qty"])
                if norm in stock_map:
                    avail = int(stock_map[norm]["current_qty"])
                    rate = float(stock_map[norm]["avg_rate"])
                    disp_name = stock_map[norm]["display_model"]
                    status = "✅ OK" if avail >= p_qty else f"⚠️ कम स्टॉक ({avail})"
                else:
                    avail = 0
                    rate = 0.0
                    disp_name = r["Raw Model"]
                    status = "❓ स्टॉक में नहीं मिला"
                summary_rows.append({
                    "Model Name": disp_name,
                    "Norm": norm,
                    "Available Stock": avail,
                    "Reduce Qty": p_qty,
                    "After Reduce": avail - p_qty,
                    "Cover Rate (₹)": rate,
                    "Status": status
                })

            res_df = st.data_editor(pd.DataFrame(summary_rows), use_container_width=True, hide_index=True)
            total_red = int(res_df["Reduce Qty"].sum())
            st.info(f"*सेलर:* {selected_seller} | *कुल प्रिंट कवर घटेंगे:* {total_red:,} पीस")

            if st.button("🚀 स्टॉक घटाएं और सेलर लेजर में जोड़ें", type="primary"):
                batch_id = f"OUT-{uuid.uuid4().hex[:8].upper()}"
                with get_conn() as conn, conn.cursor() as cur:
                    for _, r in res_df.iterrows():
                        norm = r["Norm"]
                        disp = r["Model Name"]
                        qty = int(r["Reduce Qty"])
                        rate = float(r["Cover Rate (₹)"])
                        amt = qty * rate

                        cur.execute("SELECT current_qty FROM stock_master WHERE norm_model = %s;", (norm,))
                        ex = cur.fetchone()
                        if ex:
                            cur.execute("UPDATE stock_master SET current_qty = current_qty - %s WHERE norm_model = %s;", (qty, norm))
                        else:
                            cur.execute("""
                                INSERT INTO stock_master (norm_model, display_model, current_qty, avg_rate, default_seller)
                                VALUES (%s, %s, %s, 0, %s);
                            """, (norm, disp, -qty, selected_seller))

                        cur.execute("""
                            INSERT INTO transactions (batch_id, tx_date, tx_type, seller_name, norm_model, display_model, qty, unit_rate, total_amount)
                            VALUES (%s, %s, 'DISPATCH', %s, %s, %s, %s, %s, %s);
                        """, (batch_id, dispatch_date, selected_seller, norm, disp, qty, rate, amt))
                    conn.commit()
                st.success(f"{selected_seller} के खाते में {total_red} कवर दर्ज हो गए और स्टॉक घट गया! (Batch: {batch_id})")

# ==============================================================================
# TAB 4: DAILY LOSS & WASTAGE (PENDING, WRONG, WASTAGE)
# ==============================================================================
elif menu == "⚠️ Daily Loss & Wastage":
    st.header("⚠️ डेली लॉस और वेस्टेज एंट्री (1-Minute Closing)")
    loss_date = st.date_input("तारीख चुनें:", value=date.today())
    stock_df = fetch_stock()
    model_list = stock_df["display_model"].tolist() if not stock_df.empty else []

    col1, col2, col3 = st.columns(3)

    with col1:
        st.subheader("1️⃣ पेंडिंग कवर (Empty Bed)")
        st.caption("खाली जगह पर मशीन चली (सिर्फ़ इंक + बिजली लागत जुड़ेगी, स्टॉक नहीं घटेगा)")
        pending_qty = st.number_input("पेंडिंग / खाली प्रिंट संख्या:", min_value=0, step=1, key="pend_q")
        if st.button("Save Pending Print"):
            if pending_qty > 0:
                with get_conn() as conn, conn.cursor() as cur:
                    cur.execute("""
                        INSERT INTO transactions (batch_id, tx_date, tx_type, display_model, qty, unit_rate, total_amount, notes)
                        VALUES (%s, %s, 'PENDING_PRINT', 'Empty Bed Print', %s, 0, 0, 'Ink & Electricity Only');
                    """, (f"PND-{uuid.uuid4().hex[:6].upper()}", loss_date, pending_qty))
                    conn.commit()
                st.success(f"{pending_qty} पेंडिंग प्रिंट रिकॉर्ड हो गए!")

    with col2:
        st.subheader("2️⃣ रॉन्ग कवर प्रिंट")
        st.caption("गलत मॉडल पर प्रिंट हुआ (कवर कीमत + इंक जुड़ेगी और स्टॉक घटेगा)")
        w_model = st.selectbox("गलत प्रिंट हुआ मॉडल:", [""] + model_list, key="wrong_m")
        w_qty = st.number_input("गलत प्रिंट क्वांटिटी:", min_value=0, step=1, key="wrong_q")
        if st.button("Save Wrong Print"):
            if w_model and w_qty > 0:
                norm = normalize_model(w_model)
                row = stock_df[stock_df["norm_model"] == norm].iloc[0]
                rate = float(row["avg_rate"])
                with get_conn() as conn, conn.cursor() as cur:
                    cur.execute("UPDATE stock_master SET current_qty = current_qty - %s WHERE norm_model = %s;", (w_qty, norm))
                    cur.execute("""
                        INSERT INTO transactions (batch_id, tx_date, tx_type, norm_model, display_model, qty, unit_rate, total_amount)
                        VALUES (%s, %s, 'WRONG_PRINT', %s, %s, %s, %s, %s);
                    """, (f"WRG-{uuid.uuid4().hex[:6].upper()}", loss_date, norm, w_model, w_qty, rate, w_qty * rate))
                    conn.commit()
                st.success(f"{w_model} के {w_qty} पीस Wrong Print में जुड़े और स्टॉक घटा!")

    with col3:
        st.subheader("3️⃣ वेस्टेज / कटे कवर")
        st.caption("प्रिंटिंग में कटा या खराब हुआ कवर (100% कवर + इंक लॉस और स्टॉक घटेगा)")
        wst_model = st.selectbox("खराब हुआ मॉडल:", [""] + model_list, key="wst_m")
        wst_qty = st.number_input("वेस्टेज क्वांटिटी:", min_value=0, step=1, key="wst_q")
        if st.button("Save Wastage"):
            if wst_model and wst_qty > 0:
                norm = normalize_model(wst_model)
                row = stock_df[stock_df["norm_model"] == norm].iloc[0]
                rate = float(row["avg_rate"])
                with get_conn() as conn, conn.cursor() as cur:
                    cur.execute("UPDATE stock_master SET current_qty = current_qty - %s WHERE norm_model = %s;", (wst_qty, norm))
                    cur.execute("""
                        INSERT INTO transactions (batch_id, tx_date, tx_type, norm_model, display_model, qty, unit_rate, total_amount)
                        VALUES (%s, %s, 'WASTAGE', %s, %s, %s, %s, %s);
                    """, (f"WST-{uuid.uuid4().hex[:6].upper()}", loss_date, norm, wst_model, wst_qty, rate, wst_qty * rate))
                    conn.commit()
                st.success(f"{wst_model} के {wst_qty} पीस Wastage में जुड़े और स्टॉक घटा!")

# ==============================================================================
# TAB 5: INK & FIXED EXPENSES
# ==============================================================================
elif menu == "💸 Ink & Fixed Expenses":
    st.header("💸 UV इंक और फिक्स खर्चे (8–10 Expense Heads)")
    exp_categories = [
        "UV Ink (White / CMYK / Varnish)",
        "Electricity Bill (बिजली बिल)",
        "Staff Salary (स्टाफ सैलरी)",
        "Packing Material (पॉलीबैग, बॉक्स, लेबल)",
        "Shop / Unit Rent (किराया)",
        "Machine Maintenance (डैम्पर, वाइपर, हेड)",
        "Internet & Office Expense",
        "Transport / Freight (भाड़ा)",
        "Other Fixed Expense 1",
        "Other Fixed Expense 2"
    ]

    with st.form("expense_form"):
        c1, c2 = st.columns(2)
        e_date = c1.date_input("खर्च की तारीख:", value=date.today())
        e_cat = c2.selectbox("खर्च की कैटेगरी चुनें:", exp_categories)

        c3, c4, c5 = st.columns(3)
        e_item = c3.text_input("विवरण (जैसे: White Ink 1L / बिजली बिल):")
        e_qty = c4.text_input("क्वांटिटी (जैसे: 2 Liter / 10 Roll):")
        e_amt = c5.number_input("कुल कीमत (₹ Amount):", min_value=0.0, step=10.0)
        e_notes = st.text_input("नोट्स (Optional):")

        if st.form_submit_button("💾 खर्च सेव करें", type="primary"):
            if e_amt > 0:
                with get_conn() as conn, conn.cursor() as cur:
                    cur.execute("""
                        INSERT INTO expenses (exp_date, category, item_name, qty_info, amount, notes)
                        VALUES (%s, %s, %s, %s, %s, %s);
                    """, (e_date, e_cat, e_item, e_qty, e_amt, e_notes))
                    conn.commit()
                st.success(f"₹{e_amt:,.2f} का खर्च ({e_cat}) सेव हो गया!")

    st.subheader("हाल के खर्चे (Recent Expenses)")
    with get_conn() as conn, conn.cursor() as cur:
        cur.execute("SELECT id, exp_date, category, item_name, qty_info, amount, notes FROM expenses ORDER BY exp_date DESC, id DESC LIMIT 50;")
        exp_df = pd.DataFrame(cur.fetchall())
    if not exp_df.empty:
        st.dataframe(exp_df, use_container_width=True, hide_index=True)

# ==============================================================================
# TAB 6: 7-DAY VELOCITY & 10-DAY AUTO ORDER SHEET
# ==============================================================================
elif menu == "🚨 10-Day Auto Order Sheet":
    st.header("🚨 7-दिन की खपत से 10-दिन का ऑटोमैटिक ऑर्डर")
    col1, col2 = st.columns(2)
    lookback_days = col1.number_input("पिछले कितने दिनों की सेल देखें?", min_value=1, max_value=30, value=7)
    target_days = col2.number_input("अगले कितने दिनों के लिए स्टॉक ऑर्डर करना है?", min_value=1, max_value=30, value=10)

    since_date = date.today() - timedelta(days=int(lookback_days))
    stock_df = fetch_stock()

    with get_conn() as conn, conn.cursor() as cur:
        cur.execute("""
            SELECT norm_model, SUM(qty) as sold_qty
            FROM transactions
            WHERE tx_type IN ('DISPATCH', 'WRONG_PRINT', 'WASTAGE')
              AND tx_date >= %s
            GROUP BY norm_model;
        """, (since_date,))
        sales_map = {r["norm_model"]: int(r["sold_qty"]) for r in cur.fetchall()}

    if not stock_df.empty:
        order_rows = []
        for _, r in stock_df.iterrows():
            norm = r["norm_model"]
            curr = int(r["current_qty"])
            sold_7d = sales_map.get(norm, 0)
            daily_avg = sold_7d / float(lookback_days)
            req_10d = round(daily_avg * float(target_days))
            to_order = max(0, req_10d - curr)

            # Also flag if below minimum limit even with low 7d sales
            if curr <= int(r["low_stock_limit"]) and to_order == 0:
                to_order = max(0, int(r["low_stock_limit"]) * 2 - curr)

            if to_order > 0 or curr <= int(r["low_stock_limit"]):
                order_rows.append({
                    "Seller": r["default_seller"] or "Unassigned",
                    "Model Name": r["display_model"],
                    "Current Stock": curr,
                    f"{lookback_days}-Day Print": sold_7d,
                    "Daily Avg": round(daily_avg, 1),
                    f"{target_days}-Day Need": req_10d,
                    "Order Qty": int(to_order),
                    "Est. Rate (₹)": float(r["avg_rate"]),
                    "Est. Cost (₹)": round(to_order * float(r["avg_rate"]), 2)
                })

        if order_rows:
            ord_df = pd.DataFrame(order_rows).sort_values(by=["Seller", "Order Qty"], ascending=[True, False])
            edited_ord = st.data_editor(ord_df, use_container_width=True, hide_index=True)

            st.info(f"*कुल ऑर्डर क्वांटिटी:* {int(edited_ord['Order Qty'].sum()):,} पीस | *अनुमानित कुल बजट:* ₹{edited_ord['Est. Cost (₹)'].sum():,.2f}")

            # Group by seller into separate Excel tabs for 1-click WhatsApp sharing
            sheets = {"All_Orders": edited_ord}
            for s_name, s_grp in edited_ord.groupby("Seller"):
                sheets[str(s_name)] = s_grp[["Model Name", "Order Qty", "Current Stock", f"{lookback_days}-Day Print"]]

            excel_data = to_excel_bytes(sheets)
            st.download_button(
                label="📥 10-Day Order Excel डाउनलोड करें (Seller-Wise Tabs के साथ)",
                data=excel_data,
                file_name=f"10Day_Cover_Order_{date.today()}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                type="primary"
            )
        else:
            st.success("अभी सभी मॉडल्स का स्टॉक पर्याप्त है! 10 दिन के लिए कोई नया ऑर्डर नहीं बन रहा।")

# ==============================================================================
# TAB 7: SELLER LEDGER & UNDO HISTORY
# ==============================================================================
elif menu == "📒 Seller Ledger & Undo":
    st.header("📒 सेलर-वाइज़ मंथली हिसाब और Undo")
    c1, c2 = st.columns(2)
    f_date = c1.date_input("From Date:", value=date.today().replace(day=1), key="led_f")
    t_date = c2.date_input("To Date:", value=date.today(), key="led_t")

    with get_conn() as conn, conn.cursor() as cur:
        cur.execute("""
            SELECT seller_name, tx_type, SUM(qty) as pieces, SUM(total_amount) as amount
            FROM transactions
            WHERE tx_date BETWEEN %s AND %s AND seller_name IS NOT NULL
            GROUP BY seller_name, tx_type
            ORDER BY seller_name;
        """, (f_date, t_date))
        led_rows = cur.fetchall()

    if led_rows:
        st.subheader("सेलर समरी (Purchase vs Dispatch)")
        led_df = pd.DataFrame(led_rows)
        st.dataframe(led_df, use_container_width=True, hide_index=True)

        # Export Full Monthly Master Report
        with get_conn() as conn, conn.cursor() as cur:
            cur.execute("SELECT * FROM transactions WHERE tx_date BETWEEN %s AND %s ORDER BY tx_date DESC;", (f_date, t_date))
            all_tx_df = pd.DataFrame(cur.fetchall())
            cur.execute("SELECT * FROM expenses WHERE exp_date BETWEEN %s AND %s ORDER BY exp_date DESC;", (f_date, t_date))
            all_exp_df = pd.DataFrame(cur.fetchall())

        master_bytes = to_excel_bytes({
            "Seller_Summary": led_df,
            "Current_Stock": fetch_stock(),
            "All_Transactions": all_tx_df,
            "Expenses": all_exp_df
        })
        st.download_button(
            "📊 पूरी मंथली मास्टर Excel रिपोर्ट डाउनलोड करें",
            data=master_bytes,
            file_name=f"KK_Master_Report_{f_date}to{t_date}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

    st.divider()
    st.subheader("↩️ ग़लत अपलोड सुधारें (Undo Batch)")
    with get_conn() as conn, conn.cursor() as cur:
        cur.execute("""
            SELECT batch_id, tx_date, tx_type, seller_name, SUM(qty) as total_qty, COUNT(*) as rows_count
            FROM transactions
            GROUP BY batch_id, tx_date, tx_type, seller_name
            ORDER BY MAX(id) DESC LIMIT 20;
        """)
        batches = cur.fetchall()

    if batches:
        b_df = pd.DataFrame(batches)
        st.dataframe(b_df, use_container_width=True, hide_index=True)
        undo_id = st.selectbox("जिस Batch को वापस (Undo) करना है उसे चुनें:", b_df["batch_id"].tolist())
        if st.button("🗑️ इस Batch को Undo करें (स्टॉक वापस पहले जैसा होगा)"):
            with get_conn() as conn, conn.cursor() as cur:
                cur.execute("SELECT tx_type, norm_model, qty FROM transactions WHERE batch_id = %s;", (undo_id,))
                items = cur.fetchall()
                for it in items:
                    if it["tx_type"] == "INWARD":
                        cur.execute("UPDATE stock_master SET current_qty = current_qty - %s WHERE norm_model = %s;", (it["qty"], it["norm_model"]))
                    elif it["tx_type"] in ("DISPATCH", "WRONG_PRINT", "WASTAGE"):
                        cur.execute("UPDATE stock_master SET current_qty = current_qty + %s WHERE norm_model = %s;", (it["qty"], it["norm_model"]))
                cur.execute("DELETE FROM transactions WHERE batch_id = %s;", (undo_id,))
                conn.commit()
            st.success(f"Batch {undo_id} सफलतापूर्वक Undo हो गया!")
            st.rerun()

# ==============================================================================
# TAB 8: SELLER & 3-CATEGORY RATES MASTER
# ==============================================================================
elif menu == "⚙️ Seller & 3-Category Rates":
    st.header("⚙️ सेलर मास्टर और 3 रेट कैटेगरी सेटअप")
    with st.form("seller_form"):
        s_name = st.text_input("सेलर / पार्टी का नाम (जैसे: Rahul Cases / VK / KK & Co):")
        c1, c2 = st.columns(2)
        c1_name = c1.text_input("Category 1 नाम:", value="Normal / Transparent")
        c1_rate = c2.number_input("Category 1 रेट (₹):", min_value=0.0, step=0.5)

        c3, c4 = st.columns(2)
        c2_name = c3.text_input("Category 2 नाम:", value="Matte / Camera Protection")
        c2_rate = c4.number_input("Category 2 रेट (₹):", min_value=0.0, step=0.5)

        c5, c6 = st.columns(2)
        c3_name = c5.text_input("Category 3 नाम:", value="Glass / Hard Case")
        c3_rate = c6.number_input("Category 3 रेट (₹):", min_value=0.0, step=0.5)

        if st.form_submit_button("💾 सेलर और रेट्स सेव / अपडेट करें", type="primary"):
            if s_name.strip():
                with get_conn() as conn, conn.cursor() as cur:
                    cur.execute("""
                        INSERT INTO sellers (name, cat1_name, cat1_rate, cat2_name, cat2_rate, cat3_name, cat3_rate)
                        VALUES (%s, %s, %s, %s, %s, %s, %s)
                        ON CONFLICT (name) DO UPDATE SET
                            cat1_name = EXCLUDED.cat1_name, cat1_rate = EXCLUDED.cat1_rate,
                            cat2_name = EXCLUDED.cat2_name, cat2_rate = EXCLUDED.cat2_rate,
                            cat3_name = EXCLUDED.cat3_name, cat3_rate = EXCLUDED.cat3_rate;
                    """, (s_name.strip(), c1_name, c1_rate, c2_name, c2_rate, c3_name, c3_rate))
                    conn.commit()
                st.success(f"{s_name} के रेट्स सेव हो गए!")

    st.subheader("सेव किए गए सेलर्स और उनके 3-कैटेगरी रेट्स")
    s_list = fetch_sellers()
    if s_list:
        st.dataframe(pd.DataFrame(s_list), use_container_width=True, hide_index=True)