import streamlit as st
import pandas as pd
import json
import os
from io import BytesIO

st.set_page_config(
    page_title="حاسبة كلف الأطعمة — مجموعة عصير تايم",
    layout="wide",
    initial_sidebar_state="collapsed"
)

st.markdown("""
<style>
    @import url('https://fonts.googleapis.com/css2?family=Cairo:wght@400;600;700&display=swap');
    
    html, body, [class*="css"] {
        font-family: 'Cairo', sans-serif;
        direction: rtl;
        text-align: right;
    }
    .stApp {
        background-color: #181b1a;
        color: #e5e7eb;
    }
    .brand-title {
        font-size: 22px;
        font-weight: 700;
        color: #ffffff;
    }
    .info-box {
        background-color: #202422;
        border: 1px solid #2e3532;
        padding: 18px;
        border-radius: 10px;
        margin-bottom: 15px;
        color: #d1d5db;
        font-size: 15px;
        line-height: 1.7;
    }
    div.stButton > button {
        background-color: #10b981;
        color: white;
        border: none;
        border-radius: 8px;
        padding: 8px 20px;
        font-weight: 700;
        font-family: 'Cairo', sans-serif;
        width: 100%;
    }
    div.stButton > button:hover {
        background-color: #059669;
        color: white;
    }
    th, td {
        text-align: right !important;
    }
</style>
""", unsafe_allow_html=True)

DATA_FILE = "asir_time_data.json"
BRANDS = ["فروج أبو العبد", "بابا كنافة", "عصير تايم"]

def load_data():
    if os.path.exists(DATA_FILE):
        with open(DATA_FILE, "r", encoding="utf-8") as f:
            return json.load(f)
    return {brand: {"materials": [], "recipes": []} for brand in BRANDS}

def save_data(data):
    with open(DATA_FILE, "w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, indent=2)

def to_excel_bytes(df):
    output = BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name="Sheet1")
    return output.getvalue()

if "db" not in st.session_state:
    st.session_state.db = load_data()

if "selected_brand" not in st.session_state:
    st.session_state.selected_brand = None

if "active_tab" not in st.session_state:
    st.session_state.active_tab = "الوصفات والكلفة"

if st.session_state.selected_brand is None:
    st.markdown("<h2 style='text-align: right; color: #fff;'>حاسبة كلف الأطعمة — مجموعة عصير تايم</h2>", unsafe_allow_html=True)
    st.markdown("<h4 style='text-align: right; color: #9ca3af; margin-top: 20px;'>اختر العلامة التجارية</h4>", unsafe_allow_html=True)
    
    cols = st.columns(3)
    for idx, brand in enumerate(BRANDS):
        with cols[idx]:
            st.markdown(f"""
            <div style="background-color:#202422; border:1px solid #2e3532; border-radius:12px; padding:30px; text-align:center; margin-bottom:10px;">
                <h3 style="color:#ffffff; margin:0;">{brand}</h3>
            </div>
            """, unsafe_allow_html=True)
            if st.button(f"دخول — {brand}", key=f"brand_{idx}"):
                st.session_state.selected_brand = brand
                st.rerun()
    st.stop()

brand = st.session_state.selected_brand
brand_data = st.session_state.db[brand]
materials_list = brand_data["materials"]
recipes_list = brand_data["recipes"]

top_col1, top_col2 = st.columns([3, 1])
with top_col1:
    st.markdown(f"<div class='brand-title'>{brand}</div>", unsafe_allow_html=True)
with top_col2:
    if st.button("تغيير العلامة"):
        st.session_state.selected_brand = None
        st.rerun()

st.markdown("<hr style='border-color: #2e3532;'>", unsafe_allow_html=True)

t_col1, t_col2, t_col3, _ = st.columns([2, 2, 2.5, 3.5])
with t_col1:
    if st.button(f"الوصفات والكلفة ({len(recipes_list)})"):
        st.session_state.active_tab = "الوصفات والكلفة"
with t_col2:
    if st.button(f"المواد ({len(materials_list)})"):
        st.session_state.active_tab = "المواد"
with t_col3:
    if st.button("مقارنة المبيعات بالمخزن"):
        st.session_state.active_tab = "مقارنة المبيعات بالمخزن"

st.write("")

mat_unit_cost = {m["المادة"]: float(m["كلفة الوحدة الأساسية"]) for m in materials_list}
mat_base_unit = {m["المادة"]: m["الوحدة الأساسية"] for m in materials_list}

if st.session_state.active_tab == "المواد":
    st.markdown("### إدارة المواد الأولية (المخزن)")
    
    with st.expander("إضافة مادة جديدة يدوياً أو رفع ملف المواد (Excel)"):
        c1, c2 = st.columns(2)
        with c1:
            st.markdown("##### إضافة مادة يدوياً")
            m_name = st.text_input("اسم المادة (مثال: دجاج، أرز، جبن)")
            m_p_unit = st.text_input("وحدة الشراء (مثال: كارتون، كيس، كغم)", value="كغم")
            m_p_price = st.number_input("سعر الشراء", min_value=0.0, step=100.0)
            m_b_unit = st.text_input("الوحدة الأساسية للوصفة (مثال: غرام، حبة، مل)", value="غرام")
            m_b_count = st.number_input("عدد الأساسية بوحدة الشراء", min_value=1.0, value=1000.0, step=10.0)
            
            if st.button("حفظ المادة"):
                if m_name.strip():
                    base_cost = round(m_p_price / m_b_count, 4)
                    materials_list = [m for m in materials_list if m["المادة"] != m_name.strip()]
                    materials_list.append({
                        "المادة": m_name.strip(),
                        "وحدة الشراء": m_p_unit,
                        "سعر الشراء": m_p_price,
                        "الوحدة الأساسية": m_b_unit,
                        "عدد الأساسية بوحدة الشراء": m_b_count,
                        "كلفة الوحدة الأساسية": base_cost
                    })
                    st.session_state.db[brand]["materials"] = materials_list
                    save_data(st.session_state.db)
                    st.success("تم حفظ المادة بنجاح!")
                    st.rerun()

        with c2:
            st.markdown("##### رفع ملف المواد (Excel)")
            sample_mat_df = pd.DataFrame([
                {"المادة": "دجاج", "وحدة الشراء": "كغم", "سعر الشراء": 5000, "الوحدة الأساسية": "غرام", "عدد الأساسية بوحدة الشراء": 1000},
                {"المادة": "أرز", "وحدة الشراء": "كيس", "سعر الشراء": 25000, "الوحدة الأساسية": "غرام", "عدد الأساسية بوحدة الشراء": 10000}
            ])
            st.download_button(
                "تحميل قالب ملف المواد (Excel)",
                data=to_excel_bytes(sample_mat_df),
                file_name="template_materials.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
            uploaded_mat = st.file_uploader("اختر ملف المواد (Excel)", type=["xlsx", "xls"], key="up_mat")
            if uploaded_mat is not None:
                if st.button("استيراد المواد من الملف"):
                    df_m = pd.read_excel(uploaded_mat)
                    new_mats = []
                    for _, row in df_m.iterrows():
                        p_price = float(row.get("سعر الشراء", 0))
                        b_count = float(row.get("عدد الأساسية بوحدة الشراء", 1)) or 1.0
                        new_mats.append({
                            "المادة": str(row.get("المادة", "")).strip(),
                            "وحدة الشراء": str(row.get("وحدة الشراء", "")),
                            "سعر الشراء": p_price,
                            "الوحدة الأساسية": str(row.get("الوحدة الأساسية", "")),
                            "عدد الأساسية بوحدة الشراء": b_count,
                            "كلفة الوحدة الأساسية": round(p_price / b_count, 4)
                        })
                    st.session_state.db[brand]["materials"] = new_mats
                    save_data(st.session_state.db)
                    st.success("تم رفع المواد بنجاح!")
                    st.rerun()

    search_mat = st.text_input("بحث...", key="search_mat", placeholder="بحث عن مادة...")
    
    if materials_list:
        df_show_mat = pd.DataFrame(materials_list)
        if search_mat:
            df_show_mat = df_show_mat[df_show_mat["المادة"].str.contains(search_mat, case=False, na=False)]
        cols_order = ["المادة", "وحدة الشراء", "سعر الشراء", "الوحدة الأساسية", "عدد الأساسية بوحدة الشراء", "كلفة الوحدة الأساسية"]
        st.dataframe(df_show_mat[cols_order], use_container_width=True, hide_index=True)
        
        if st.button("مسح جميع المواد"):
            st.session_state.db[brand]["materials"] = []
            save_data(st.session_state.db)
            st.rerun()
    else:
        st.info("لا توجد مواد مضافة حالياً. أضف المواد من الأعلى.")

elif st.session_state.active_tab == "الوصفات والكلفة":
    st.markdown("### الوصفات والكلفة")
    
    with st.expander("إضافة وصفة جديدة أو رفع ملف الوصفات (Excel)"):
        rc1, rc2 = st.columns(2)
        with rc1:
            st.markdown("##### إضافة وصفة يدوياً")
            r_name = st.text_input("اسم المنتج (الوجبة)")
            r_price = st.number_input("سعر البيع", min_value=0.0, step=250.0)
            
            mat_names = [m["المادة"] for m in materials_list]
            selected_ing = {}
            for i in range(1, 6):
                ic1, ic2 = st.columns([2, 1])
                with ic1:
                    ing_name = st.selectbox(f"المادة {i}", ["-- اختر --"] + mat_names, key=f"ing_n_{i}")
                with ic2:
                    ing_qty = st.number_input(f"الكمية {i}", min_value=0.0, step=10.0, key=f"ing_q_{i}")
                if ing_name != "-- اختر --" and ing_qty > 0:
                    selected_ing[ing_name] = ing_qty
            
            if st.button("حفظ الوصفة"):
                if r_name.strip() and selected_ing:
                    recipes_list = [r for r in recipes_list if r["المنتج"] != r_name.strip()]
                    recipes_list.append({
                        "المنتج": r_name.strip(),
                        "سعر البيع": r_price,
                        "المكونات": selected_ing
                    })
                    st.session_state.db[brand]["recipes"] = recipes_list
                    save_data(st.session_state.db)
                    st.success("تم حفظ الوصفة بنجاح!")
                    st.rerun()

        with rc2:
            st.markdown("##### رفع ملف الوصفات (Excel)")
            sample_rec_df = pd.DataFrame([
                {"المنتج": "وجبة فروج", "سعر البيع": 12000, "المادة": "دجاج", "الكمية": 500},
                {"المنتج": "وجبة فروج", "سعر البيع": 12000, "المادة": "أرز", "الكمية": 250}
            ])
            st.download_button(
                "تحميل قالب الوصفات (Excel)",
                data=to_excel_bytes(sample_rec_df),
                file_name="template_recipes.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
            uploaded_rec = st.file_uploader("اختر ملف الوصفات (Excel)", type=["xlsx", "xls"], key="up_rec")
            if uploaded_rec is not None:
                if st.button("استيراد الوصفات من الملف"):
                    df_r = pd.read_excel(uploaded_rec)
                    grouped = {}
                    for _, row in df_r.iterrows():
                        p_name = str(row.get("المنتج", "")).strip()
                        p_price = float(row.get("سعر البيع", 0))
                        m_item = str(row.get("المادة", "")).strip()
                        m_qty = float(row.get("الكمية", 0))
                        if p_name:
                            if p_name not in grouped:
                                grouped[p_name] = {"المنتج": p_name, "سعر البيع": p_price, "المكونات": {}}
                            if m_item and m_qty > 0:
                                grouped[p_name]["المكونات"][m_item] = m_qty
                    st.session_state.db[brand]["recipes"] = list(grouped.values())
                    save_data(st.session_state.db)
                    st.success("تم استيراد الوصفات بنجاح!")
                    st.rerun()

    search_rec = st.text_input("بحث...", key="search_rec", placeholder="بحث عن منتج...")
    
    if recipes_list:
        rows_rec = []
        for r in recipes_list:
            p_name = r["المنتج"]
            sell_price = float(r["سعر البيع"])
            ings = r.get("المكونات", {})
            
            total_cost = sum(float(qty) * mat_unit_cost.get(m_name, 0.0) for m_name, qty in ings.items())
            cost_pct = round((total_cost / sell_price) * 100, 2) if sell_price > 0 else 0.0
            profit = round(sell_price - total_cost, 2)
            ings_text = " ، ".join([f"{m}: {q} {mat_base_unit.get(m, '')}" for m, q in ings.items()])
            
            rows_rec.append({
                "المنتج": p_name,
                "سعر البيع": sell_price,
                "الكلفة": round(total_cost, 2),
                "نسبة الكلفة": f"{cost_pct}%",
                "الربح": profit,
                "المواد": ings_text
            })
        df_show_rec = pd.DataFrame(rows_rec)
        if search_rec:
            df_show_rec = df_show_rec[df_show_rec["المنتج"].str.contains(search_rec, case=False, na=False)]
        
        cols_rec_order = ["المنتج", "سعر البيع", "الكلفة", "نسبة الكلفة", "الربح", "المواد"]
        st.dataframe(df_show_rec[cols_rec_order], use_container_width=True, hide_index=True)
    else:
        st.info("لا توجد وصفات مضافة حالياً. أضف الوصفات من الأعلى.")

elif st.session_state.active_tab == "مقارنة المبيعات بالمخزن":
    st.markdown("""
    <div class="info-box">
        ارفع ملف المبيعات وملف سحوبات المخزن للفترة نفسها. يحسب البرنامج المواد التي كان يجب أن تُستهلك لتغطية المبيعات حسب الوصفات المعتمدة، ويقارنها بما سُحب فعلاً. الملفات لا تُحفظ.
    </div>
    """, unsafe_allow_html=True)

    with st.expander("تحميل قوالب Excel الجاهزة للمبيعات والمخزن (للتجربة)"):
        tc1, tc2 = st.columns(2)
        with tc1:
            sample_sales = pd.DataFrame([
                {"المنتج": "وجبة فروج", "الكمية المباعة": 50}
            ])
            st.download_button(
                "قالب ملف المبيعات (Excel)",
                data=to_excel_bytes(sample_sales),
                file_name="sales_sample.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
        with tc2:
            sample_store = pd.DataFrame([
                {"المادة": "دجاج", "الكمية المسحوبة": 27000},
                {"المادة": "أرز", "الكمية المسحوبة": 13000}
            ])
            st.download_button(
                "قالب ملف سحوبات المخزن (Excel)",
                data=to_excel_bytes(sample_store),
                file_name="store_withdrawals_sample.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

    u_col1, u_col2, u_col3 = st.columns([2, 2, 1])
    with u_col1:
        sales_file = st.file_uploader("ملف المبيعات (Excel)", type=["xlsx", "xls"], key="sales_excel")
    with u_col2:
        store_file = st.file_uploader("ملف سحوبات المخزن (Excel)", type=["xlsx", "xls"], key="store_excel")
    with u_col3:
        tolerance_pct = st.number_input("نسبة السماح %", min_value=0.0, max_value=100.0, value=5.0, step=1.0)

    btn_col1, _ = st.columns([1.5, 8.5])
    with btn_col1:
        audit_clicked = st.button("تدقيق")

    if audit_clicked:
        if sales_file is None or store_file is None:
            st.warning("ارفع الملفين أولاً")
        elif not recipes_list:
            st.error("يرجى إضافة الوصفات والمواد أولاً حتى يتمكن البرنامج من الحساب.")
        else:
            df_sales = pd.read_excel(sales_file)
            df_store = pd.read_excel(store_file)

            sales_prod_col = "المنتج" if "المنتج" in df_sales.columns else df_sales.columns[0]
            sales_qty_col = "الكمية المباعة" if "الكمية المباعة" in df_sales.columns else df_sales.columns[1]

            store_mat_col = "المادة" if "المادة" in df_store.columns else df_store.columns[0]
            store_qty_col = "الكمية المسحوبة" if "الكمية المسحوبة" in df_store.columns else df_store.columns[1]

            expected_usage = {}
            recipes_map = {r["المنتج"]: r["المكونات"] for r in recipes_list}

            for _, s_row in df_sales.iterrows():
                prod = str(s_row[sales_prod_col]).strip()
                sold_qty = float(s_row[sales_qty_col]) if pd.notna(s_row[sales_qty_col]) else 0.0
                if prod in recipes_map:
                    for mat_name, ing_qty in recipes_map[prod].items():
                        expected_usage[mat_name] = expected_usage.get(mat_name, 0.0) + (sold_qty * float(ing_qty))

            actual_withdrawn = {}
            for _, st_row in df_store.iterrows():
                mat = str(st_row[store_mat_col]).strip()
                w_qty = float(st_row[store_qty_col]) if pd.notna(st_row[store_qty_col]) else 0.0
                actual_withdrawn[mat] = actual_withdrawn.get(mat, 0.0) + w_qty

            all_mats = sorted(set(list(expected_usage.keys()) + list(actual_withdrawn.keys())))
            audit_rows = []
            total_waste_cost = 0.0

            for mat in all_mats:
                exp_q = round(expected_usage.get(mat, 0.0), 2)
                act_q = round(actual_withdrawn.get(mat, 0.0), 2)
                diff_q = round(act_q - exp_q, 2)
                diff_pct = round((abs(diff_q) / exp_q) * 100, 2) if exp_q > 0 else (100.0 if act_q > 0 else 0.0)
                
                u_cost = mat_unit_cost.get(mat, 0.0)
                diff_cost = round(diff_q * u_cost, 2)
                if diff_cost > 0:
                    total_waste_cost += diff_cost

                if abs(diff_q) == 0 or diff_pct <= tolerance_pct:
                    status = "ضمن السماح"
                elif diff_q > 0:
                    status = "هدر / سحب زائد"
                else:
                    status = "سحب أقل من المبيعات"

                audit_rows.append({
                    "المادة": mat,
                    "الوحدة": mat_base_unit.get(mat, ""),
                    "الاستهلاك المفترض": exp_q,
                    "المسحوب فعلاً": act_q,
                    "الفرق": diff_q,
                    "نسبة الفرق %": f"{diff_pct}%",
                    "كلفة الفرق": diff_cost,
                    "الحالة": status
                })

            st.markdown("### نتيجة تدقيق المبيعات بالمخزن")
            st.dataframe(pd.DataFrame(audit_rows), use_container_width=True, hide_index=True)
            st.markdown(f"*إجمالي كلفة الهدر / السحب الزائد:* {round(total_waste_cost, 2)}")