# -*- coding: utf-8 -*-
"""
حساب الخامات الإجمالية المطلوبة للشراء
--------------------------------------
يقرأ ملف الإكسل الذي يحتوي على:
  - تابة "الخامات للمنتجات" : قائمة الخامات لكل منتج في كل مرحلة إنتاج (القسم المستهلك)
  - تابة "الاعداد"          : العدد المطلوب إنتاجه من كل منتج

وينتج ملف إكسل فيه:
  1) ملخص
  2) الإجمالي              (إجمالي كل خامة على مستوى المصنع)
  3) الإجمالي لكل قسم       (قائمة الشراء مقسمة على الأقسام)
  4) فرعي - تفصيل المنتجات  (كل سطر: منتج × خامة)
  5) ملاحظات ومشاكل البيانات

الاستخدام:
    python calc_materials.py <ملف_المدخلات.xlsx> [ملف_المخرجات.xlsx]
                             [--extra-bom خامات.csv] [--extra-counts اعداد.csv]

الملفان الاختياريان يضيفان منتجات غير موجودة في ملف الإكسل:
    --extra-bom     أعمدة: المنتج، كود المنتج، القسم، كود الخامة، اسم الخامة،
                    الكمية للوحدة، الوحدة
    --extra-counts  أعمدة: اسم المنتج، العدد
"""

import re
import sys
from collections import defaultdict

import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.table import Table, TableStyleInfo

# تابات ملف المصدر الأصلي
SHEET_MATERIALS = "الخامات للمنتجات اريكة"
SHEET_COUNTS = "الاعداد "
HEADER_ROW = 4  # صف العناوين في تابة الخامات (1-based)

# تابات المدخلات داخل الملف الناتج — الملف يقرأ نفسه، فيمكن التعديل عليه وإعادة التشغيل
IN_MATERIALS = "مدخلات - قائمة الخامات"
IN_COUNTS = "مدخلات - الأعداد المطلوبة"

IN_MATERIALS_COLS = ["المنتج", "الموديل", "كود المنتج", "القسم الإنتاجي",
                     "كود الخامة", "اسم الخامة", "الكمية للوحدة", "الوحدة", "المصدر"]
IN_COUNTS_COLS = ["اسم المنتج", "العدد"]

# حساب القشرة: تابات ملف القشرة ثم تابة المدخلات داخل الملف الناتج
VEN_SHEET_COUNTS = "الاعداد "
VEN_COUNTS_LABEL = "تابة الأعداد في ملف القشرة"
VEN_SHEET_NEED = "احتياج القشرة"
IN_VENEER = "مدخلات - احتياج القشرة"
IN_VENEER_COLS = ["المنتج (كما ورد)", "النص الأصلي", "المنتج المطابق",
                  "نوع القشرة", "متر للوحدة", "العدد المطلوب", "مصدر العدد"]

# أسماء في تابة القشرة مكتوبة بصيغة مختلفة عن تابة الأعداد
VENEER_ALIASES = {
    "palma dresser": "Palma Drawer Dresser",
    "beech side": "Beech Side Tables",
    "mirlia dresser": "Marlia Drawer Dresser",
    "sleek coffebar coffee": "Sleek Coffe Bar",
    "malia tv": "Malia TV unit",
    "fluted round": "Fluted Round Middle Table",
    "rattan corner shelf coffee": "Rattan Corner Shelf",
    "costa rattan desk coffee": "Costa Rattan Desk",
    "turin chair": "فوتية تورين",
    # سطر واحد يغطي المقاسين معًا — العدد = عدد الأطقم
    "odeno middle table set": "Odeon Middle Table large",
}

# كلمات الكميات المكتوبة نصًا في عمود القشرة
QTY_WORDS = {"نص": 0.5, "نصف": 0.5, "ربع": 0.25, "واحد": 1.0}

COLS = ["code", "collection", "prod_en", "model", "piece_type",
        "furx_name", "old_code", "old_name", "category", "version",
        "dept", "mat_code", "mat_name", "mat_type", "qty_per_unit",
        "unit", "status"]

# توحيد أسماء الأقسام المكتوبة بأكثر من صيغة
DEPT_MAP = {
    "دهانات": "الدهانات",
    "الدهانات": "الدهانات",
    "تغليف": "التغليف",
    "التغليف": "التغليف",
    "نجاره نوم وسفره": "نجارة النوم والسفرة",
    "نجاره نوم و السفره": "نجارة النوم والسفرة",
    "نجارة النوم والسفرة": "نجارة النوم والسفرة",
    "تشطيب نوم وسفرة": "تشطيب النوم والسفرة",
    "قشره": "القشرة",
    "القشرة": "القشرة",
    "التفصيل والكسوة": "التفصيل والكسوة",
    "التفصيل والكسوه": "التفصيل والكسوة",
    "تفصيل وكسوه": "التفصيل والكسوة",
    "نجارة التنجيد": "نجارة التنجيد",
    "السفنجة": "السفنجة",
    "السفنجه": "السفنجة",
    "تشطيب التنجيد": "تشطيب التنجيد",
    "تشطيب تنجيد": "تشطيب التنجيد",
}

# توحيد وحدات القياس
UNIT_MAP = {"علبه": "علبة", "meter": "متر"}

UNSPECIFIED = "غير محدد"

# المنتجات التي ليس لها اسم إنجليزي في تابة الخامات ونربطها يدويًا بتابة الأعداد
MODEL_TO_EN = {"كوستا راتان": "Costa Rattan Desk"}


# --------------------------------------------------------------------------
# أدوات مساعدة
# --------------------------------------------------------------------------
def clean_text(value):
    """إزالة المسافات الزائدة وتحويل القيم الفارغة إلى نص فارغ."""
    if pd.isna(value):
        return ""
    return re.sub(r"\s+", " ", str(value)).strip()


def match_key(value):
    """مفتاح مطابقة أسماء المنتجات: بدون أقواس ولا مسافات زائدة ولا حروف كبيرة."""
    text = clean_text(value)
    text = re.sub(r"[\[\]()_]", " ", text)  # بعض الأسماء بها _ أو سطر جديد
    return re.sub(r"\s+", " ", text).strip().lower()


def code_to_text(value):
    """كود الخامة كنص: بدون .0 من قراءة الأرقام وبدون بادئة النظام (RM- / FB-)."""
    if pd.isna(value):
        return ""
    if isinstance(value, float) and value.is_integer():
        return str(int(value))
    # النظام الجديد يكتب الأكواد بالصيغة RM-11003 بينما الملف القديم يكتبها 11003
    return re.sub(r"^[A-Za-z]{2,3}-", "", clean_text(value))


# --------------------------------------------------------------------------
# قراءة المدخلات
# --------------------------------------------------------------------------
def load_input(path):
    """يقرأ الملف الأصلي أو ملفًا ناتجًا عن هذا السكربت (تابات المدخلات بداخله)."""
    names = pd.ExcelFile(path).sheet_names
    if IN_MATERIALS in names and IN_COUNTS in names:
        return load_consolidated(path)

    materials = pd.read_excel(path, sheet_name=SHEET_MATERIALS, header=HEADER_ROW - 1)
    if len(materials.columns) != len(COLS):
        raise SystemExit(
            f"تابة الخامات بها {len(materials.columns)} عمود بينما المتوقع {len(COLS)}."
        )
    materials.columns = COLS
    materials["source_row"] = materials.index + HEADER_ROW + 1

    counts = pd.read_excel(path, sheet_name=SHEET_COUNTS)
    counts.columns = ["product", "count"]
    counts = counts[counts["product"].notna()].copy()
    counts["product"] = counts["product"].map(clean_text)
    counts["count"] = pd.to_numeric(counts["count"], errors="coerce").fillna(0)
    return materials, counts


def load_consolidated(path):
    """قراءة تابتي المدخلات من ملف أنتجه هذا السكربت."""
    src = pd.read_excel(path, sheet_name=IN_MATERIALS)
    missing = [c for c in IN_MATERIALS_COLS[:-1] if c not in src.columns]
    if missing:
        raise SystemExit(f"ينقص تابة «{IN_MATERIALS}» الأعمدة: {', '.join(missing)}")

    materials = pd.DataFrame({col: "" for col in COLS}, index=src.index)
    materials["prod_en"] = src["المنتج"]
    materials["model"] = src["الموديل"]
    materials["code"] = src["كود المنتج"]
    materials["dept"] = src["القسم الإنتاجي"]
    materials["mat_code"] = src["كود الخامة"]
    materials["mat_name"] = src["اسم الخامة"]
    materials["mat_type"] = "خامة"
    materials["qty_per_unit"] = src["الكمية للوحدة"]
    materials["unit"] = src["الوحدة"]
    materials["status"] = "له قائمة"
    materials["source_row"] = src.get("المصدر", pd.Series("", index=src.index))

    counts = pd.read_excel(path, sheet_name=IN_COUNTS)
    counts = counts.iloc[:, :2]
    counts.columns = ["product", "count"]
    counts = counts[counts["product"].notna()].copy()
    counts["product"] = counts["product"].map(clean_text)
    counts["count"] = pd.to_numeric(counts["count"], errors="coerce").fillna(0)
    return materials, counts


def load_extra(bom_path, counts_path):
    """قراءة منتجات إضافية من ملفي CSV وتحويلها لنفس شكل تابة الخامات."""
    if not bom_path:
        return None, None

    bom = pd.read_csv(bom_path)
    required = ["المنتج", "القسم", "كود الخامة", "اسم الخامة", "الكمية للوحدة", "الوحدة"]
    missing = [c for c in required if c not in bom.columns]
    if missing:
        raise SystemExit(f"ينقص ملف الخامات الإضافي الأعمدة: {', '.join(missing)}")

    rows = pd.DataFrame({col: "" for col in COLS}, index=bom.index)
    rows["prod_en"] = bom["المنتج"]
    rows["model"] = bom["المنتج"]
    rows["code"] = bom.get("كود المنتج", "")
    rows["dept"] = bom["القسم"]
    rows["mat_code"] = bom["كود الخامة"]
    rows["mat_name"] = bom["اسم الخامة"]
    rows["mat_type"] = "خامة"
    rows["qty_per_unit"] = bom["الكمية للوحدة"]
    rows["unit"] = bom["الوحدة"]
    rows["status"] = "له قائمة"
    rows["source_row"] = [f"{bom_path} صف {i + 2}" for i in bom.index]

    extra_counts = None
    if counts_path:
        extra_counts = pd.read_csv(counts_path)
        extra_counts.columns = ["product", "count"]
        extra_counts["product"] = extra_counts["product"].map(clean_text)
        extra_counts["count"] = pd.to_numeric(extra_counts["count"], errors="coerce").fillna(0)
    return rows, extra_counts


def prepare(materials, counts):
    """تنظيف البيانات وربط كل خامة بالعدد المطلوب من منتجها."""
    df = materials.copy()

    df["product"] = df["prod_en"].map(clean_text)
    df["model"] = df["model"].map(clean_text)
    # المنتجات بدون اسم إنجليزي: نستخدم الربط اليدوي وإلا نستخدم الاسم العربي
    df["product"] = [
        prod if prod else MODEL_TO_EN.get(model, model)
        for prod, model in zip(df["product"], df["model"])
    ]
    df["product_code"] = df["code"].map(clean_text)
    df["dept"] = df["dept"].map(clean_text).map(lambda d: DEPT_MAP.get(d, d) or UNSPECIFIED)
    df["mat_code"] = df["mat_code"].map(code_to_text)
    df["mat_name"] = df["mat_name"].map(clean_text)
    df["unit"] = df["unit"].map(clean_text).map(lambda u: UNIT_MAP.get(u, u) or UNSPECIFIED)
    df["qty_per_unit"] = pd.to_numeric(df["qty_per_unit"], errors="coerce").fillna(0.0)

    # مفتاح الخامة: الكود إن وُجد، وإلا الاسم
    df["mat_key"] = [
        code if code else f"بدون كود: {name}"
        for code, name in zip(df["mat_code"], df["mat_name"])
    ]

    # اسم موحّد لكل مفتاح خامة (الاسم الأكثر تكرارًا)
    canonical = (
        df[df["mat_name"] != ""]
        .groupby("mat_key")["mat_name"]
        .agg(lambda s: s.value_counts().idxmax())
        .to_dict()
    )
    df["mat_name_std"] = [
        canonical.get(key, name or "(بدون اسم)")
        for key, name in zip(df["mat_key"], df["mat_name"])
    ]

    # ربط الأعداد المطلوبة
    wanted = {match_key(p): c for p, c in zip(counts["product"], counts["count"])}
    df["product_count"] = df["product"].map(lambda p: wanted.get(match_key(p), 0.0))
    df["total_qty"] = df["qty_per_unit"] * df["product_count"]
    return df


# --------------------------------------------------------------------------
# التجميع
# --------------------------------------------------------------------------
def build_tables(df):
    used = df[df["product_count"] > 0].copy()

    detail = (
        used.groupby(
            ["product", "product_code", "model", "product_count",
             "dept", "mat_key", "mat_code", "mat_name_std", "unit"],
            as_index=False, dropna=False,
        )
        .agg(qty_per_unit=("qty_per_unit", "sum"), total_qty=("total_qty", "sum"))
        .sort_values(["product", "dept", "mat_name_std"], kind="stable")
    )

    by_dept = (
        used.groupby(["dept", "mat_key", "mat_code", "mat_name_std", "unit"],
                     as_index=False, dropna=False)
        .agg(total_qty=("total_qty", "sum"),
             products=("product", pd.Series.nunique))
        .sort_values(["dept", "mat_name_std"], kind="stable")
    )

    grand = (
        used.groupby(["mat_key", "mat_code", "mat_name_std", "unit"],
                     as_index=False, dropna=False)
        .agg(total_qty=("total_qty", "sum"),
             products=("product", pd.Series.nunique),
             depts=("dept", lambda s: " + ".join(sorted(set(s)))))
        .sort_values("total_qty", ascending=False, kind="stable")
    )
    return detail, by_dept, grand


def build_issues(df, counts):
    """كل ما يحتاج مراجعة بشرية في ملف المدخلات."""
    rows = []

    matched = {match_key(p) for p in df["product"]}
    for product, count in zip(counts["product"], counts["count"]):
        if match_key(product) not in matched:
            rows.append(["منتج مطلوب بلا قائمة خامات", product,
                         f"مطلوب {count:g} قطعة لكنه غير موجود في تابة الخامات — لم يُحسب."])

    for model, group in df[df["product_count"] == 0].groupby("model"):
        rows.append(["منتج بلا عدد مطلوب", model,
                     f"له {len(group)} سطر خامات لكنه غير موجود في تابة الأعداد — لم يُحسب."])

    missing_code = df[(df["mat_code"] == "") & (df["product_count"] > 0)]
    for _, r in missing_code.iterrows():
        rows.append(["خامة بدون كود", r["mat_name_std"],
                     f"صف {r['source_row']} — منتج {r['product']} / {r['dept']}. جُمعت بالاسم."])

    missing_unit = df[(df["unit"] == UNSPECIFIED) & (df["product_count"] > 0)]
    for _, r in missing_unit.iterrows():
        rows.append(["خامة بدون وحدة", r["mat_name_std"],
                     f"صف {r['source_row']} — منتج {r['product']} / {r['dept']}."])

    no_name = df[(df["mat_name"] == "") & (df["product_count"] > 0)]
    for _, r in no_name.iterrows():
        rows.append(["خامة بدون اسم", f"كود {r['mat_code']}",
                     f"صف {r['source_row']} — منتج {r['product']} / {r['dept']}."])

    used = df[df["product_count"] > 0]
    for key, group in used.groupby("mat_key"):
        units = sorted(set(group["unit"]))
        if len(units) > 1:
            rows.append(["وحدة مختلفة لنفس الخامة", group["mat_name_std"].iloc[0],
                         "مسجلة بوحدات: " + " و ".join(units) + " — فُصلت في صفوف مستقلة."])
        names = sorted({n for n in group["mat_name"] if n})
        if len(names) > 1:
            rows.append(["اسمان لنفس الكود", f"كود {key}",
                         " / ".join(names) + " — وُحّدت على الاسم الأكثر تكرارًا."])

    for name, group in used[used["mat_code"] != ""].groupby("mat_name"):
        codes = sorted(set(group["mat_code"]))
        if len(codes) > 1:
            rows.append(["كودان لنفس الاسم", name,
                         "الأكواد: " + " / ".join(codes) + " — عُوملت كخامتين منفصلتين."])

    dupes = used.groupby(["product", "dept", "mat_key", "unit"]).size().reset_index(name="n")
    dupes = dupes[dupes["n"] > 1]
    for product, group in dupes.groupby("product"):
        codes = " / ".join(sorted(set(group["mat_key"])))
        total_rows = len(used[used["product"] == product])
        note = (f"{len(group)} خامة مكتوبة أكثر من مرة داخل نفس المنتج والقسم "
                f"(جُمعت معًا). الأكواد: {codes}.")
        if len(group) >= 5:
            note += (f" ⚠ المنتج له {total_rows} سطر خامات — يبدو أن قائمة الخامات "
                     "مسجلة مرتين كاملة (نسختان أو مقاسان). راجعها قبل الشراء "
                     "لأن الكميات قد تكون مضاعفة.")
        rows.append(["خامات مكررة داخل المنتج", product, note])

    return pd.DataFrame(rows, columns=["النوع", "البند", "التفاصيل"])



# --------------------------------------------------------------------------
# حساب القشرة (ملف منفصل: تابة أعداد + تابة "احتياج القشرة" بنص حر)
# --------------------------------------------------------------------------
def parse_veneer(text):
    """تحويل نص مثل «2 متر ارو مسنن» أو «نص متر ارو مفجر» إلى (النوع، الأمتار)."""
    value = re.sub(r"\s+", " ", clean_text(text))
    if not value:
        return "", None

    if "مسنن" in value:
        kind = "ارو مسنن"
    elif "مفجر" in value:
        kind = "ارو مفجر"
    elif "امريكي" in value:
        kind = "ارو امريكي"
    else:
        kind = ""

    number = re.match(r"^([\d.]+)", value)
    if number:
        return kind, float(number.group(1))
    for word, amount in QTY_WORDS.items():
        if value.startswith(word):
            return kind, amount
    if value.startswith("متر"):  # «متر ارو مسنن» تعني مترًا واحدًا
        return kind, 1.0
    return kind, None


def load_veneer(path, main_counts):
    """قراءة ملف القشرة وربط كل منتج بعدده. main_counts احتياطي للأسماء الناقصة."""
    need = pd.read_excel(path, sheet_name=VEN_SHEET_NEED).iloc[:, :2]
    need.columns = ["product_raw", "veneer_text"]
    need = need[need["product_raw"].notna()].copy()

    own = pd.read_excel(path, sheet_name=VEN_SHEET_COUNTS).iloc[:, :2]
    own.columns = ["product", "count"]
    own = own[own["product"].notna()]
    own_counts = {match_key(p): (clean_text(p), c)
                  for p, c in zip(own["product"], own["count"])}
    fallback = {match_key(p): (clean_text(p), c)
                for p, c in zip(main_counts["product"], main_counts["count"])}

    rows = []
    for _, r in need.iterrows():
        raw = clean_text(r["product_raw"])
        kind, per_unit = parse_veneer(r["veneer_text"])
        key = match_key(VENEER_ALIASES.get(match_key(raw), raw))

        if key in own_counts:
            matched, count, source = own_counts[key][0], own_counts[key][1], VEN_COUNTS_LABEL
        elif key in fallback:
            matched, count, source = fallback[key][0], fallback[key][1], "تابة الأعداد الرئيسية"
        else:
            matched, count, source = "", 0, "غير موجود"

        rows.append({
            "product_raw": raw,
            "veneer_text": clean_text(r["veneer_text"]),
            "product": matched,
            "kind": kind or UNSPECIFIED,
            "per_unit": per_unit,
            "count": pd.to_numeric(count, errors="coerce") or 0,
            "count_source": source,
        })

    veneer = pd.DataFrame(rows)
    veneer["per_unit"] = pd.to_numeric(veneer["per_unit"], errors="coerce")
    veneer["total_m"] = veneer["per_unit"].fillna(0) * veneer["count"]
    return veneer, own


def same_veneer_kind(new_kind, bom_kind):
    """هل النوعان نفس الخامة؟ على افتراض أن «ارو مسنن» هو «ارو امريكي» في القائمة القديمة."""
    if bom_kind == "غير موجود":
        return "—"
    if "مفجر" in new_kind and "مفجر" in bom_kind:
        return "متطابق"
    if "مسنن" in new_kind and "امريكي" in bom_kind:
        return "متطابق"
    return "مختلف"


def load_veneer_tab(path):
    """قراءة تابة مدخلات القشرة من ملف أنتجه هذا السكربت (لتشغيل الملف على نفسه)."""
    src = pd.read_excel(path, sheet_name=IN_VENEER)
    veneer = pd.DataFrame({
        "product_raw": src["المنتج (كما ورد)"].map(clean_text),
        "veneer_text": src["النص الأصلي"].map(clean_text),
        "product": src["المنتج المطابق"].map(clean_text),
        "kind": src["نوع القشرة"].map(clean_text),
        "per_unit": pd.to_numeric(src["متر للوحدة"], errors="coerce"),
        "count": pd.to_numeric(src["العدد المطلوب"], errors="coerce").fillna(0),
        "count_source": src["مصدر العدد"].map(clean_text),
    })
    veneer["total_m"] = veneer["per_unit"].fillna(0) * veneer["count"]
    own = pd.DataFrame({"product": veneer["product"], "count": veneer["count"]})
    return veneer, own[own["product"] != ""]


def build_veneer_tables(veneer, detail):
    """إجمالي كل نوع قشرة + مقارنة مع الأمتار المستخرجة من قائمة الخامات."""
    totals = (
        veneer.groupby("kind", as_index=False)
        .agg(total_m=("total_m", "sum"),
             products=("product_raw", "count"),
             pieces=("count", "sum"))
        .sort_values("total_m", ascending=False)
    )

    bom = detail[detail["mat_name_std"].str.contains("قشرة", na=False)]
    bom_by_product = (
        bom.groupby("product")
        .agg(bom_m=("total_qty", "sum"),
             bom_kind=("mat_name_std", lambda s: " + ".join(sorted(set(s)))))
    )

    compare = []
    for _, r in veneer.iterrows():
        if not r["product"]:
            continue
        key = match_key(r["product"])
        hit = [i for i in bom_by_product.index if match_key(i) == key]
        bom_m = float(bom_by_product.loc[hit[0], "bom_m"]) if hit else None
        bom_kind = bom_by_product.loc[hit[0], "bom_kind"] if hit else "غير موجود"
        compare.append([r["product"], r["kind"], round(float(r["total_m"]), 2),
                        bom_kind, round(bom_m, 2) if bom_m is not None else "—",
                        round(float(r["total_m"]) - bom_m, 2) if bom_m is not None else "—",
                        same_veneer_kind(r["kind"], bom_kind)])
    compare = pd.DataFrame(compare, columns=["المنتج", "نوع القشرة (هذا الملف)",
                                             "أمتار (هذا الملف)", "القشرة في قائمة الخامات",
                                             "أمتار (قائمة الخامات)", "الفرق", "توافق النوع"])
    return totals, compare


def veneer_issues(veneer, counts):
    rows = []
    for _, r in veneer.iterrows():
        if r["per_unit"] is None or pd.isna(r["per_unit"]):
            rows.append(["القشرة: كمية غير مقروءة", r["product_raw"],
                         f"النص «{r['veneer_text']}» لا يبدأ برقم — لم يُحسب."])
        if r["kind"] == UNSPECIFIED:
            rows.append(["القشرة: نوع غير معروف", r["product_raw"],
                         f"النص «{r['veneer_text']}» لا يذكر مسنن ولا مفجر."])
        if not r["product"]:
            rows.append(["القشرة: منتج بلا عدد", r["product_raw"],
                         "غير موجود في تابة الأعداد — حُسب بعدد صفر."])
        elif r["count_source"] != VEN_COUNTS_LABEL:
            rows.append(["القشرة: العدد من مصدر آخر", r["product_raw"],
                         f"غير موجود في تابة أعداد ملف القشرة؛ أُخذ العدد "
                         f"({r['count']:g}) من {r['count_source']} باسم «{r['product']}»."])
        if match_key(r["product_raw"]) in ("odeno middle table set",):
            rows.append(["القشرة: سطر يغطي مقاسين", r["product_raw"],
                         f"سطر واحد للمقاسين. حُسب على أنه {r['count']:g} طقم × "
                         f"{r['per_unit']:g} متر = {r['total_m']:g} متر. لو المقصود "
                         f"{r['per_unit']:g} متر لكل مقاس على حدة فالرقم يتضاعف إلى "
                         f"{r['total_m'] * 2:g} متر."])

    listed = {match_key(p) for p in veneer["product"] if p}
    for product, count in zip(counts["product"], counts["count"]):
        if count > 0 and match_key(product) not in listed:
            rows.append(["القشرة: منتج بلا سطر قشرة", clean_text(product),
                         f"مطلوب {count:g} قطعة ولا يوجد له سطر في تابة «{VEN_SHEET_NEED}» "
                         "— اعتُبر بلا قشرة."])
    return pd.DataFrame(rows, columns=["النوع", "البند", "التفاصيل"])


# --------------------------------------------------------------------------
# كتابة ملف الإكسل
# --------------------------------------------------------------------------
THIN = Side(style="thin", color="D0D0D0")
BORDER = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)
HEAD_FILL = PatternFill("solid", fgColor="1F4E79")
IN_HEAD_FILL = PatternFill("solid", fgColor="7F6000")  # تابات المدخلات بلون مختلف
HEAD_FONT = Font(bold=True, color="FFFFFF", size=11)
TITLE_FONT = Font(bold=True, size=14, color="1F4E79")


def write_sheet(wb, title, headers, rows, widths, number_cols=(), table_name=None,
                head_fill=HEAD_FILL):
    ws = wb.create_sheet(title)
    ws.sheet_view.rightToLeft = True
    ws.append(headers)
    for cell in ws[1]:
        cell.fill = head_fill
        cell.font = HEAD_FONT
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    ws.row_dimensions[1].height = 30

    for row in rows:
        ws.append(row)

    for idx, width in enumerate(widths, start=1):
        ws.column_dimensions[get_column_letter(idx)].width = width

    for row in ws.iter_rows(min_row=2, max_row=ws.max_row, max_col=len(headers)):
        for cell in row:
            cell.border = BORDER
            cell.alignment = Alignment(horizontal="center", vertical="center")
            if cell.column in number_cols:
                cell.number_format = "#,##0.00"

    if ws.max_row > 1:
        ws.freeze_panes = "A2"
        ws.auto_filter.ref = f"A1:{get_column_letter(len(headers))}{ws.max_row}"
        if table_name:
            table = Table(displayName=table_name,
                          ref=f"A1:{get_column_letter(len(headers))}{ws.max_row}")
            table.tableStyleInfo = TableStyleInfo(
                name="TableStyleMedium2", showRowStripes=True)
            # الفلتر التلقائي والجدول لا يجتمعان في نفس النطاق
            ws.auto_filter.ref = None
            ws.add_table(table)
    return ws


def write_summary(wb, df, detail, by_dept, grand, issues, counts, veneer_pack=None):
    ws = wb.create_sheet("ملخص", 0)
    ws.sheet_view.rightToLeft = True
    ws.sheet_view.showGridLines = False

    used = df[df["product_count"] > 0]
    products = used["product"].nunique()
    pieces = (
        used.groupby("product")["product_count"].first().sum()
    )

    def title(text, row):
        ws.cell(row=row, column=1, value=text).font = TITLE_FONT

    def head(row, headers):
        for col, text in enumerate(headers, start=1):
            cell = ws.cell(row=row, column=col, value=text)
            cell.fill = HEAD_FILL
            cell.font = HEAD_FONT
            cell.alignment = Alignment(horizontal="center", vertical="center")

    ws["A1"] = "ملخص احتياجات الخامات للشراء"
    ws["A1"].font = Font(bold=True, size=18, color="1F4E79")

    r = 3
    title("أرقام سريعة", r); r += 1
    quick = [
        ("عدد المنتجات المطلوبة", products),
        ("إجمالي عدد القطع المطلوب إنتاجها", int(pieces)),
        ("عدد أصناف الخامات المختلفة", grand["mat_key"].nunique()),
        ("عدد الأقسام المستهلكة", by_dept["dept"].nunique()),
        ("عدد أسطر قائمة الشراء التفصيلية", len(by_dept)),
        ("عدد الملاحظات التي تحتاج مراجعة", len(issues)),
    ]
    for label, value in quick:
        ws.cell(row=r, column=1, value=label).font = Font(bold=True)
        cell = ws.cell(row=r, column=2, value=value)
        cell.number_format = "#,##0"
        cell.alignment = Alignment(horizontal="center")
        r += 1

    r += 2
    title("الاحتياج حسب القسم", r); r += 1
    head(r, ["القسم", "عدد أصناف الخامات", "عدد المنتجات التي يخدمها"]); r += 1
    dept_stats = (
        used.groupby("dept")
        .agg(items=("mat_key", pd.Series.nunique), prods=("product", pd.Series.nunique))
        .sort_values("items", ascending=False)
    )
    for dept, row in dept_stats.iterrows():
        ws.cell(row=r, column=1, value=dept)
        ws.cell(row=r, column=2, value=int(row["items"])).alignment = Alignment(horizontal="center")
        ws.cell(row=r, column=3, value=int(row["prods"])).alignment = Alignment(horizontal="center")
        for col in range(1, 4):
            ws.cell(row=r, column=col).border = BORDER
        r += 1

    r += 2
    title("إجمالي الكميات حسب وحدة القياس", r); r += 1
    head(r, ["الوحدة", "إجمالي الكمية", "عدد الأصناف"]); r += 1
    unit_stats = (
        grand.groupby("unit")
        .agg(qty=("total_qty", "sum"), items=("mat_key", pd.Series.nunique))
        .sort_values("qty", ascending=False)
    )
    for unit, row in unit_stats.iterrows():
        ws.cell(row=r, column=1, value=unit)
        cell = ws.cell(row=r, column=2, value=float(row["qty"]))
        cell.number_format = "#,##0.00"
        cell.alignment = Alignment(horizontal="center")
        ws.cell(row=r, column=3, value=int(row["items"])).alignment = Alignment(horizontal="center")
        for col in range(1, 4):
            ws.cell(row=r, column=col).border = BORDER
        r += 1

    if veneer_pack:
        _, totals, _ = veneer_pack
        r += 2
        title("احتياج القشرة حسب النوع", r); r += 1
        head(r, ["نوع القشرة", "إجمالي الأمتار", "عدد المنتجات"]); r += 1
        for _, row in totals.iterrows():
            ws.cell(row=r, column=1, value=row["kind"])
            cell = ws.cell(row=r, column=2, value=round(float(row["total_m"]), 2))
            cell.number_format = "#,##0.00"
            cell.alignment = Alignment(horizontal="center")
            ws.cell(row=r, column=3, value=int(row["products"])).alignment = Alignment(
                horizontal="center")
            for col in range(1, 4):
                ws.cell(row=r, column=col).border = BORDER
            r += 1
        ws.cell(row=r, column=1, value="الإجمالي العام").font = Font(bold=True)
        cell = ws.cell(row=r, column=2, value=round(float(totals["total_m"].sum()), 2))
        cell.number_format = "#,##0.00"
        cell.font = Font(bold=True)
        cell.alignment = Alignment(horizontal="center")
        r += 1

    r += 2
    title("أعلى 15 خامة من حيث الكمية المطلوبة", r); r += 1
    head(r, ["اسم الخامة", "الوحدة", "الكمية الإجمالية"]); r += 1
    for _, row in grand.head(15).iterrows():
        ws.cell(row=r, column=1, value=row["mat_name_std"])
        ws.cell(row=r, column=2, value=row["unit"]).alignment = Alignment(horizontal="center")
        cell = ws.cell(row=r, column=3, value=float(row["total_qty"]))
        cell.number_format = "#,##0.00"
        cell.alignment = Alignment(horizontal="center")
        for col in range(1, 4):
            ws.cell(row=r, column=col).border = BORDER
        r += 1

    r += 2
    title("كيف تُقرأ التابات", r); r += 1
    notes = [
        "الإجمالي: كل خامة مرة واحدة — الكمية الكلية المطلوب شراؤها للمصنع كله.",
        "الإجمالي لكل قسم: نفس الكميات موزّعة على الأقسام (قائمة شراء كل قسم).",
        "فرعي - تفصيل المنتجات: أصل الحساب — كل سطر منتج × خامة (الكمية للوحدة × العدد).",
        "ملاحظات ومشاكل البيانات: أسطر ناقصة أو متعارضة في البيانات تحتاج مراجعة.",
        f"{IN_MATERIALS}: كل أسطر الخامات لكل منتج — هنا تضيف منتجًا أو تعدّل كمية.",
        f"{IN_COUNTS}: العدد المطلوب من كل منتج — غيّر الرقم فقط. العدد صفر = لا يُحسب.",
        "",
        "طريقة الحساب: الكمية الإجمالية = الكمية للوحدة × العدد المطلوب من المنتج.",
        "الكميات لا تُجمع عبر وحدات مختلفة (متر لا يُجمع مع كيلو) — كل وحدة في سطر مستقل.",
        "التابتان البنّيتان مدخلات تُعدَّل، والزرقاء نتائج محسوبة — لا تكتب فيها.",
        "بعد أي تعديل على تابتي المدخلات، شغّل السكربت على نفس الملف لتحديث النتائج.",
    ]
    for note in notes:
        if note:
            ws.cell(row=r, column=1, value="• " + note)
        r += 1

    ws.column_dimensions["A"].width = 65
    ws.column_dimensions["B"].width = 22
    ws.column_dimensions["C"].width = 26
    return ws


def write_veneer_sheets(wb, veneer, totals, compare):
    """ثلاث تابات للقشرة: الإجمالي لكل نوع، تفصيل المنتجات، ومقارنة بقائمة الخامات."""
    grand = float(totals["total_m"].sum())
    write_sheet(
        wb, "القشرة - الإجمالي",
        ["نوع القشرة", "إجمالي الأمتار المطلوبة", "عدد المنتجات", "إجمالي القطع"],
        [[r["kind"], round(float(r["total_m"]), 2), int(r["products"]), int(r["pieces"])]
         for _, r in totals.iterrows()] +
        [["الإجمالي العام", round(grand, 2), int(totals["products"].sum()), ""]],
        [22, 26, 15, 15], number_cols={2},
    )
    last = wb["القشرة - الإجمالي"].max_row
    for cell in wb["القشرة - الإجمالي"][last]:
        cell.font = Font(bold=True)
        cell.fill = PatternFill("solid", fgColor="DDEBF7")

    write_sheet(
        wb, "القشرة - تفصيل المنتجات",
        ["المنتج", "نوع القشرة", "متر للوحدة", "العدد المطلوب",
         "إجمالي الأمتار", "النص الأصلي"],
        [[r["product"] or r["product_raw"], r["kind"],
          "—" if pd.isna(r["per_unit"]) else float(r["per_unit"]),
          int(r["count"]), round(float(r["total_m"]), 2), r["veneer_text"]]
         for _, r in veneer.sort_values(["kind", "total_m"], ascending=[True, False]).iterrows()],
        [32, 16, 14, 15, 16, 24], number_cols={3, 5}, table_name="VeneerDetail",
    )

    write_sheet(
        wb, "القشرة - مقارنة بقائمة الخامات",
        list(compare.columns), compare.values.tolist(),
        [32, 20, 18, 34, 20, 12, 14], number_cols={3, 5, 6}, table_name="VeneerCompare",
    )


def write_inputs(wb, df, counts):
    """تابتا المدخلات: كل أسطر الخامات والأعداد — قابلة للتعديل وإعادة التشغيل."""
    rows = df.sort_values(["product", "dept", "mat_name"], kind="stable")
    write_sheet(
        wb, IN_MATERIALS, IN_MATERIALS_COLS,
        [[r["product"], r["model"], r["product_code"], r["dept"], r["mat_code"],
          r["mat_name"], float(r["qty_per_unit"]), r["unit"], str(r["source_row"])]
         for _, r in rows.iterrows()],
        [30, 30, 16, 22, 14, 40, 14, 14, 26],
        number_cols={7}, table_name="InputMaterials", head_fill=IN_HEAD_FILL,
    )

    known = {match_key(p) for p in counts["product"]}  # noqa: E501
    extra = [(p, 0) for p in sorted(set(df["product"])) if match_key(p) not in known]
    write_sheet(
        wb, IN_COUNTS, IN_COUNTS_COLS,
        [[p, int(c)] for p, c in zip(counts["product"], counts["count"])] +
        [[p, c] for p, c in extra],
        [40, 14], number_cols={2}, table_name="InputCounts", head_fill=IN_HEAD_FILL,
    )


def write_veneer_input(wb, veneer):
    write_sheet(
        wb, IN_VENEER, IN_VENEER_COLS,
        [[r["product_raw"], r["veneer_text"], r["product"], r["kind"],
          "" if pd.isna(r["per_unit"]) else float(r["per_unit"]),
          int(r["count"]), r["count_source"]]
         for _, r in veneer.iterrows()],
        [30, 22, 30, 16, 14, 15, 26],
        number_cols={5}, table_name="InputVeneer", head_fill=IN_HEAD_FILL,
    )


def write_output(path, df, detail, by_dept, grand, issues, counts, veneer_pack=None):
    wb = Workbook()
    wb.remove(wb.active)

    write_sheet(
        wb, "الإجمالي",
        ["كود الخامة", "اسم الخامة", "الوحدة", "الكمية الإجمالية المطلوبة",
         "عدد المنتجات", "الأقسام المستهلكة"],
        [[r["mat_code"] or "—", r["mat_name_std"], r["unit"], round(float(r["total_qty"]), 3),
          int(r["products"]), r["depts"]] for _, r in grand.iterrows()],
        [14, 42, 14, 24, 13, 46], number_cols={4}, table_name="TotalMaterials",
    )

    write_sheet(
        wb, "الإجمالي لكل قسم",
        ["القسم", "كود الخامة", "اسم الخامة", "الوحدة",
         "الكمية المطلوبة", "عدد المنتجات"],
        [[r["dept"], r["mat_code"] or "—", r["mat_name_std"], r["unit"],
          round(float(r["total_qty"]), 3), int(r["products"])]
         for _, r in by_dept.iterrows()],
        [24, 14, 42, 14, 18, 13], number_cols={5}, table_name="DeptMaterials",
    )

    write_sheet(
        wb, "فرعي - تفصيل المنتجات",
        ["المنتج", "الموديل", "كود المنتج", "العدد المطلوب", "القسم",
         "كود الخامة", "اسم الخامة", "الوحدة", "الكمية للوحدة", "الكمية الإجمالية"],
        [[r["product"], r["model"], r["product_code"] or "—", int(r["product_count"]),
          r["dept"], r["mat_code"] or "—", r["mat_name_std"], r["unit"],
          round(float(r["qty_per_unit"]), 4), round(float(r["total_qty"]), 3)]
         for _, r in detail.iterrows()],
        [30, 28, 14, 13, 22, 13, 40, 13, 14, 16],
        number_cols={9, 10}, table_name="DetailLines",
    )

    write_sheet(
        wb, "ملاحظات ومشاكل البيانات",
        ["النوع", "البند", "التفاصيل"],
        issues.values.tolist(),
        [28, 40, 90], table_name="DataIssues",
    )
    for row in wb["ملاحظات ومشاكل البيانات"].iter_rows(min_row=2, min_col=3, max_col=3):
        for cell in row:
            cell.alignment = Alignment(horizontal="right", vertical="center", wrap_text=True)

    if veneer_pack:
        veneer, totals, compare = veneer_pack
        write_veneer_sheets(wb, veneer, totals, compare)

    write_inputs(wb, df, counts)
    if veneer_pack:
        write_veneer_input(wb, veneer_pack[0])
    write_summary(wb, df, detail, by_dept, grand, issues, counts, veneer_pack)
    wb.save(path)


def parse_args(argv):
    positional, options = [], {}
    i = 0
    while i < len(argv):
        if argv[i] in ("--extra-bom", "--extra-counts", "--veneer"):
            if i + 1 >= len(argv):
                raise SystemExit(f"{argv[i]} يحتاج مسار ملف بعده.")
            options[argv[i]] = argv[i + 1]
            i += 2
        else:
            positional.append(argv[i])
            i += 1
    return positional, options


def main():
    positional, options = parse_args(sys.argv[1:])
    if not positional:
        raise SystemExit(__doc__)
    src = positional[0]
    dst = positional[1] if len(positional) > 1 else "احتياجات_الخامات.xlsx"

    materials, counts = load_input(src)
    extra_rows, extra_counts = load_extra(options.get("--extra-bom"),
                                          options.get("--extra-counts"))
    if extra_rows is not None:
        materials = pd.concat([materials, extra_rows], ignore_index=True)
    if extra_counts is not None:
        counts = pd.concat([counts, extra_counts], ignore_index=True)

    df = prepare(materials, counts)
    detail, by_dept, grand = build_tables(df)
    issues = build_issues(df, counts)

    veneer_pack = None
    veneer_path = options.get("--veneer")
    has_tab = IN_VENEER in pd.ExcelFile(src).sheet_names
    if veneer_path or has_tab:
        if veneer_path:
            veneer, own_counts = load_veneer(veneer_path, counts)
        else:
            veneer, own_counts = load_veneer_tab(src)
        totals, compare = build_veneer_tables(veneer, detail)
        issues = pd.concat([issues, veneer_issues(veneer, counts)], ignore_index=True)
        veneer_pack = (veneer, totals, compare)

    write_output(dst, df, detail, by_dept, grand, issues, counts, veneer_pack)

    used = df[df["product_count"] > 0]
    print(f"تم إنشاء: {dst}")
    print(f"  منتجات محسوبة      : {used['product'].nunique()}")
    print(f"  أصناف خامات        : {grand['mat_key'].nunique()}")
    print(f"  أسطر تفصيلية       : {len(detail)}")
    print(f"  أسطر لكل قسم       : {len(by_dept)}")
    print(f"  ملاحظات للمراجعة   : {len(issues)}")
    if veneer_pack:
        for _, row in veneer_pack[1].iterrows():
            print(f"  قشرة {row['kind']:<10}: {row['total_m']:,.2f} متر")


if __name__ == "__main__":
    main()
