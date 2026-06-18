# 從業績總表 Excel 產生 data.js（業務 × 客戶 × 產品系列 明細）
# 用法：python convert_xlsx.py "G:\My Drive\業績\2026\新業代業績總表(含target)_20260528 (忠季).xlsx"
import sys, json, re, os, datetime
import pandas as pd

path = sys.argv[1]
DIVS = ["Div1", "Div2", "Div3"]

# 年度與資料截止日：自動從檔名的 8 位數日期抓取（例：..._20260528... → 2026 / 2026-05-28）
_m = re.findall(r"(20\d{6})", os.path.basename(path))
if _m:
    _d = _m[-1]
    YEAR = _d[0:4]
    UPDATED = f"{_d[0:4]}-{_d[4:6]}-{_d[6:8]}"
else:
    YEAR = str(datetime.date.today().year)
    UPDATED = datetime.date.today().isoformat()

df = pd.read_excel(path, sheet_name="業績 DATA", dtype={"年": str, "月": str})
d = df[(df["年"] == YEAR) & (df["處別"].isin(DIVS))].copy()
d["業績金額"] = pd.to_numeric(d["業績金額"], errors="coerce")
d["目標金額"] = pd.to_numeric(d["目標金額"], errors="coerce")
d["月n"] = pd.to_numeric(d["月"], errors="coerce")
d["客戶名稱"] = d["客戶名稱"].fillna("（未填客戶）").astype(str).str.strip()

dup = d.groupby("業代名稱")["處別"].nunique()
assert (dup <= 1).all(), "有業代出現在多個處別: " + str(dup[dup > 1].index.tolist())

last_month = int(d[d["業績金額"].notna()]["月n"].max())


def mode_or_blank(s):
    s = s.dropna().astype(str).str.strip()
    s = s[s != ""]
    return s.mode().iloc[0] if not s.empty else ""


people = []
for (grp, name), pers in d.groupby(["處別", "業代名稱"]):
    customers = []
    for cust, csub in pers.groupby("客戶名稱"):
        products = []
        for series, ps in csub.groupby("產品系列"):
            target = int(round(ps["目標金額"].sum()))
            monthly = []
            for m in range(1, 13):
                if m <= last_month:
                    monthly.append(int(round(ps.loc[ps["月n"] == m, "業績金額"].sum())))
                else:
                    monthly.append(None)
            if target == 0 and not any(monthly[:last_month]):
                continue  # 該系列無目標也無業績，略過
            products.append({"series": series, "yearTarget": target, "monthly": monthly})
        if not products:
            continue
        products.sort(key=lambda x: -(x["yearTarget"] or sum(v or 0 for v in x["monthly"])))
        customers.append({
            "name": cust,
            "channel": mode_or_blank(csub["通路"]),
            "system": mode_or_blank(csub["體系名稱"]),
            "city": mode_or_blank(csub["縣市"]),
            "products": products,
        })
    # 客戶依「年度目標＋累計業績」排序，讓重點客戶排前面
    customers.sort(key=lambda c: -sum(
        (x["yearTarget"] or 0) + sum(v or 0 for v in x["monthly"]) for x in c["products"]))
    people.append({"name": name, "group": grp, "customers": customers})

people.sort(key=lambda p: (p["group"], -sum(
    (x["yearTarget"] or 0) for c in p["customers"] for x in c["products"])))


def js_str(s):
    return json.dumps(s, ensure_ascii=False)


def js_product(x):
    monthly = ", ".join("null" if v is None else str(v) for v in x["monthly"])
    return ('          { series: %s, yearTarget: %d, monthly: [%s] }'
            % (js_str(x["series"]), x["yearTarget"], monthly))


def js_customer(c):
    prods = ",\n".join(js_product(x) for x in c["products"])
    return ('      { name: %s, channel: %s, system: %s, city: %s, products: [\n%s\n      ] }'
            % (js_str(c["name"]), js_str(c["channel"]), js_str(c["system"]),
               js_str(c["city"]), prods))


def js_person(p):
    custs = ",\n".join(js_customer(c) for c in p["customers"])
    return ('    { name: %s, group: %s, customers: [\n%s\n    ] }'
            % (js_str(p["name"]), js_str(p["group"]), custs))


header = """\
// ============================================================
// 客戶業績資料檔 — 由 convert_xlsx.py 從業績總表 Excel 自動產生
// 重新產生：python convert_xlsx.py "<Excel檔路徑>"
// ============================================================
//
// 結構說明：
//   每位業代底下依「客戶」拆分，每個客戶再依「產品系列」拆分：
//     channel    : 通路（醫學中心／區域醫院／地區醫院／診所…）
//     system     : 體系名稱   city: 縣市
//     yearTarget : 該客戶該系列年度目標（元）＝ 1~12 月目標金額加總
//     monthly    : 該客戶該系列 1~12 月實際業績（元），尚無資料的月份為 null
//   範圍：處別 Div1/Div2/Div3，不含直營與經銷
//
"""

body = ",\n".join(js_person(p) for p in people)
js = (header
      + "const CUSTOMER_DATA = {\n"
      + f"  year: {YEAR},\n"
      + f'  updatedAt: "{UPDATED}",\n'
      + "  salespeople: [\n" + body + "\n  ]\n};\n")

with open("data.js", "w", encoding="utf-8") as f:
    f.write(js)

n_cust = sum(len(p["customers"]) for p in people)
total_t = sum(x["yearTarget"] for p in people for c in p["customers"] for x in c["products"])
total_a = sum(v for p in people for c in p["customers"] for x in c["products"]
              for v in x["monthly"] if v)
print(f"people={len(people)} customers={n_cust} months=1-{last_month} "
      f"target={total_t:,} actual={total_a:,}")
