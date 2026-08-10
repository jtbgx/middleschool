# -*- coding: utf-8 -*-
import glob
import json
import re
import unicodedata
from collections import Counter

import openpyxl


NINE_DISTRICTS = [
    "上城区",
    "拱墅区",
    "西湖区",
    "滨江区",
    "钱塘区",
    "萧山区",
    "余杭区",
    "临平区",
    "富阳区",
]
REGION_PREFIX = [
    "杭州市",
    "杭州",
    "上城区",
    "拱墅区",
    "西湖区",
    "滨江区",
    "钱塘区",
    "萧山区",
    "富阳区",
    "临平区",
    "余杭区",
    "临安区",
    "桐庐县",
    "淳安县",
    "建德市",
]


def norm2(s):
    if not s:
        return ""
    s = unicodedata.normalize("NFKC", str(s))
    s = s.replace(" ", "").replace("\u3000", "")
    for p in REGION_PREFIX:
        if s.startswith(p):
            s = s[len(p):]
    s = re.sub(r"[（(][^）)]*[)）]", "", s)
    s = s.replace("初级中学", "初中")
    return s


def variants(s):
    base = norm2(s)
    out = {base}
    if base:
        out.add(base.replace("实验学校", "学校"))
        out.add(base.replace("中学教育集团", "中学"))
        out.add(base.replace("附属学校", "学校"))
        out.add(base.replace("初级中学", "初中"))
        out.add(base.replace("中学", "初中"))
    return out


def sim_match(n1, n2):
    if not n1 or not n2:
        return False
    if n1 == n2:
        return True
    v1, v2 = variants(n1), variants(n2)
    if v1 & v2:
        return True
    if len(n1) >= 4 and n1 in n2:
        return True
    if len(n2) >= 4 and n2 in n1:
        return True
    return False


def find_key(items, district, name):
    n = norm2(name)
    for k, _ in items:
        if k[0] == district and sim_match(k[1], n):
            return k
    return None


def is_school_name(name):
    if not name:
        return False
    if not re.search(r"[\u4e00-\u9fa5]", name):
        return False
    return any(w in name for w in ("中学", "学校", "教育集团"))


path = [p for p in glob.glob(r"E:\ai\codex-project\shengxuechuzhong\*.xlsx") if "分配生" in p][0]
wb = openpyxl.load_workbook(path, data_only=True, read_only=True)

# ---- 256 所名单骨架 ----
ws = wb.worksheets[14]
base_rows = []
for r in ws.iter_rows(min_col=1, max_col=6, values_only=True):
    seq = r[0]
    if seq is None or str(seq).strip() in ("", "序号"):
        continue
    district = str(r[1]).strip() if r[1] else ""
    name = str(r[2]).strip() if r[2] else ""
    if not name or not district or not is_school_name(name):
        continue
    base_rows.append(
        {
            "district": district,
            "name": name,
            "nature": str(r[3]).strip() if r[3] else "",
            "type": str(r[4]).strip() if r[4] else "",
            "note": str(r[5]).strip() if r[5] else "",
        }
    )

# ---- 分配生总表（6 区，T1-T5）----
ws = wb.worksheets[2]
sheet2_rows = []
cur_d = cur_t = None
for r in ws.iter_rows(min_col=1, max_col=11, values_only=True):
    seq = r[0]
    if seq is None or str(seq).strip() in ("", "序号"):
        continue
    d = str(r[1]).strip() if r[1] else ""
    t = str(r[2]).strip() if r[2] else ""
    name = str(r[3]).strip() if r[3] else ""
    if d:
        cur_d = d
    if t:
        m = re.search(r"\d+", t)
        cur_t = int(m.group()) if m else None
    if not name or not is_school_name(name):
        continue
    sheet2_rows.append(
        {
            "district": cur_d,
            "name": name,
            "tier": cur_t,
            "nature": str(r[4]).strip() if r[4] else "",
            "type": str(r[5]).strip() if r[5] else "",
            "cnt26": str(r[6]).strip() if r[6] is not None else "",
            "quota26": str(r[7]).strip() if r[7] is not None else "",
            "note": str(r[10]).strip() if len(r) > 10 and r[10] else "",
        }
    )

# ---- 拱墅区（梯队排名 1-4）----
ws = wb.worksheets[4]
sheet4_rows = []
cur_t = None
for r in ws.iter_rows(min_col=1, max_col=10, max_row=80, values_only=True):
    seq = r[0]
    name = str(r[3]).strip() if len(r) > 3 and r[3] else ""
    if str(seq).strip() == "序号":
        continue
    if not name or not is_school_name(name):
        continue
    t = str(r[2]).strip() if len(r) > 2 and r[2] else ""
    if t:
        m = re.search(r"\d+", t)
        cur_t = int(m.group()) if m else None
    if cur_t is not None and not (1 <= cur_t <= 5):
        print("WARN 忽略异常梯队值:", cur_t, name)
        cur_t = None
    sheet4_rows.append(
        {
            "district": "拱墅区",
            "name": name,
            "tier": cur_t,
            "nature": str(r[5]).strip() if len(r) > 5 and r[5] else "",
            "type": str(r[6]).strip() if len(r) > 6 and r[6] else "",
            "cnt26": str(r[8]).strip() if len(r) > 8 and r[8] is not None else "",
            "quota26": str(r[9]).strip() if len(r) > 9 and r[9] is not None else "",
            "note": str(r[7]).strip() if len(r) > 7 and r[7] else "",
        }
    )

# ---- 富阳区（名单与办学信息）----
ws = wb.worksheets[11]
sheet11_rows = []
for r in ws.iter_rows(min_col=1, max_col=6, max_row=27, values_only=True):
    seq = r[0]
    if seq is None or str(seq).strip() in ("", "序号"):
        continue
    name = str(r[2]).strip() if len(r) > 2 and r[2] else ""
    if not name or not is_school_name(name):
        continue
    sheet11_rows.append(
        {
            "district": "富阳区",
            "name": name,
            "tier": None,
            "nature": str(r[3]).strip() if len(r) > 3 and r[3] else "",
            "type": str(r[4]).strip() if len(r) > 4 and r[4] else "",
            "cnt26": "",
            "quota26": "",
            "note": str(r[5]).strip() if len(r) > 5 and r[5] else "",
        }
    )

# ---- 民间榜 125 所（按重高率分档补梯队）----
ws = wb.worksheets[13]
sheet125_rows = []
for r in ws.iter_rows(min_col=1, max_col=10, values_only=True):
    district = str(r[0]).strip() if r[0] else ""
    name = str(r[1]).strip() if len(r) > 1 and r[1] else ""
    if not district or district == "区域" or not name:
        continue
    if district not in NINE_DISTRICTS:
        continue
    try:
        rate = float(r[5]) / 100.0
    except (TypeError, ValueError):
        rate = None
    tier = None
    if rate is not None:
        if rate >= 0.40:
            tier = 1
        elif rate >= 0.30:
            tier = 2
        elif rate >= 0.20:
            tier = 3
        else:
            tier = 4
    sheet125_rows.append(
        {
            "district": district,
            "name": name,
            "tier": tier,
            "nature": str(r[2]).strip() if len(r) > 2 and r[2] else "",
            "type": "",
            "cnt26": "",
            "quota26": "",
            "note": str(r[9]).strip() if len(r) > 9 and r[9] else "",
        }
    )

# ---- 萧山工作表确认是否有初中名单 ----
ws = wb.worksheets[10]
cx_hits = []
for r in ws.iter_rows(min_col=1, max_col=10, max_row=236, values_only=True):
    joined = " ".join(str(c) for c in r if c is not None)
    if "初中" in joined:
        cx_hits.append([str(c) if c is not None else "" for c in r][:6])
print("萧山工作表含初中字样的行数:", len(cx_hits))
for row in cx_hits:
    print("   ", row)


# ---- 合并 ----
records_by_key = {}


def add(rec):
    k = (rec["district"], norm2(rec["name"]))
    old = records_by_key.get(k)
    if old:
        if old["tier"] is None and rec["tier"] is not None:
            old["tier"] = rec["tier"]
        for f in ("nature", "type", "cnt26", "quota26", "note"):
            if not old[f] and rec[f]:
                old[f] = rec[f]
        return
    records_by_key[k] = dict(rec)


def apply_source(rows):
    for item in rows:
        fk = find_key(list(records_by_key.items()), item["district"], item["name"])
        if fk:
            target = records_by_key[fk]
            if target["tier"] is None and item["tier"] is not None:
                target["tier"] = item["tier"]
            for f in ("nature", "type", "cnt26", "quota26", "note"):
                if not target[f] and item[f]:
                    target[f] = item[f]
        else:
            add(item)


for b in base_rows:
    add(
        {
            "district": b["district"],
            "tier": None,
            "name": b["name"],
            "nature": b["nature"],
            "type": b["type"],
            "cnt26": "",
            "quota26": "",
            "note": b["note"],
        }
    )

apply_source(sheet2_rows)
apply_source(sheet4_rows)
apply_source(sheet11_rows)
apply_source(sheet125_rows)

records = list(records_by_key.values())
SKIP_NAMES = {"采荷濮家"}
total_before = len(records)
records = [r for r in records if r["name"] not in SKIP_NAMES]
removed = total_before - len(records)
if removed:
    print("已移除重复校区条目:", removed)
records.sort(key=lambda x: (NINE_DISTRICTS.index(x["district"]), x["name"]))

by_d = Counter(x["district"] for x in records)
by_t = Counter(x["tier"] for x in records)
print("总校数:", len(records))
print("按区:", dict(by_d))
print("按梯队:", dict(by_t))
print("无梯队数量:", by_t.get(None, 0))
for x in records:
    if x["tier"] is None:
        print("  无梯队:", x["district"], x["name"])

# ---- 输出网页数据 ----
out_js = r"E:\ai\codex-project\shengxuechuzhong\data.js"
with open(out_js, "w", encoding="utf-8") as f:
    f.write("window.HZ_SCHOOLS = ")
    json.dump(records, f, ensure_ascii=False, indent=1)
    f.write(";\n")
print("written:", out_js)

# ---- 输出整理版 Excel ----
out_xlsx = r"E:\ai\codex-project\shengxuechuzhong\杭州9区初中梯队总表_整理版.xlsx"
wb2 = openpyxl.Workbook()
ws2 = wb2.active
ws2.title = "9区初中梯队总表"
headers = ["序号", "所属区域", "梯队", "学校名称", "办学性质", "办学类型", "2026报考人数", "2026分配生名额", "备注"]
ws2.append(headers)
for i, rec in enumerate(records, 1):
    tier_text = "T%d" % rec["tier"] if rec["tier"] else "暂无"
    ws2.append(
        [
            i,
            rec["district"],
            tier_text,
            rec["name"],
            rec["nature"],
            rec["type"],
            rec["cnt26"],
            rec["quota26"],
            rec["note"],
        ]
    )
widths = [6, 10, 8, 44, 16, 26, 12, 14, 46]
for idx, w in enumerate(widths, 1):
    ws2.column_dimensions[openpyxl.utils.get_column_letter(idx)].width = w
wb2.save(out_xlsx)
print("written:", out_xlsx)
wb.close()
