# -*- coding: utf-8 -*-
"""解析《挥发性风味物质_完整版_阈值_CAS对照表.html》，
转成与 curated_extras.json 同源的 JSON，再并入主数据库。

本次重点补充：
  - 「主要来源」列（肉类高汤/乳制品=主要来源；茶汤=主要茶类；柑橘=主要柑橘）→ 字段 dist
  - 细分类别（含硫化合物 / 吡嗪类 / 内酯类 等），依据 h3 小节与 tr.cat 子类别推断
字段：en/cn/cas/cat/thr/med/note(气味描述)/dist(主要来源)/subcat/src/syn
"""
import json, re, os
from bs4 import BeautifulSoup
import importlib.util

HERE = os.path.dirname(os.path.abspath(__file__))
spec = importlib.util.spec_from_file_location("fc", os.path.join(HERE, "flavor_core.py"))
fc = importlib.util.module_from_spec(spec); spec.loader.exec_module(fc)

SRC = "/root/uploads/1787656343439922873-挥发性风味物质_完整版_阈值_CAS对照表.html"
OUT = os.path.join(HERE, "volatile_flavor_html.json")

html = open(SRC, encoding="utf-8").read()
soup = BeautifulSoup(html, "html.parser")


def main_name(en):
    en = en.strip()
    return re.sub(r"\s*\(.*?\)\s*", " ", en).strip()


def clean_cn(cn):
    return cn.replace("（", "(").replace("）", ")").strip()


def parse_thr(thr_text):
    """从阈值单元格提取 (thr_水相, med_完整)。水相优先；无数值记为 '—'。"""
    t = thr_text.strip()
    if not t or t == "—":
        return "—", "—"
    segs = re.split(r"[；;]", t)
    water_seg = None
    other_seg = None
    for s in segs:
        s = s.strip()
        if not s:
            continue
        if "水" in s or "ppb" in s.lower() or "μg/kg" in s or "ppm" in s.lower():
            if "空气" in s and "水" not in s:
                other_seg = other_seg or s
                continue
            water_seg = water_seg or s
        else:
            other_seg = other_seg or s
    chosen = water_seg or other_seg
    if chosen is None:
        return "—", t
    m = re.search(r"[-+]?[\d.]+(?:[eE][-+]?\d+)?\s*[^；;（）()\s]*", chosen)
    val = m.group(0).strip() if m else chosen
    return val, t


SECTION_HINT = ['肉类高汤', '乳制品', '茶汤', '柑橘']


def subcat_to_cat(h):
    """由 h3 小节标题或 tr.cat 子类别标题推断细分类别。"""
    if not h:
        return None
    if '含硫' in h:
        return '含硫化合物'
    if '吡嗪' in h:
        return '吡嗪类'
    if '内酯' in h:
        return '内酯类'
    if '酮' in h:
        return '酮类'
    if '醇' in h:
        return '醇类'
    if '酯' in h:
        return '酯类'
    if '醛' in h:
        return '醛类'
    if '萜' in h:
        return '萜烯类'
    # 下列为混合组，按单个物质名自动判别
    if '酸' in h or '发酵产物' in h or 'C6 醛醇' in h or '其他重要香气' in h:
        return None
    return None


# 当前 h2 段 -> 规范食品类别（仅 4 个食品段用于「主要来源」分组；其余段忽略）
SECTION_CANON = {'肉类高汤': '肉类', '乳制品': '乳制品', '茶汤': '茶类', '柑橘': '柑橘类'}


def canon_section(h2):
    for s in SECTION_HINT:
        if s in h2:
            return SECTION_CANON.get(s)
    return None


agg = {}          # en(归一) -> 合并记录
order = []        # 保持出现顺序
current_h2 = ''
current_sec = None
current_h3 = ''
current_subcat = ''

for el in soup.find_all(['h2', 'h3', 'tr']):
    if el.name == 'h2':
        current_h2 = el.get_text(strip=True)
        for s in SECTION_HINT:
            if s in current_h2:
                current_h2 = s
                break
        current_sec = canon_section(current_h2)
        current_h3 = ''
        current_subcat = ''
        continue
    if el.name == 'h3':
        current_h3 = el.get_text(strip=True)
        current_subcat = ''
        continue
    if el.name == 'tr':
        tds = el.find_all('td')
        if len(tds) < 6:
            # 可能是 tr.cat 子类别分隔行（1 个 colspan 单元格）
            if len(tds) == 1:
                txt = tds[0].get_text(strip=True)
                if txt and not txt.startswith('主要'):
                    current_subcat = txt
            continue
        cn = clean_cn(tds[0].get_text(strip=True))
        en = tds[1].get_text(strip=True).strip()
        if not cn or not en:
            continue
        # 跳过可能的非物质行
        if cn in ('茶类', '特征维度', '关键化合物'):
            continue
        # 仅食品段（4 类）参与「主要来源」聚合
        if current_sec is None:
            continue
        cas = tds[2].get_text(strip=True)
        cas = "" if (cas.startswith("—") or cas == "") else cas.strip()
        thr_text = tds[3].get_text(strip=True)
        fla = tds[4].get_text(strip=True).strip()
        dist = tds[5].get_text(strip=True).strip()   # 该食品段的主要来源
        mn = main_name(en)
        syn = [mn.lower()]
        for p in re.findall(r"\(([^)]*)\)", en):
            for x in re.split(r"[/,]", p):
                x = x.strip().lower()
                if x and x not in syn:
                    syn.append(x)
        thr_val, med_full = parse_thr(thr_text)
        cat = subcat_to_cat(current_h3)
        if cat is None and current_subcat:
            cat = subcat_to_cat(current_subcat)
        if cat is None:
            cat = fc.auto_classify(mn)
        # 聚合：同一物质跨食品段合并 dist_by_cat
        rec = agg.get(mn)
        if rec is None:
            rec = {
                "en": mn, "cn": cn, "cas": cas, "cat": cat,
                "thr": thr_val, "med": med_full, "note": fla,
                "dist_by_cat": {}, "subcat": current_subcat or current_h3,
                "src": "HTML表(风味物质完整版)", "syn": syn,
            }
            agg[mn] = rec
            order.append(mn)
        # 阈值/描述：首个有效值优先
        if not rec["thr"] or rec["thr"] == "—":
            rec["thr"] = thr_val
        if not rec["med"] or rec["med"] == "—":
            rec["med"] = med_full
        if not rec["note"] or rec["note"] == "—":
            rec["note"] = fla
        # 类别：取最细（非其他优先；内酯类从酯类析出）
        if cat and cat != '其他' and (rec["cat"] in (None, '', '其他') or (rec["cat"] == '酯类' and cat == '内酯类')):
            rec["cat"] = cat
        # 主要来源：按段并入 dist_by_cat
        if dist and dist != '—':
            lst = rec["dist_by_cat"].setdefault(current_sec, [])
            for p in re.split(r'[、，,/／；;]', dist):
                p = p.strip()
                if p and p not in lst:
                    lst.append(p)

rows = []
for mn in order:
    rec = agg[mn]
    # dist_by_cat: {肉类:'肉类、油脂氧化', ...}；dist 保留并集（向后兼容）
    dist_by_cat = {k: '、'.join(v) for k, v in rec["dist_by_cat"].items()}
    dist_union = '；'.join(dist_by_cat[k] for k in ['肉类', '乳制品', '茶类', '柑橘类'] if k in dist_by_cat)
    item = {
        "en": rec["en"], "cn": rec["cn"], "cas": rec["cas"], "cat": rec["cat"],
        "thr": rec["thr"], "med": rec["med"], "note": rec["note"],
        "dist": dist_union, "dist_by_cat": dist_by_cat,
        "subcat": rec["subcat"], "src": rec["src"], "syn": rec["syn"],
    }
    rows.append(item)

out = {"_doc": "挥发性风味物质完整版（肉类高汤/乳制品/茶汤/柑橘）阈值CAS对照表，解析自 HTML，含主要来源(按食品段分组 dist_by_cat)与细分类别。", "items": rows}
json.dump(out, open(OUT, "w", encoding="utf-8"), ensure_ascii=False, indent=2)

# 类别统计
from collections import Counter
c = Counter(r['cat'] for r in rows)
print("生成:", OUT, "条目:", len(rows))
print("--- 类别分布 ---")
for k, v in c.most_common():
    print("  %s: %d" % (k, v))
