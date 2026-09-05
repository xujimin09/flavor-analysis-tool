# -*- coding: utf-8 -*-
"""
把 FlavorThresholdDB (XQplayer/HXQLab, v1.5.0) 的 aroma_data_merged.json
融入本程序：生成 ftdb_compounds.json（映射到本程序的化合物 schema）。
- 以 CAS 去重，约 2207 个唯一化合物
- 化学类别用 flavor_core.auto_classify 自动归类（醛/酮/酯/酸/醇/萜烯/其他）
- 阈值优先「水相」(μg/kg≈ppb, 兼容 OAV)，仅空气/其他介质的阈值放备注(标注单位, 不计入OAV)
- 提取括号内同义名作为 syn；保留中文名、CAS、风味类别与风味描述
"""
import json, re, os
import importlib.util

HERE = os.path.dirname(os.path.abspath(__file__))
# 载入 auto_classify
spec = importlib.util.spec_from_file_location("fc", os.path.join(HERE, "flavor_core.py"))
fc = importlib.util.module_from_spec(spec); spec.loader.exec_module(fc)

SRC = "/tmp/aroma_merged.json"
OUT = os.path.join(HERE, "ftdb_compounds.json")

def base_name(en):
    en = (en or "").strip()
    return re.sub(r'\s*\(.*?\)\s*', ' ', en).strip()

def synonyms(en):
    out = []
    for m in re.findall(r'\(([^)]*)\)', en or ''):
        for part in re.split(r'[,;]', m):
            p = part.strip().lower()
            if p:
                out.append(p)
    out.append(base_name(en).lower())
    return out

def fix_space_dec(s):
    # FTDB 小数值用空格作小数分隔: "0.000 02" -> "0.00002"
    return re.sub(r'(\d)\.(\d+)\s+(\d{1,4})', lambda m: m.group(1)+'.'+m.group(2)+m.group(3), s)

def extract_value(s):
    s = fix_space_dec(s)
    if '⇒' in s or '=>' in s:          # 交叉引用行, 无本物阈值
        return None
    toks = s.split()
    nums = [(i, t) for i, t in enumerate(toks)
            if re.fullmatch(r'\d+\.\d+|\d+', t) and not (1900 <= float(t) <= 2099)]
    if not nums:
        return None
    # 检测区间 A - B (取中点)
    for i in range(len(nums) - 1, 0, -1):
        ci, cv = nums[i]; pi, pv = nums[i - 1]
        if pi == ci - 2 and toks[pi + 1] in ('-', '–', '—'):
            return (float(pv) + float(cv)) / 2
    return float(nums[-1][1])

def parse_medium(thr_list):
    """返回 (value, count) —— 取所有数值的最小值(最灵敏报告阈值)。"""
    vals = []
    for s in thr_list:
        v = extract_value(s)
        if v and v > 0:
            vals.append(v)
    if not vals:
        return None, 0
    return min(vals), len(vals)

# ---- 聚合：以 CAS 为主键 ----
raw = json.load(open(SRC, encoding='utf-8'))
agg = {}
order = 0
for r in raw:
    cas = r.get('cas') or ('EN:' + base_name(r.get('english_name')))
    g = agg.setdefault(cas, {
        'en': r.get('english_name'), 'cn': r.get('chinese_name'),
        'med': {}, 'flavor_cat': set(), 'flavor_desc_cn': [], 'flavor_desc': [], 'order': order
    })
    order += 1
    med = r.get('medium') or '其他介质'
    g['med'].setdefault(med, []).extend(r.get('threshold_data') or [])
    for c in (r.get('flavor_categories') or []):
        g['flavor_cat'].add(c)
    for d in (r.get('flavor_desc_cn') or []):
        if d and d not in g['flavor_desc_cn']:
            g['flavor_desc_cn'].append(d)
    for d in (r.get('flavor_desc') or []):
        if d and d not in g['flavor_desc']:
            g['flavor_desc'].append(d)
    if not g['cn'] and r.get('chinese_name'):
        g['cn'] = r['chinese_name']

records = []
skipped_no_name = 0
for cas, g in agg.items():
    en = (g['en'] or '').strip()
    if not en:
        skipped_no_name += 1
        continue
    cn = (g['cn'] or '').strip()
    # 仅当含中文字符才视为有效中文名；否则保留英文全名（避免英文碎片/乱码当中文名）
    if not re.search(r'[一-鿿]', cn):
        cn = ''
    syn = synonyms(en)
    cat = fc.auto_classify(en)
    # 阈值：水 > 空气 > 其他
    w = g['med'].get('水'); a = g['med'].get('空气'); o = g['med'].get('其他介质')
    thr, med, thr_count = '—', '—', 0
    if w:
        v, c = parse_medium(w); thr, med, thr_count = (f'≈{v:g}' if v else '—'), '水', c
    elif a:
        v, c = parse_medium(a)
        if v:
            med = '空气(mg/m³)'; thr_count = c
            # 空气阈值量纲不同, 不计入 OAV, 仅作备注
    elif o:
        v, c = parse_medium(o)
        if v:
            med = '其他介质'; thr_count = c
    # 风味描述：优先中文 flavor_desc_cn，缺则回退英文 flavor_desc（中英对照）
    fcat = '、'.join(sorted(g['flavor_cat']))
    fdesc_cn = [d for d in (g['flavor_desc_cn'] or []) if d and not re.fullmatch(r'[A-Za-z0-9 /]+', d)]
    fdesc_en = [d for d in (g['flavor_desc'] or []) if d]
    fdesc = '、'.join(fdesc_cn[:4]) if fdesc_cn else ''
    fdesc_e = '、'.join(fdesc_en[:4]) if fdesc_en else ''
    note_parts = []
    if fcat:
        note_parts.append('风味类别: ' + fcat)
    if fdesc:
        note_parts.append('描述: ' + fdesc)
    elif fdesc_e:
        note_parts.append('描述: ' + fdesc_e)
    if med.startswith('空气') or med == '其他介质':
        # 把空气/其他阈值值补进备注(避免量纲混用, 不进 OAV)
        v, c = parse_medium(g['med'].get(med.replace('(mg/m³)', '').strip() if med.startswith('空气') else '其他介质') or [])
        if v:
            note_parts.append(f'（{med}阈值≈{v:g}, 量纲不同未计入OAV）')
    note = '；'.join(note_parts) if note_parts else '—'
    rec = dict(en=en, cn=cn or '(无中文名)', cat=cat, thr=thr, med=med,
               note=note, src='FTDB v1.5.0', syn=syn, cas=cas, order=g['order'])
    records.append(rec)

records.sort(key=lambda x: x['order'])
json.dump(records, open(OUT, 'w', encoding='utf-8'), ensure_ascii=False, indent=1)
print(f"生成 {OUT}")
print(f"  FTDB 唯一化合物(去重后): {len(agg)}")
print(f"  有效记录(有英文名): {len(records)}  跳过无名: {skipped_no_name}")
# 统计
from collections import Counter
c = Counter(r['cat'] for r in records)
print('  类别分布:', dict(c))
print('  含水相阈值(可OAV):', sum(1 for r in records if r['thr'] != '—'))
print('  无阈值(仅空气/其他或缺失):', sum(1 for r in records if r['thr'] == '—'))
