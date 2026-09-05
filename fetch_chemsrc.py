#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""按 CAS 号从 chemsrc.com 抓取权威中文名（酯类/内酯类修正用）。"""
import json, re, sys, time, os, html as _html
import requests
from concurrent.futures import ThreadPoolExecutor, as_completed

H = {
    'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 '
                  '(KHTML, like Gecko) Chrome/120.0 Safari/537.36',
    'Referer': 'https://www.chemsrc.com/',
}
API = 'https://search.chemsrc.com/api/search?keyword=%s&source=chemsrc&page=1'
OUT = '/workspace/flavor_tool/chemsrc_cache.json'


def _strip(s):
    s = re.sub(r'<[^>]+>', ' ', s or '')
    s = s.replace('&nbsp;', ' ').replace('&amp;', '&')
    return re.sub(r'\s+', ' ', s).strip()


def fetch(cas):
    cas = (cas or '').strip()
    if not cas:
        return None
    try:
        r = requests.get(API % cas, headers=H, timeout=30)
        if r.status_code != 200:
            return {'cas': cas, 'ok': False, 'err': 'http%d' % r.status_code}
        t = r.text
        rec = {'cas': cas, 'ok': True, 'url': r.url}
        m = re.search(r'<title>(.*?)</title>', t, re.S)
        rec['title_cn'] = _strip(m.group(1)).split('_')[0].strip() if m else ''
        # JSON-LD
        cn = en = formula = ''
        for pm in re.finditer(
                r'"@type":"PropertyValue","name":"(中文名称|英文名称|分子式)","value":"([^"]*)"', t):
            k, v = pm.group(1), pm.group(2)
            if k == '中文名称':
                cn = cn or v
            elif k == '英文名称':
                en = en or v
            else:
                formula = formula or v
        rec['cn'] = _html.unescape(cn or rec['title_cn'])
        rec['en'] = _html.unescape(en)
        rec['formula'] = formula
        # 中文别名
        m = re.search(r'中文别名', t)
        if m:
            seg = _html.unescape(_strip(t[m.end():m.end() + 600]))
            rec['cn_alias'] = seg[:200]
        # 英文别名
        m = re.search(r'英文名?\s*</td>\s*<td[^>]*>(.*?)</td>', t, re.S)
        return rec
    except Exception as e:
        return {'cas': cas, 'ok': False, 'err': str(e)[:80]}


def main():
    items = json.load(open('/tmp/ester_lactone.json'))
    cas_list = sorted({(c.get('cas') or '').strip() for c in items if (c.get('cas') or '').strip()})
    print('待抓取 CAS: %d' % len(cas_list))
    cache = {}
    if os.path.exists(OUT):
        cache = json.load(open(OUT))
    todo = [c for c in cas_list if c not in cache or not cache[c].get('ok')]
    print('缓存命中 %d，待抓取 %d' % (len(cas_list) - len(todo), len(todo)))
    done = 0
    with ThreadPoolExecutor(max_workers=6) as ex:
        futs = {ex.submit(fetch, c): c for c in todo}
        for f in as_completed(futs):
            c = futs[f]
            try:
                rec = f.result()
            except Exception as e:
                rec = {'cas': c, 'ok': False, 'err': str(e)[:80]}
            cache[c] = rec
            done += 1
            if done % 25 == 0:
                print('  进度 %d/%d' % (done, len(todo)), flush=True)
                json.dump(cache, open(OUT, 'w'), ensure_ascii=False, indent=0)
    json.dump(cache, open(OUT, 'w'), ensure_ascii=False, indent=0)
    ok = sum(1 for v in cache.values() if v.get('ok'))
    print('完成：成功 %d / 失败 %d -> %s' % (ok, len(cache) - ok, OUT))


if __name__ == '__main__':
    main()
