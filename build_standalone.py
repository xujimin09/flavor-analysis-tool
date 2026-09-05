# -*- coding: utf-8 -*-
"""把 flavor_core 的数据库与最新前端逻辑打包成一个「真正离线可用」的单文件 HTML。

特性：
  - 内置 __N__ 种化合物数据库（内联 JSON，无需服务器）
  - 单物质检索 / 批量检索 / GC-MS 报告解析 / 多报告对比 / 浓度计算 / 数据库浏览
  - OAV / ROAV 计算、气味活性判定（关键致香 / 潜在贡献）
  - 类别覆盖卡片（可点击展开） + 多报告对比按类别排序、同类数值降序、深色数值
  - Excel(.xlsx) 解析库 SheetJS 已内联 → 离线可读 GC-MS 报告
  - PDF(.pdf) 解析库 pdf.js 已内联 → 离线可读 PDF 报告
  - Word(.docx) 解析走 CDN mammoth，离线时降级提示「另存 CSV」
  - 无需联网（除 Word 解析外），无需服务器，双击即用
"""
import json, os, base64, re

HERE = os.path.dirname(os.path.abspath(__file__))
spec = importlib_spec = None

import importlib.util
spec = importlib.util.spec_from_file_location("fc", os.path.join(HERE, "flavor_core.py"))
fc = importlib.util.module_from_spec(spec)
spec.loader.exec_module(fc)

# 取出当前数据库（含 custom 合并）
data = json.dumps(fc.COMPOUNDS, ensure_ascii=False)
# 类别配色（与 app.py 保持一致；flavor_core 未导出该变量）
CAT_COLOR = {
    '醛类': '#E8590C', '酮类': '#B8860B', '酯类': '#2F9E44', '酸类': '#1971C2',
    '醇类': '#5C5C5C', '萜烯类': '#6741D9', '含硫化合物': '#E03131', '吡嗪类': '#0C8599',
    '内酯类': '#C2255C', '其他': '#868E96',
}

# ---- 内联第三方解析库 ----
def read_vendor(name):
    p = os.path.join(HERE, "vendor", name)
    if not os.path.exists(p):
        return None
    with open(p, "r", encoding="utf-8", errors="replace") as f:
        return f.read()

XLSX_JS = read_vendor("xlsx.full.min.js") or ""
PDF_JS  = read_vendor("pdf.min.js") or ""
PDF_WORKER_JS = read_vendor("pdf.worker.min.js") or ""
MAMMOTH_JS = read_vendor("mammoth.browser.min.js") or ""

# pdf.js CMaps 内联（离线抽取中文 / 各类 CID 字体 PDF 文本所必需）
CMAP_JS = ""
_cmap_dir = os.path.join(HERE, "vendor", "cmaps")
if os.path.isdir(_cmap_dir):
    import base64 as _b64
    _cmap_obj = {}
    for _fn in sorted(os.listdir(_cmap_dir)):
        if _fn.endswith(".bcmap"):
            with open(os.path.join(_cmap_dir, _fn), "rb") as _f:
                _cmap_obj[_fn] = _b64.b64encode(_f.read()).decode("ascii")
    if _cmap_obj:
        CMAP_JS = ("/* pdf.js CMaps 内联（离线中文/CID 字体 PDF 文本抽取必需） */\n"
                   "var CMAP_DATA=" + json.dumps(_cmap_obj, ensure_ascii=False) + ";\n")

USE_XLSX = "/* SheetJS 内联（离线可用） */\n" + XLSX_JS if XLSX_JS else ""
USE_PDF  = "/* pdf.js 内联（离线可用） */\n" + PDF_JS if PDF_JS else ""
# pdf.js worker：直接在主线程内联为普通脚本。
# 关键：pdf.js 3.x 的 PDFWorker._initialize() 判断
#   if(!isWorkerDisabled && !PDFWorker._mainThreadWorkerMessageHandler) { 创建 Worker }
#   else this._setupFakeWorker();
# 而 _mainThreadWorkerMessageHandler 取自 globalThis.pdfjsWorker?.WorkerMessageHandler。
# 因此只要把 worker 脚本以普通 <script> 内联到主线程，window.pdfjsWorker 就存在，
# pdf.js 会自动走「fake worker（主线程）」模式，完全不创建 Worker。
# —— 这解决了 file:// 协议下 new Worker(blob:null/...) 被浏览器拒绝、
#    以及 isSameOrigin() 返回 false 后 createCDNWrapper() 用 importScripts 二次加载失败的问题。
USE_PDF_WORKER = "/* pdf.js worker 内联（主线程模式，file:// 下同样可用） */\n" + PDF_WORKER_JS if PDF_WORKER_JS else ""
USE_MAMMOTH = "/* mammoth 内联（离线可用） */\n" + MAMMOTH_JS if MAMMOTH_JS else ""

# ---- 拉取 app.py 中新增的「处理组对比」模块（CSS + JS）----
# CSS：从「/* ===== 处理组对比（单 PDF 多处理组自动识别） ===== */」到「/* 类别徽章」前
# JS ：从「/* ===================== 处理组对比（单 PDF 多处理组自动识别） ===================== */」
#      到「/* ===================== 多报告批量对比 ===================== */」前（仅多组对比函数，避免与离线端既有 doMulti/renderMulti 重复）
GC_MULTI_CSS = ""
GC_MULTI_JS = ""
GC_MULTI_HTML = ""
BASE_STYLE = ""          # 在线版 app.py 的基础 <style>（单一来源，避免离线版分叉）
BASE_SKELETON = ""       # 在线版 app.py 的 tabs + 面板骨架
try:
    src = open(os.path.join(HERE, "app.py"), encoding="utf-8").read()
    m = re.search(r'PAGE\s*=\s*r"""(.*?)"""', src, re.S)
    page = m.group(1) if m else ""
    if page:
        m = re.search(r'/\* ===== 处理组对比.*?\*/\s*(.*?)(?=\n\s*/\* 类别徽章)', page, re.S)
        GC_MULTI_CSS = m.group(1).strip() if m else ""
        # JS 段：仅匹配脚本区那个「=」很多的注释（CSS 段只有 5 个 =，避免误命中）
        m = re.search(r'/\* ={8,} 处理组对比.*?\*/\s*(.*?)(?=\n/\* ={8,} 多报告批量对比)', page, re.S)
        GC_MULTI_JS = m.group(1).strip() if m else ""
        # 基础 UI 单一来源：直接复用在线版 app.py 的 <style> 与 tabs+面板骨架，避免离线版再次分叉
        m = re.search(r'<style>(.*?)</style>', page, re.S)
        BASE_STYLE = m.group(1).strip() if m else ""
        m = re.search(r'<div class="tabs">.*?(?=<script translate="no">)', page, re.S)
        BASE_SKELETON = m.group(0).strip() if m else ""
        if BASE_SKELETON:
            # 离线报告解析依赖 xlsxWarn 节点（app.py 骨架无此节点），补回
            BASE_SKELETON = BASE_SKELETON.replace(
                '<input type="file" id="reportfile"',
                '<div class="warn" id="xlsxWarn" style="display:none">⚠ 解析库未加载（可能离线），无法读取该格式；请把报告另存为 CSV（第一列=名称，第二列=响应比）后再上传。</div>\n    <input type="file" id="reportfile"')
    if not GC_MULTI_CSS or not GC_MULTI_JS:
        print("[warn] 处理组对比模块提取不完整：CSS=%d JS=%d" % (len(GC_MULTI_CSS), len(GC_MULTI_JS)))
except Exception as _e:
    print("[warn] 拉取 app.py 处理组对比模块失败:", _e)

# ---- 拉取 app.py 中的「Nature 风格可视化（热力图 / PCA）+ 物质管理」脚本 ----
NATURE_VIZ_JS = ""
try:
    src = open(os.path.join(HERE, "app.py"), encoding="utf-8").read()
    m = re.search(r'PAGE\s*=\s*r"""(.*?)"""', src, re.S)
    page = m.group(1) if m else ""
    if page:
        m = re.search(r'/\* ===== NATURE_VIZ_START ===== \*/\s*(.*?)\n\s*/\* ===== NATURE_VIZ_END ===== \*/', page, re.S)
        NATURE_VIZ_JS = m.group(1).strip() if m else ""
    if not NATURE_VIZ_JS:
        print("[warn] Nature 可视化脚本提取为空")
except Exception as _e:
    print("[warn] 拉取 Nature 可视化脚本失败:", _e)

# 旧版走 CDN 的标签已弃用（改为全部内联，零外部依赖）
MAMMOTH_TAG = ""

COMPOUNDS_N = len(fc.COMPOUNDS)

# ============ HTML 模板 ============
HTML = r"""<!DOCTYPE html>
<html lang="zh-CN">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<meta name="google" content="notranslate">
<meta http-equiv="Content-Language" content="zh-CN, en">
<title>风味物质检索 · GC-MS 解析 · 对比（离线版）</title>
<style>
  :root{
    --bg:#eef3f9; --card:#fff; --glass:rgba(255,255,255,.78); --line:#e3e8ef; --line-soft:#eef1f6;
    --txt:#1f2937; --mut:#6b7280; --mut-2:#9aa3b2; --teal:#7dd3e5; --coral:#f47c6b; --good:#5dbe6c; --warn:#f59e0b;
    --teal-bg:#eaf7fb; --teal-d:#1f7a93; --coral-bg:#fde9e5; --good-bg:#e7f6ea; --lav:#cdb4f0; --rose:#f5a3a3;
    --radius:16px; --radius-sm:12px; --shadow:0 8px 30px rgba(31,42,61,.08); --shadow-sm:0 2px 12px rgba(31,42,61,.06);
    --glass-line:rgba(255,255,255,.7);
    --c-醛类:#e06b7d; --cb-醛类:#fde6ec; --c-酮类:#d49a4e; --cb-酮类:#fdf1de; --c-酯类:#2f9e44; --cb-酯类:#e2efda;
    --c-酸类:#1971c2; --cb-酸类:#e5f3f8; --c-醇类:#5c5c5c; --cb-醇类:#eef1f6; --c-萜烯类与含氧萜类:#6741d9; --cb-萜烯类与含氧萜类:#ece4f7;
    --c-含硫化合物:#e03131; --cb-含硫化合物:#fbe6e6; --c-吡嗪类:#0c8599; --cb-吡嗪类:#e1f4f7; --c-内酯类:#c2255c; --cb-内酯类:#fce3ee;
    --c-含氮杂环其他含氮化合物:#7950f2; --cb-含氮杂环其他含氮化合物:#f3f0ff;
    --c-呋喃呋喃酮类:#e67700; --cb-呋喃呋喃酮类:#fff0d9;
    --c-噻唑类噻唑啉类:#40c057; --cb-噻唑类噻唑啉类:#e6fcf5;
    --c-其他:#868e96; --cb-其他:#eef1f6;
  }
  *{box-sizing:border-box}
  body{margin:0;background:radial-gradient(1200px 600px at 80% -10%,#e8f0ff 0,transparent 60%),radial-gradient(900px 500px at -10% 10%,#fdece9 0,transparent 55%),var(--bg);color:var(--txt);font-family:Arial,"Helvetica Neue","PingFang SC","Microsoft YaHei",sans-serif;font-size:14px;line-height:1.5;padding:24px 16px 56px;-webkit-font-smoothing:antialiased}
  .page{max-width:1180px;margin:0 auto}
  h1{font-size:26px;font-weight:800;margin:0 0 4px;display:flex;align-items:center;gap:12px}
  h1 .logo{width:34px;height:34px;border-radius:12px;background:linear-gradient(135deg,var(--teal),var(--coral));color:#fff;display:inline-flex;align-items:center;justify-content:center;font-size:17px;font-weight:800}
  .sub{color:var(--mut);font-size:13.5px;margin:0 0 18px}
  .sub b{color:var(--teal-d);font-weight:700}
  .tabs{display:flex;gap:6px;flex-wrap:wrap;background:var(--glass);backdrop-filter:blur(12px);-webkit-backdrop-filter:blur(12px);padding:6px;border-radius:16px;box-shadow:var(--shadow);margin-bottom:20px;border:1px solid var(--glass-line)}
  .tab{padding:10px 18px;border-radius:12px;cursor:pointer;background:transparent;color:var(--mut);font-size:14px;font-weight:600}
  .tab:hover{color:var(--teal-d);background:var(--teal-bg)}
  .tab.active{background:var(--teal);color:#1f2a3d;box-shadow:0 4px 12px rgba(125,211,229,.4);font-weight:700}
  .card{background:var(--glass);backdrop-filter:blur(16px);-webkit-backdrop-filter:blur(16px);border-radius:var(--radius);padding:22px 24px;margin-bottom:18px;box-shadow:var(--shadow);border:1px solid var(--glass-line)}
  .card h2{font-size:20px;font-weight:700;margin:0 0 12px;display:flex;align-items:center;gap:10px}
  .card h2 .ic{width:30px;height:30px;border-radius:10px;background:var(--teal-bg);color:var(--teal-d);display:inline-flex;align-items:center;justify-content:center;font-size:15px}
  label{font-weight:600;display:block;margin:10px 0 4px;color:var(--txt)}
  input[type=text],textarea,select{width:100%;padding:10px 12px;border:1px solid var(--line);border-radius:10px;font-size:14px;font-family:inherit;background:#fff;color:var(--txt)}
  textarea{min-height:130px;resize:vertical}
  button{background:var(--teal);color:#1f2a3d;border:0;border-radius:var(--radius-sm);padding:11px 22px;font-size:14px;font-weight:700;cursor:pointer;transition:all .2s;box-shadow:0 4px 12px rgba(125,211,229,.35);letter-spacing:.02em}
  button:hover{background:var(--coral);color:#fff;transform:translateY(-1px);box-shadow:0 6px 18px rgba(244,124,107,.35)}
  button.ghost{background:#fff;border:1.5px solid var(--line);color:var(--teal-d);box-shadow:none;margin-left:6px}
  button.ghost:hover{background:var(--teal-bg);border-color:var(--teal);color:var(--teal-d);transform:none}
  .row{display:flex;gap:10px;align-items:flex-end;flex-wrap:wrap}
  .row>*{flex:1;min-width:120px}
  .hint{font-size:12.5px;color:var(--mut);margin:6px 0}
  .warn{background:#fff7e6;border:1px solid #ffe0a3;color:#8a5a00;padding:8px 12px;border-radius:8px;font-size:12.5px;margin-top:10px}
  .badge{display:inline-block;padding:3px 10px;border-radius:8px;font-size:11.5px;font-weight:700;white-space:nowrap;letter-spacing:.02em;background:var(--cb-其他);color:var(--c-其他)}
  .b-醛类{background:var(--cb-醛类);color:var(--c-醛类)} .b-酮类{background:var(--cb-酮类);color:var(--c-酮类)} .b-酯类{background:var(--cb-酯类);color:var(--c-酯类)}
  .b-酸类{background:var(--cb-酸类);color:var(--c-酸类)} .b-醇类{background:var(--cb-醇类);color:var(--c-醇类)} .b-萜烯类与含氧萜类{background:var(--cb-萜烯类与含氧萜类);color:var(--c-萜烯类与含氧萜类)}
  .b-含硫化合物{background:var(--cb-含硫化合物);color:var(--c-含硫化合物)} .b-吡嗪类{background:var(--cb-吡嗪类);color:var(--c-吡嗪类)} .b-内酯类{background:var(--cb-内酯类);color:var(--c-内酯类)}
  .b-含氮杂环其他含氮化合物{background:var(--cb-含氮杂环其他含氮化合物);color:var(--c-含氮杂环其他含氮化合物)} .b-呋喃呋喃酮类{background:var(--cb-呋喃呋喃酮类);color:var(--c-呋喃呋喃酮类)} .b-噻唑类噻唑啉类{background:var(--cb-噻唑类噻唑啉类);color:var(--c-噻唑类噻唑啉类)} .b-其他{background:var(--cb-其他);color:var(--c-其他)}
  /* 表格容器：.scroll 与 .wrap 统一为圆角 + 细边框 + 轻投影 + 滚动条美化 */
  .scroll,.wrap{max-height:66vh;overflow:auto;border:1px solid var(--line);border-radius:var(--radius);background:#fff;box-shadow:var(--shadow-sm)}
  .scroll::-webkit-scrollbar,.wrap::-webkit-scrollbar{height:10px;width:10px}
  .scroll::-webkit-scrollbar-thumb,.wrap::-webkit-scrollbar-thumb{background:#d8dee6;border-radius:8px;border:2px solid #fff}
  .scroll::-webkit-scrollbar-thumb:hover,.wrap::-webkit-scrollbar-thumb:hover{background:#c2ccd6}
  /* 基础表格：行 / 列 / 文字分布适中，清晰易读（与在线版统一「记忆格式」） */
  table{border-collapse:separate;border-spacing:0;width:auto;min-width:100%;font-size:12.5px;line-height:1.55;margin:0;color:var(--txt)}
  th,td{padding:10px 14px;text-align:left;vertical-align:top;border-bottom:1px solid var(--line-soft);word-break:break-word;overflow-wrap:anywhere}
  th{background:var(--card-soft);color:var(--mut);font-weight:700;font-size:11px;position:sticky;top:0;z-index:2;border-bottom:1.5px solid var(--line);white-space:nowrap;letter-spacing:.03em}
  .scroll thead th:first-child,.wrap thead th:first-child{border-top-left-radius:var(--radius)}
  .scroll thead th:last-child,.wrap thead th:last-child{border-top-right-radius:var(--radius)}
  tbody tr:nth-child(even) td{background:#fcfdfe}
  tbody tr:hover td{background:#eef5fb}
  tr:last-child td{border-bottom:0}
  td.num,th.num{text-align:right;font-variant-numeric:tabular-nums;white-space:nowrap}
  td.cn{font-weight:800;color:var(--txt)}
  /* 描述 / 来源 列：行高舒展、字号适中、长文本换行不挤压 */
  td.desc,td.gc-desc{text-align:left;white-space:normal;color:var(--txt);font-size:12px;line-height:1.55;max-width:240px}
  td.src,td.gc-src{text-align:left;white-space:normal;color:var(--mut);font-size:12px;line-height:1.6;max-width:180px}
  td.src b,td.gc-src b{color:var(--txt);font-weight:700}
  /* 报告解析主表：固定列宽，压缩中英文/物质列，保证后续数据列·来源·风味描述清晰（与在线版一致） */
  #report-table{table-layout:fixed;width:100%}
  #report-table th,#report-table td{font-variant-numeric:tabular-nums}
  #report-table th:nth-child(1),#report-table td:nth-child(1){width:34px;text-align:center;white-space:nowrap;color:var(--mut-2);font-variant-numeric:tabular-nums}
  #report-table th:nth-child(2),#report-table td:nth-child(2){width:13%;min-width:104px;text-align:left;white-space:normal;word-break:break-word;overflow-wrap:anywhere}
  #report-table th:nth-child(3),#report-table td:nth-child(3){width:8%;min-width:68px;text-align:left;white-space:normal;word-break:break-word}
  #report-table th:nth-child(4),#report-table td:nth-child(4){text-align:left;white-space:nowrap;min-width:80px}
  #report-table th:nth-child(5),#report-table td:nth-child(5){width:8%;min-width:58px;text-align:left;white-space:normal}
  #report-table th:nth-child(6),#report-table th:nth-child(7),#report-table th:nth-child(10),#report-table th:nth-child(11),#report-table th:nth-child(12),
  #report-table td:nth-child(6),#report-table td:nth-child(7),#report-table td:nth-child(10),#report-table td:nth-child(11),#report-table td:nth-child(12){text-align:right;white-space:nowrap}
  #report-table th:nth-child(8),#report-table th:nth-child(9),
  #report-table td:nth-child(8),#report-table td:nth-child(9){text-align:right;white-space:normal}
  #report-table th:nth-child(12),#report-table td:nth-child(12){text-align:center;white-space:nowrap;min-width:88px}
  #report-table th:nth-child(13),#report-table td:nth-child(13){width:19%;min-width:196px;text-align:left;white-space:normal}
  #report-table th:nth-child(14),#report-table td:nth-child(14){width:13%;min-width:150px;text-align:left;white-space:normal}
  /* 相对于对照列：上调(≥2×)红、下调(≤0.5×)蓝（与在线版一致） */
  .fc-up{color:#e0533a;font-weight:700}
  .fc-down{color:#2a7de1;font-weight:700}
  /* ===== Nature 风格可视化（热力图 / PCA）+ 物质管理 ===== */
  .viz-wrap{margin-top:18px;border-top:1px dashed var(--line);padding-top:14px}
  .viz-head{display:flex;align-items:center;gap:10px;flex-wrap:wrap;margin:4px 0 10px}
  .viz-head h3{margin:0;font-size:15px;display:flex;align-items:center;gap:8px}
  .viz-head h3 .ic{width:24px;height:24px;border-radius:8px;background:var(--teal-bg);color:var(--teal);display:inline-flex;align-items:center;justify-content:center;font-size:13px}
  .viz-toolbar{display:flex;flex-wrap:wrap;gap:8px;align-items:center;margin-bottom:10px}
  .viz-toolbar label{font-size:12.5px;color:var(--mut);font-weight:600}
  .viz-toolbar select{padding:6px 9px;border-radius:8px;border:1px solid var(--line);background:var(--card);color:var(--txt);font-size:13px}
  .viz-chips{display:flex;flex-wrap:wrap;gap:6px;margin:2px 0 8px}
  .viz-chip{display:inline-flex;align-items:center;gap:5px;padding:3px 10px;border-radius:18px;border:1px solid var(--line);background:var(--card);font-size:12px;font-weight:600;cursor:pointer;user-select:none;transition:all .15s}
  .viz-chip .dot{width:9px;height:9px;border-radius:50%}
  .viz-chip.off{opacity:.4;background:#f1f3f6}
  .viz-chip:hover{border-color:var(--teal)}
  .viz-chip em.cn{font-style:normal;font-size:10.5px;color:var(--mut);background:var(--card-soft);border-radius:8px;padding:0 6px;margin-left:2px;font-weight:600}
  .viz-chip.off em.cn{color:var(--mut-2)}
  .viz-search{flex:1;min-width:160px;padding:7px 10px;border-radius:9px;border:1px solid var(--line);background:#fff;color:var(--txt);font-size:13px}
  .viz-main{display:grid;grid-template-columns:minmax(320px,1.4fr) minmax(300px,1fr);gap:14px}
  @media(max-width:900px){ .viz-main{grid-template-columns:1fr} }
  .viz-fig{background:#fff;border:1px solid var(--line);border-radius:var(--radius);padding:12px 12px 8px;box-shadow:var(--shadow-sm)}
  .viz-fig h4{margin:0 0 6px;font-size:13px;color:var(--txt);font-weight:700;display:flex;justify-content:space-between;align-items:baseline}
  .viz-fig h4 small{font-weight:500;color:var(--mut);font-size:11px}
  .viz-fig svg{width:100%;height:auto;display:block}
  .viz-mgr{background:#fff;border:1px solid var(--line);border-radius:var(--radius);padding:10px 12px;box-shadow:var(--shadow-sm);display:flex;flex-direction:column;max-height:520px}
  .viz-mgr .mgr-top{display:flex;gap:6px;flex-wrap:wrap;margin-bottom:8px}
  .viz-mgr .mgr-list{overflow:auto;flex:1;border-top:1px solid var(--line-soft)}
  .viz-mrow{display:flex;align-items:center;gap:8px;padding:5px 6px;border-bottom:1px solid var(--line-soft);font-size:12.5px}
  .viz-mrow:hover{background:#f6fafc}
  .viz-mrow input{width:15px;height:15px;accent-color:var(--teal);flex:0 0 auto}
  .viz-mrow .mn{flex:1;min-width:0;overflow:hidden;text-overflow:ellipsis;white-space:nowrap}
  .viz-mrow .mn b{color:var(--txt)}
  .viz-mrow .mc{font-size:10.5px;color:var(--mut);flex:0 0 auto;padding:1px 7px;border-radius:7px;background:var(--card-soft)}
  .viz-mrow .mv{font-size:11px;color:var(--mut-2);font-variant-numeric:tabular-nums;flex:0 0 auto;min-width:54px;text-align:right}
  .viz-mrow.del{opacity:.45}
  .viz-mrow.del .mn{text-decoration:line-through}
  .viz-hint{font-size:11.5px;color:var(--mut-2);margin-top:6px;line-height:1.5}
  .viz-legend{display:flex;flex-wrap:wrap;gap:10px;margin-top:6px;font-size:11px;color:var(--mut)}
  .viz-legend span{display:inline-flex;align-items:center;gap:4px}
  .viz-legend i{width:11px;height:11px;border-radius:3px;display:inline-block}
  /* 类别卡片网格 */
  .cat-bars{background:var(--card);border-radius:var(--radius);padding:18px 20px;box-shadow:var(--shadow-sm);border:1px solid var(--line);margin-bottom:16px}
  .cat-bars h3{font-size:15px;font-weight:700;margin:0 0 12px;display:flex;align-items:center;gap:8px}
  .cat-bars h3 .ic{width:24px;height:24px;border-radius:8px;background:var(--teal-bg);color:var(--teal);display:inline-flex;align-items:center;justify-content:center;font-size:13px}
  .cat-grid{display:grid;grid-template-columns:repeat(auto-fill,minmax(118px,1fr));gap:10px}
  .cat-chip{cursor:pointer;background:var(--teal-bg);border:1px solid var(--line);border-radius:12px;padding:12px 13px;display:flex;flex-direction:column;gap:9px;transition:transform .15s ease,box-shadow .15s ease,background .15s ease;min-height:78px}
  .cat-chip:hover{transform:translateY(-2px);box-shadow:0 6px 16px rgba(125,211,229,.28);background:#e6f6fa}
  .cat-chip .cc-head{display:flex;align-items:center;gap:7px}
  .cat-chip .cc-dot{width:11px;height:11px;border-radius:50%;flex:0 0 auto;box-shadow:0 1px 3px rgba(0,0,0,.18)}
  .cat-chip .cc-dot.b-醛类{background:var(--c-醛类)} .cat-chip .cc-dot.b-酮类{background:var(--c-酮类)} .cat-chip .cc-dot.b-酯类{background:var(--c-酯类)} .cat-chip .cc-dot.b-酸类{background:var(--c-酸类)} .cat-chip .cc-dot.b-醇类{background:var(--c-醇类)} .cat-chip .cc-dot.b-萜烯类与含氧萜类{background:var(--c-萜烯类与含氧萜类)} .cat-chip .cc-dot.b-含硫化合物{background:var(--c-含硫化合物)} .cat-chip .cc-dot.b-吡嗪类{background:var(--c-吡嗪类)} .cat-chip .cc-dot.b-内酯类{background:var(--c-内酯类)} .cat-chip .cc-dot.b-含氮杂环其他含氮化合物{background:var(--c-含氮杂环其他含氮化合物)} .cat-chip .cc-dot.b-呋喃呋喃酮类{background:var(--c-呋喃呋喃酮类)} .cat-chip .cc-dot.b-噻唑类噻唑啉类{background:var(--c-噻唑类噻唑啉类)} .cat-chip .cc-dot.b-其他{background:var(--c-其他)}
  .cat-chip .cc-name{font-size:13.5px;font-weight:700;line-height:1.25}
  .cat-chip .cc-num{margin-top:auto;font-size:24px;font-weight:800;color:var(--teal-d);letter-spacing:-.03em;font-variant-numeric:tabular-nums;line-height:1;display:flex;align-items:baseline;gap:3px}
  .cat-chip .cc-num small{font-size:12px;font-weight:700;color:var(--mut)}
  .cat-row.is-active{background:#e6f6fa;box-shadow:0 6px 16px rgba(125,211,229,.34);border-color:var(--teal)}
  .cat-detail{margin-top:14px} .cat-detail:empty{margin-top:0} .cat-detail .card{margin:0}
  .oavflag{font-weight:700;font-size:12.5px;white-space:nowrap}
  .oavflag.f-key{color:var(--good);background:var(--good-bg);padding:2px 8px;border-radius:8px}
  .oavflag.f-pot{color:var(--coral);background:var(--coral-bg);padding:2px 8px;border-radius:8px}
  .oavflag.f-na{color:var(--mut-2)}
  .num-dark{color:var(--txt);font-weight:700;font-variant-numeric:tabular-nums;font-size:13px;line-height:1.5}
  .multi-toolbar{display:flex;align-items:center;gap:14px;flex-wrap:wrap;margin:4px 0 14px}
  .multi-toolbar select{padding:7px 10px;border-radius:8px;border:1px solid var(--line);background:var(--card);color:var(--txt);font-size:13px}
  .field{display:flex;flex-direction:column}
  @media (max-width:760px){ table th,table td{font-size:11.5px} th,td{padding:5px 6px} .cat-grid{grid-template-columns:repeat(auto-fill,minmax(96px,1fr))} }
/* ===== 处理组对比（单 PDF 多处理组自动识别） ===== */
__GC_MULTI_CSS__
</style>
</head>
<body class="notranslate" translate="no">
<div class="page">
  <h1><span class="logo">香</span>风味物质检索 · GC-MS 解析 · 对比（离线版）</h1>
  <p class="sub">内置 <b>__N__</b> 种化合物数据库｜英文/中文/模糊匹配｜OAV/ROAV 气味活性｜内标法浓度｜多报告对比｜Excel/PDF 离线可解析（双击即用，无需服务器）</p>
  <div class="tabs">
    <div class="tab active" data-t="single">① 单物质检索</div>
    <div class="tab" data-t="batch">② 批量检索</div>
    <div class="tab" data-t="report">③ GC-MS 报告解析</div>
    <div class="tab" data-t="multi">④ 多报告对比</div>
    <div class="tab" data-t="browse">⑤ 数据库浏览</div>
  </div>

  <!-- ① 单物质 -->
  <div class="panel card" id="p-single">
    <label>输入物质名（英文 / 中文 / 允许拼写偏差）</label>
    <input type="text" id="q" placeholder="例如：Butanoic acid、丁酸、Limolene、2-Heptanone">
    <div class="row">
      <div style="flex:0 0 auto"><button onclick="doSingle()">检索</button></div>
      <div style="flex:1"><span class="hint">快捷：</span>
        <button class="ghost" onclick="byQ('Butanoic acid')">丁酸</button>
        <button class="ghost" onclick="byQ('Limolene')">Limolene</button>
        <button class="ghost" onclick="byQ('Unknown ketone XYZ')">未知物</button></div>
    </div>
    <div id="single-out"></div>
  </div>

  <!-- ② 批量 -->
  <div class="panel card" id="p-batch" style="display:none">
    <label>批量物质名（每行一个，可「名称,响应比」成对输入）</label>
    <textarea id="batch" placeholder="Butanoic acid&#10;2-Heptanone&#10;2-Nonanone&#10;gamma-Nonalactone&#10;丁酸"></textarea>
    <div class="row">
      <div style="flex:0 0 auto"><button onclick="doBatch()">批量检索</button></div>
      <div style="flex:0 0 auto"><button class="ghost" onclick="document.getElementById('batchfile').click()">从文件导入</button><input type="file" id="batchfile" accept=".txt,.csv" style="display:none" onchange="loadBatchFile()"></div>
    </div>
    <div id="batch-out"></div>
  </div>

  <!-- ③ 报告 -->
  <div class="panel card" id="p-report" style="display:none">
    <label>上传 GC-MS 报告（.xlsx / .pdf；Word .docx 需联网解析库）</label>
    <input type="file" id="reportfile" accept=".xlsx,.xls,.pdf,.docx,.md,.markdown,.txt">
    <div class="warn" id="xlsxWarn" style="display:none">⚠ 解析库未加载（可能离线），无法读取该格式；请把报告另存为 CSV（第一列=名称，第二列=响应比）后再上传。</div>
    <div class="row" style="margin-top:12px">
      <span class="hint" style="width:100%">内标法浓度参数（直接使用响应比 rr 计算：C = 内标浓度 × 内标加量 × 响应比 ÷ 样品量，统一到 μg/L）：</span>
      <div class="field"><label class="hint">内标物质<input type="text" id="r-ist" value="2-辛醇" style="width:90px"></label></div>
      <div class="field"><label class="hint">内标浓度<input type="text" id="r-cis" value="10" style="width:90px"></label></div>
      <div class="field"><label class="hint">浓度单位<input type="text" id="r-cisu" value="mg/L" style="width:64px"></label></div>
      <div class="field"><label class="hint">内标加量<input type="text" id="r-vis" value="50" style="width:90px"></label></div>
      <div class="field"><label class="hint">体积单位<input type="text" id="r-visu" value="μL" style="width:50px"></label></div>
      <div class="field"><label class="hint">样品量<input type="text" id="r-ms" value="2" style="width:90px"></label></div>
      <div class="field"><label class="hint">取样量单位<input type="text" id="r-msu" value="g" style="width:46px"></label></div>
      <div class="field" style="flex:0 0 auto"><label>&nbsp;</label><button class="ghost" onclick="recalcReportConc()">重算浓度</button></div>
    </div>
    <div id="report-out"></div>
  </div>

  <!-- ④ 多报告对比 -->
  <div class="panel card" id="p-multi" style="display:none">
    <p class="hint">逐个上传 GC-MS 报告（.xlsx / .pdf / .docx / .txt），每个文件上传后填入<b>样品名称</b>；全部添加后点「开始对比分析」，系统按各报告自带浓度/OAV/ROAV 横向对比共有/特有物质与关键致香物。</p>
    <div class="row" style="align-items:flex-start">
      <div style="flex:0 0 auto"><span class="up-btn" style="display:inline-block;padding:11px 20px;background:var(--teal);color:#1f2a3d;border-radius:12px;font-weight:700;cursor:pointer;position:relative;overflow:hidden">选择文件<input type="file" id="multifile" accept=".xlsx,.xls,.pdf,.docx,.md,.markdown,.txt" onchange="onMultiFile(event)" style="position:absolute;inset:0;opacity:0;width:100%;height:100%;cursor:pointer;border:0;font-size:0"></span></div>
    </div>
    <div class="row" id="sample-list" style="margin-top:12px" style="flex-wrap:wrap"></div>
    <div class="row" style="margin-top:12px;align-items:center;flex-wrap:wrap;gap:10px">
      <span class="hint">统一内标参数：</span>
      内标物质<input type="text" id="m-ist" value="2-辛醇" style="width:64px" oninput="document.getElementById('m-ist').dataset.touched='1';">
      内标浓度<input type="text" id="m-cis" value="10" style="width:60px"> <input type="text" id="m-cisu" value="mg/L" style="width:56px">
      内标加量<input type="text" id="m-vis" value="50" style="width:60px"> <input type="text" id="m-visu" value="μL" style="width:46px">
      样品量<input type="text" id="m-ms" value="2" style="width:60px"> <input type="text" id="m-msu" value="g" style="width:42px">
    </div>
    <div class="row" style="margin-top:12px;align-items:center">
      <div style="flex:0 0 auto"><button onclick="doMulti()">开始对比分析</button></div>
      <div style="flex:0 0 auto"><button class="ghost" onclick="clearMulti()">清空</button></div>
    </div>
    <div id="multi-out"></div>
  </div>

  <!-- ⑥ 浏览 -->
  <div class="panel card" id="p-browse" style="display:none">
    <label>按类别筛选 / 搜索</label>
    <div class="row">
      <div style="flex:0 0 auto;min-width:160px"><select id="bcat" onchange="browseDB()">
        <option value="">全部类别</option>
        <option>醛类</option><option>酮类</option><option>酯类</option><option>酸类</option><option>醇类</option>
        <option>萜烯类</option><option>含硫化合物</option><option>吡嗪类</option><option>内酯类</option><option>其他</option>
      </select></div>
      <div style="flex:1"><input type="text" id="bkey" placeholder="搜索：如 丁酸 / Limonene / 奶酪" oninput="browseDB()"></div>
      <div style="flex:0 0 auto"><span class="hint" id="bcount"></span></div>
    </div>
    <div id="browse-out"></div>
  </div>
</div>

<!-- ===== 内联解析库（离线） ===== -->
<script>
__USE_XLSX__
</script>
<script>
__USE_PDF_WORKER__
</script>
<script>
__USE_PDF__
</script>
<script>
__USE_CMAPS__
</script>
<script>
__USE_MAMMOTH__
</script>

<script translate="no">
"use strict";
// 防注入/翻译扩展损坏事件的兜底
(function(){
  try{
    window.addEventListener('error', function(e){ console.warn('[flavor-tool-offline] captured error', e.message); });
  }catch(_){ }
})();
const DB_RAW = __DATA__;
const CAT_COLOR = __CAT_COLOR__;
const CAT = CAT_COLOR;

// ---------- 工具 ----------
function normalize(s){ if(typeof s!=='string') return ''; s=s.toLowerCase().replace(/[’‘]/g,"'"); return s.replace(/[^a-z0-9一-鿿]/g,''); }
const TERP=['terpene','terpen','萜','limonene','pinene','myrcene','ocimene','linalool','geraniol','citronellol','nerol','menthol','borneol','carveol','terpineol','caryophyllene','farnesol','bisabolol','camphene','terpinene','phellandrene','sabinene','carene','thujone','cedrol','fenchone','camphor'];
function autoClassify(name){ const n=name.toLowerCase();
  if(n.includes('lactone')||n.includes('内酯')) return '内酯类';
  if(n.includes('pyrazine')||n.includes('吡嗪')) return '吡嗪类';
  if(n.includes('thiazole')||n.includes('thiazoline')||n.includes('噻唑')) return '噻唑类 / 噻唑啉类';
  if(['sulf','thiol','mercapt','disulfide','trisulfide','硫','furfurylthiol','furanthiol'].some(h=>n.includes(h))) return '含硫化合物';
  if(n.includes('furan')||n.includes('呋喃')||n.includes('furaneol')||n.includes('furanone')||n.includes('furfural')||n.includes('furfuryl')) return '呋喃/呋喃酮类';
  const NHET=['pyridine','pyrrole','pyrrolidine','pyrimidine','indole','quinoline','oxazole','isoxazole','pyrazole','imidazole','azole','morpholine','piperidine','piperazine','吡啶','吡咯','吲哚','喹啉','恶唑','异恶唑','吡唑','咪唑','氮杂','吗啉','哌啶','哌嗪'];
  if(NHET.some(h=>n.includes(h))) return '含氮杂环/其他含氮化合物';
  if(TERP.some(h=>n.includes(h))) return '萜烯类与含氧萜类';
  if(n.includes('aldehyde')||n.includes('醛')||/anal$/.test(n)||/\bal$/.test(n)) return '醛类';
  if(n.includes('acid')||n.includes('oic')||n.includes('羧酸')||n.includes('酸')) return '酸类';
  if(n.includes('ester')||n.includes('ate')||n.includes('酯')) return '酯类';
  if(n.includes('ketone')||n.includes('酮')||/\bone$/.test(n)||n.includes('diacetyl')||n.includes('acetoin')) return '酮类';
  if(n.includes('alcohol')||n.includes('醇')||/(an|en|in|ol)ol$/.test(n)||n.endsWith('ol')) return '醇类';
  return '其他';
}
const DB={}; DB_RAW.forEach(c=>{ [c.en,...(c.syn||[]),c.cn].forEach(k=>{ const nk=normalize(k);
  // 跳过退化键：空 / 纯数字(如 '2') / 单字符(如 '醇')，否则「包含匹配」会让这些键成为大量查询的子串而批量误命中
  if(nk && !/^\d+$/.test(nk) && nk.length>=2) DB[nk]=c; }); });
const KEYS=Object.keys(DB);
// 信息完整度：优先展示「有风味描述 + 有阈值 + 有来源」的物质（阈值权重2，其余各1，满分4）
function hasField(c,k){ const v=c&&c[k]; if(v==null) return false; const s=String(v).trim(); return s!==''&&!['—','-','–','暂无','无','nan','None','null'].includes(s); }
function completeness(c){ if(!c) return 0; return (hasField(c,'odor')?1:0)+(hasField(c,'thr')?2:0)+(hasField(c,'source')?1:0); }
function compLabel(c){ const s=completeness(c); return s>=4?'信息完整':(s>=2?'部分缺失':'信息待补'); }
function sortByCompleteness(rows){ return (rows||[]).slice().sort((a,b)=>completeness(b)-completeness(a)); }
function ratio(a,b){ const m=a.length,n=b.length; if(!m||!n) return 0;
  const dp=Array.from({length:m+1},()=>new Array(n+1).fill(0));
  for(let i=1;i<=m;i++) for(let j=1;j<=n;j++) dp[i][j]= a[i-1]===b[j-1]? dp[i-1][j-1]+1 : Math.max(dp[i-1][j],dp[i][j-1]);
  return 2*dp[m][n]/(m+n);
}
// 基于原名的词边界信息：b4[i]=True 表示 nq[i] 之前是分隔符/行首（词边界）
function _buildNormBoundary(name){
  const nq=[]; const b4=[]; let prevSep=true; const s=name||'';
  for(let i=0;i<s.length;i++){ const ch=s[i].toLowerCase();
    if(/[a-z0-9一-鿿]/.test(ch)){ nq.push(ch); b4.push(prevSep); prevSep=false; } else { prevSep=true; } }
  return {nq:nq.join(''), b4:b4};
}
// k 是否在 nq 中以「整词」出现（前后均为词边界）；返回 'start'(词首,优先)/'end'(词尾)/null
function _boundaryHit(nq,b4,k){
  if(!k) return null; const L=k.length, N=nq.length; let i=nq.indexOf(k);
  while(i>=0){ if(b4[i] && (i+L>=N || b4[i+L])) return b4[i]?'start':'end'; i=nq.indexOf(k,i+1); }
  return null;
}
// 单个候选：精确 -> 整词边界包含(词首优先,最长键) -> 模糊
function _tryMatchJS(q,nq,b4,name,fuzzy){
  if(!q) return null;
  if(DB[q]) return Object.assign({match:'精确',score:1,query:name},DB[q]);
  let preK=null, sufK=null;
  for(const k of KEYS){ if(!k||k.length<3) continue;
    const hit=_boundaryHit(nq,b4,k);
    if(hit==='start'){ if(preK===null||k.length>preK.length) preK=k; }
    else if(hit==='end'){ if(sufK===null||k.length>sufK.length) sufK=k; } }
  if(preK!==null) return Object.assign({match:'包含匹配',score:0.9,query:name},DB[preK]);
  if(sufK!==null) return Object.assign({match:'包含匹配',score:0.9,query:name},DB[sufK]);
  if(fuzzy){ let best=null,bs=0.82; for(const k of KEYS){ const r=ratio(q,k); if(r>bs){bs=r;best=k;} }
    if(best){ const r=ratio(q,best); return Object.assign({match:'模糊匹配',score:Math.round(r*100)/100,query:name},DB[best]); } }
  return null;
}
function match(name){
  const q=normalize(name);
  if(!q) return {en:name,cn:'(空)',cat:'其他',thr:'—',med:'—',odor:'',source:'—',match:'无效',score:0,query:name};
  // 候选查询：全名 -> 各括号片段 -> 去括号后的主体（去重）
  const nb=_buildNormBoundary(name);
  const cands=[q];
  const segs=(name||'').match(/\(([^()]*)\)/g)||[];
  segs.forEach(function(s){ const seg=s.replace(/[()]/g,''); const qs=normalize(seg); if(qs.length>=3) cands.push(qs); });
  const base=normalize((name||'').replace(/\([^()]*\)/g,' '));
  if(base && base!==q) cands.push(base);
  const seen={};
  for(let i=0;i<cands.length;i++){ const cand=cands[i]; if(seen[cand]) continue; seen[cand]=1;
    // 模糊仅在全名候选上运行（控制开销）；括号/主体片段靠精确+整词包含覆盖
    const hit=_tryMatchJS(cand, nb.nq, nb.b4, name, cand===q);
    if(hit) return hit; }
  const cat=autoClassify(name);
  return {en:name,cn:'(未收录)',cat:cat,thr:'—',med:'—',odor:'',source:'—',match:'未匹配(自动分类)',score:0,query:name};
}
const SUP={'⁰':'^0','¹':'^1','²':'^2','³':'^3','⁴':'^4','⁵':'^5','⁶':'^6','⁷':'^7','⁸':'^8','⁹':'^9'};
function toNum(p){ const m=p.match(/([0-9.]+)\s*[x×]\s*10\^?\s*([0-9]+)/); if(m) return parseFloat(m[1])*Math.pow(10,parseFloat(m[2])); return parseFloat(p); }
function parseThr(s){ if(typeof s!=='string') return null; let t=s.trim();
  if(['—','-','', '暂无','无'].includes(t)) return null;
  t=t.split(/[（(]/)[0].replace(/≈|~|\s/g,'');
  if(['—','-',''].includes(t)) return null;
  t=t.split('').map(ch=>SUP[ch]||ch).join('');
  if(t.includes('–')||t.includes('-')){ const vals=[]; t.split(/[–\-]/).forEach(p=>{p=p.trim(); if(p){const v=toNum(p); if(!isNaN(v)) vals.push(v);}}); return vals.length? vals.reduce((a,b)=>a+b,0)/vals.length : null; }
  const v=toNum(t); return isNaN(v)?null:v;
}
function computeOAV(rows){ const pairs=[];
  rows.forEach(r=>{ const T=parseThr(r.thr); if(T==null||!(T>0)){ r.oav=null; r.roav=null; r.oav_flag='—'; return; }
    let val=null; if(typeof r.conc_ugkg==='number'&&r.conc_ugkg!=null) val=r.conc_ugkg; else if(typeof r.conc==='number'&&r.conc!=null) val=r.conc; else if(typeof r.rr==='number'&&r.rr!=null) val=r.rr;
    if(val==null){ r.oav=null; r.roav=null; r.oav_flag='—'; return; } r.oav=val/T; pairs.push(r); });
  const mx=pairs.length? Math.max.apply(null,pairs.map(r=>r.oav)) : 0;
  pairs.forEach(r=>{ r.roav=mx? r.oav/mx*100 : 0; if(r.roav>=10) r.oav_flag='关键致香'; else if(r.roav>=1) r.oav_flag='潜在贡献'; else r.oav_flag='—'; });
}
// 内标法浓度: C = (内标浓度 × 内标体积 × 响应比) / 样品取样量
// 响应比 rr = 化合物峰面积/内标峰面积(无量纲)，已含内标峰面积，故不再除以内标峰面积
// 单位换算: 把 (内标浓度单位 × 内标体积单位)/样品量单位 折算到 μg/L（水相近似）
const _MASS={'pg':1e-12,'ng':1e-9,'µg':1e-6,'μg':1e-6,'ug':1e-6,'mg':1e-3,'g':1.0,'kg':1e3};
const _VOL={'pl':1e-12,'nl':1e-9,'µl':1e-3,'µL':1e-3,'μl':1e-3,'μL':1e-3,'ul':1e-3,'UL':1e-3,'mL':1.0,'ml':1.0,'L':1e3,'l':1e3};
function jsParseConcUnit(expr){ if(!expr) return null;
  let s=expr.replace(/[ ×xX()（）]/g,''); if(!s) return null;
  let conv=1;
  for(let tok of s.split('*')){
    let sign=1;
    if(tok.startsWith('/')){ sign=-1; tok=tok.slice(1); }
    if(!tok) continue;
    if(tok.indexOf('/')>=0){ // 复合单位 mass/vol、mass/mass、vol/mass
      const parts=tok.split('/'), num=parts[0], den=parts[1];
      if(_MASS[num]!=null && _VOL[den]!=null) conv*=Math.pow(_MASS[num],sign)*Math.pow(_VOL[den],-sign);
      else if(_MASS[num]!=null && _MASS[den]!=null) conv*=Math.pow(_MASS[num],sign)*Math.pow(_MASS[den],-sign);
      else if(_VOL[num]!=null && _MASS[den]!=null) conv*=Math.pow(_VOL[num],sign)*Math.pow(_MASS[den],-sign);
      else return null;
    } else if(_MASS[tok]!=null) conv*=Math.pow(_MASS[tok],sign);
    else if(_VOL[tok]!=null) conv*=Math.pow(_VOL[tok],sign);
    else return null;
  }
  return conv*1e9; // 折算到 μg/L（水相近似，μg/kg≈μg/L）
}
function conc(cis,vis,resp,ms,cisu,visu,msu){
  const a=parseFloat(cis),b=parseFloat(vis),c=parseFloat(resp),d=parseFloat(ms);
  if([a,b,c,d].some(v=>isNaN(v))||d===0) return null;
  const raw=a*b*c/d;
  const unitExpr='('+(cisu||'mg/L')+')*('+(visu||'μL')+')/('+(msu||'g')+')';
  const k=jsParseConcUnit(unitExpr);
  return (k==null)? null : raw*k;
}
// 浓度单位换算系数: 原单位 -> 目标单位（基于 μg/kg = 1 基准）
const _CU={ 'μg/kg':1, 'mg/kg':1e3, 'ng/g':1, 'ng/kg':1e-3, 'g/kg':1e6 };
function concUnitFactor(src,dst){ const a=_CU[src], b=_CU[dst]; if(a==null||b==null) return null; return b/a; }
function catCls(cat){ return 'b-'+(cat||'其他').replace(/[\\/\\s]/g,''); }
function badge(cat){ return '<span class="badge '+catCls(cat)+'">'+(cat||'其他')+'</span>'; }
function srcFmt(s){ return String(s??'').split('\n').map(l=>{ const m=l.match(/^([^：:]+)[：:]/); return m? '<b>'+m[1]+'</b>'+l.slice(m[1].length) : l; }).join('<br>'); }
function mcls(m){ if(m==='精确')return 'ok'; if(m.includes('自动'))return 'bad'; return 'warn'; }
function escapeHtml(s){ return (s||'').replace(/[&<>"]/g,c=>({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;'}[c])); }

// ---------- 摘要（类别卡片） ----------
function renderSummary(rows, opts){ opts=opts||{}; if(!rows||!rows.length) return '';
  const total=rows.length; const cats={}; rows.forEach(r=>{ if(r.cat) cats[r.cat]=(cats[r.cat]||0)+1; });
  const catCount=Object.keys(cats).length;
  const catArr=Object.entries(cats).sort((a,b)=>b[1]-a[1]); const maxCat=catArr.length?catArr[0][1]:1;
  let catHtml='<div class="cat-bars"><h3><span class="ic">▤</span>类别覆盖<span class="sub-note">（共 '+catCount+' 类 · 点击类别展开物质）</span></h3><div class="cat-grid">';
  catArr.forEach(([k,v])=>{ catHtml+='<div class="cat-chip cat-row" data-cat="'+k+'" onclick="filterSummaryCat(this)" title="点击查看「'+k+'」全部物质">'
    +'<span class="cc-head"><span class="cc-dot '+catCls(k)+'"></span><span class="cc-name">'+k+'</span></span>'
    +'<span class="cc-num">'+v+'<small>种</small></span></div>'; });
  catHtml+='</div></div>';
  return catHtml+'<div class="cat-detail" id="cat-detail"></div>';
}
let _lastSummaryRows=null; function _setSummaryRows(r){ _lastSummaryRows=r; }
function filterSummaryCat(el){ const cat=el.getAttribute('data-cat'); if(!cat) return; const box=document.getElementById('cat-detail'); if(!box) return;
  document.querySelectorAll('.cat-row').forEach(c=>c.classList.remove('is-active'));
  if(box._cat===cat){ box.innerHTML=''; box._cat=null; return; }
  el.classList.add('is-active'); box._cat=cat;
  // 先按信息完整度（有风味描述/阈值/来源优先），再按 OAV 由大到小
  const rows=(_lastSummaryRows||[]).filter(r=>r.cat===cat).slice().sort((a,b)=>{
    const dc=completeness(b)-completeness(a); if(dc!==0) return dc;
    const av=r=>(typeof r.oav==='number'&&!isNaN(r.oav))?r.oav:0; return av(b)-av(a); });
  const flsOf=r=>{ const f=r.oav_flag||'—'; return f==='关键致香'?'f-key':(f==='潜在贡献'?'f-pot':'f-na'); };
  let h='<div class="card" style="margin:0"><h3><span class="ic">▤</span>'+cat+' · 共 '+rows.length+' 种物质（按 OAV 由大到小）</h3><div class="scroll"><table><thead><tr><th>#</th><th>英文名</th><th>中文名</th><th>类别</th><th>阈值(μg/L)</th><th>浓度(μg/L)</th><th>OAV</th><th>ROAV</th><th>气味活性</th><th>气味描述</th></tr></thead><tbody>';
  h+=rows.map((r,i)=>{ const cu=(r.conc==null||isNaN(r.conc))?'—':Number(r.conc).toPrecision(4).replace(/\.?0+$/,'');
    const oav=(r.oav==null)?'—':r.oav.toPrecision(4).replace(/\.?0+$/,''); const roav=(r.roav==null)?'—':r.roav.toFixed(2);
    const flag=r.oav_flag||'—'; return '<tr><td>'+(i+1)+'</td><td><b>'+r.en+'</b></td><td class="cn">'+r.cn+'</td><td>'+badge(r.cat)+'</td><td>'+r.thr+'</td><td class="num-dark">'+cu+'</td><td class="num-dark">'+oav+'</td><td class="num-dark">'+roav+'</td><td class="oavflag '+flsOf(r)+'">'+flag+'</td><td class="desc">'+r.odor+'</td></tr>'; }).join('');
  h+='</tbody></table></div></div>'; box.innerHTML=h;
}

// ---------- 表格 ----------
function tableHTML(rows){ if(!rows.length) return '<p class="hint">无结果</p>';
  rows=sortByCompleteness(rows);  // 信息完整（有气味描述/阈值/来源）优先展示
  let h='<div class="scroll"><table><thead><tr><th>#</th><th>英文名</th><th>中文名</th><th>CAS</th><th>类别</th><th>阈值(μg/L)</th><th>介质</th><th>匹配</th><th>气味描述</th><th>来源</th></tr></thead><tbody>';
  rows.forEach((r,i)=>{
    h+='<tr><td>'+(i+1)+'</td><td><b>'+r.en+'</b></td><td class="cn">'+r.cn+'</td><td>'+(r.cas||'')+'</td><td>'+badge(r.cat)+'</td><td>'+r.thr+'</td><td>'+r.med+'</td><td class="m '+mcls(r.match)+'">'+r.match+'</td><td class="desc">'+(r.odor||'')+'</td><td class="src">'+srcFmt(r.source)+'</td></tr>'; });
  h+='</tbody></table></div>'; return h;
}

// ---------- ① 单物质 ----------
function byQ(v){ document.getElementById('q').value=v; doSingle(); }
function doSingle(){ const q=document.getElementById('q').value.trim(); if(!q) return;
  const r=match(q); _setSummaryRows([r]);
  document.getElementById('single-out').innerHTML=renderSummary([r],{scope:'单物质查询'})+'<p class="hint">命中 1 种。若需 OAV/ROAV，请在「⑤ 浓度计算」或「③ 报告解析」中提供浓度/响应比。</p>'+tableHTML([r]);
}
// ---------- ② 批量 ----------
function loadBatchFile(){ const el=document.getElementById('batchfile'); const f=el.files[0]; if(!f) return; const rd=new FileReader();
  rd.onload=e=>{ const names=(e.target.result||'').split(/\r?\n/).map(s=>s.trim()).filter(Boolean); document.getElementById('batch').value=names.join('\n'); el.value=''; if(!names.length) alert('未识别到物质名'); };
  rd.readAsText(f,'utf-8');
}
function doBatch(){ const txt=document.getElementById('batch').value; const names=txt.split(/\n|\r/).map(s=>s.trim()).filter(Boolean); if(!names.length) return;
  const rows=names.map(n=>{ const p=n.split(/[,，]/); const m=match(p[0].trim()); if(p[1]!==undefined){ const v=parseFloat(p[1]); if(!isNaN(v)) m.rr=v; } return m; });
  _setSummaryRows(rows); document.getElementById('batch-out').innerHTML=renderSummary(rows,{scope:'批量检索'})+tableHTML(rows);
}

// ---------- ③ 报告 ----------
let lastReport=[];
var reportCatFilter='';  // GC-MS 主数据表（report-table）类别筛选；''=全部
function applyReportConc(){
  const cis=document.getElementById('r-cis').value.trim();
  const vis=document.getElementById('r-vis').value.trim();
  const ms=document.getElementById('r-ms').value.trim();
  const cisu=document.getElementById('r-cisu').value.trim()||'mg/L';
  const visu=document.getElementById('r-visu').value.trim()||'μL';
  const msu=document.getElementById('r-msu').value.trim()||'g';
  const ok = cis!==''&&vis!==''&&ms!==''&&!isNaN(parseFloat(cis))&&!isNaN(parseFloat(vis))&&!isNaN(parseFloat(ms))&&parseFloat(ms)!==0;
  lastReport.forEach(r=>{
    if(ok){
      const resp=(typeof r.rr==='number'&&r.rr!=null)?r.rr:null;
      const c=resp!=null?conc(cis,vis,resp,ms,cisu,visu,msu):null;
      r.conc=c; r.conc_unit='μg/L'; r.conc_ugkg=c;
    } else if(typeof r.conc0==='number'&&r.conc0!=null){ r.conc=r.conc0; r.conc_unit=(r.conc_unit||'μg/L'); r.conc_ugkg=r.conc0; }
    else { r.conc=null; r.conc_unit=null; r.conc_ugkg=null; }
  });
  lastReport._conc_ok=ok; lastReport._conc_unit=ok?'μg/L':null;
  computeOAV(lastReport);
}
function recalcReportConc(){ if(!lastReport.length){alert('请先解析报告');return;} applyReportConc(); renderReport(); }
function renderReport(){ if(!lastReport.length){ document.getElementById('report-out').innerHTML='<p class="hint">无结果</p>'; return; }
  let h='';
  const istA=(lastReport&&lastReport.istdArea!=null)?lastReport.istdArea:(lastReport&&lastReport[0]&&lastReport[0].istdArea!=null?lastReport[0].istdArea:null);
  if(istA!=null){ const istName=document.getElementById('r-ist').value.trim()||'内标'; h+='<p class="hint" style="color:var(--teal)">已识别内标物质「'+istName+'」的峰面积 A₁ = '+istA+'（rr 已含该值，公式不再单独除以 A₁）。</p>'; }
  h+=renderSummary(lastReport,{scope:'GC-MS 报告',noKpi:true}); _setSummaryRows(lastReport);
  // 类别筛选（下拉单选，与多报告模块风格一致）
  const cats=[...new Set(lastReport.map(r=>r.cat||'其他'))];
  const catOrder=(typeof CAT!=='undefined'&&CAT)?(Array.isArray(CAT)?CAT.map(c=>c.name||c):Object.keys(CAT)):[];
  cats.sort((a,b)=>{var ia=catOrder.indexOf(a),ib=catOrder.indexOf(b);ia=ia<0?999:ia;ib=ib<0?999:ib;return ia-ib;});
  if(reportCatFilter && cats.indexOf(reportCatFilter)<0) reportCatFilter='';
  const fsel=reportCatFilter||'';
  let fbar='<div class="gc-catbar" style="margin:8px 0 4px"><label>类别筛选：<select id="r-cat-filter" onchange="reportCatFilter=this.value;renderReport();"><option value="">全部类别</option>';
  cats.forEach(c=>{ fbar+='<option value="'+c+'"'+(c===fsel?' selected':'')+'>'+c+'</option>'; });
  fbar+='</select></label><span class="hint">按类别筛选主数据表物质</span></div>';
  h+=fbar;
  const rows=lastReport.filter(r=> !reportCatFilter || (r.cat||'其他')===reportCatFilter);
  h+='<div class="scroll"><table id="report-table"><thead><tr><th>#</th><th>英文名</th><th>中文名</th><th>CAS</th><th>类别</th><th>RT(min)</th><th>响应比 rr</th><th>阈值(μg/L)</th><th>浓度(μg/L)</th><th>OAV</th><th>ROAV</th><th>气味活性</th><th>气味描述</th><th>来源</th></tr></thead><tbody>';
  rows.forEach((r,i)=>{ const cu=(r.conc==null||isNaN(r.conc))?'—':Number(r.conc).toPrecision(4).replace(/\.?0+$/,'')+' '+(r.conc_unit||'μg/L');
    const oav=(r.oav==null)?'—':r.oav.toPrecision(4).replace(/\.?0+$/,''); const roav=(r.roav==null)?'—':r.roav.toFixed(2);
    const flag=r.oav_flag||'—'; const fls=flag==='关键致香'?'f-key':(flag==='潜在贡献'?'f-pot':'f-na');
    h+='<tr><td>'+(i+1)+'</td><td><b>'+r.en+'</b></td><td class="cn">'+r.cn+'</td><td>'+(r.cas||'')+'</td><td>'+badge(r.cat)+'</td><td>'+(r.rt||'')+'</td><td>'+(r.rr||'')+'</td><td>'+r.thr+'</td><td>'+cu+'</td><td>'+oav+'</td><td>'+roav+'</td><td class="oavflag '+fls+'">'+flag+'</td><td class="desc">'+(r.odor||'')+'</td><td class="src">'+srcFmt(r.source)+'</td></tr>'; });
  h+='</tbody></table></div>'; document.getElementById('report-out').innerHTML=h;
}
function recalcReportConc(){ if(!lastReport.length){alert('请先解析报告');return;} applyReportConc(); renderReport(); }
function parseXlsxRows(buf){ const wb=XLSX.read(buf,{type:'array'}); const ws=wb.Sheets[wb.SheetNames[0]]; const js=XLSX.utils.sheet_to_json(ws,{header:1});
  let hi=-1; for(let i=0;i<js.length;i++){ if(js[i]&&js[i].some(c=>typeof c==='string'&&c.trim()==='化合物')){hi=i;break;} } if(hi<0) return null;
  const head=js[hi].map(c=>typeof c==='string'?c.trim():''); const ci=head.indexOf('化合物'),ri=head.indexOf('响应比');
  const rows=[]; for(let i=hi+1;i<js.length;i++){ const r=js[i]; if(!r||!r[ci]||typeof r[ci]!=='string'||!r[ci].trim()) continue;
    const comp=r[ci].replace(/\n/g,' ').trim(); if(['样品类型','化合物','定量分析完成报告','样品色谱图'].includes(comp)) continue;
    const m=match(comp); const v=r[ri]; m.rr=(typeof v==='number')? v : undefined; rows.push(m); } return rows;
}
// 归一化并记录「归一化下标 -> 原文下标」的映射，便于回原文取数（保留小数点）
// 归一化并记录「归一化下标 -> 原文下标」的映射，便于回原文取数（保留小数点）
function normalizeMap(s){ if(typeof s!=='string') return {n:'',m:[]}; const ch=[],mp=[];
  for(let i=0;i<s.length;i++){ let c=s[i].toLowerCase(); if(c==='\u2019'||c==='\u2018') c="'";
    if(c.length===1&&((c>='a'&&c<='z')||(c>='0'&&c<='9')||(c>='\u4e00'&&c<='\u9fff'))){ ch.push(c); mp.push(i); } }
  return {n:ch.join(''),m:mp}; }
// 归一化名 -> 物质对象 的索引
let _NAME_IDX=null;
function _nameIdx(){ if(_NAME_IDX) return _NAME_IDX; const idx={};
  DB_RAW.forEach(c=>{ const names=[c.en].concat(c.syn||[]).concat(c.cn?[c.cn]:[]);
    names.forEach(nm=>{ if(!nm) return; const nk=normalize(nm); if(!nk) return;
      // 剔除碎片候选：纯数字跳过；ASCII 名要求 ≥3 字符；中文名要求 ≥2 字（保留「丁酸/己醛」）
      if(/^[0-9]+$/.test(nk)) return;
      const ascii=/^[\x00-\x7F]+$/.test(nk);
      if(ascii? nk.length<3 : nk.length<2) return;
      if(!(nk in idx)) idx[nk]=c; }); });
  _NAME_IDX={idx:idx,cands:Object.keys(idx).sort((a,b)=>b.length-a.length)}; return _NAME_IDX; }
// 数值列标签（按优先级）
const RR_LABELS=['响应比','响应','含量','浓度','ratio','resp','conc'];
const RR_HEAD=['响应比','含量','浓度','ratio','conc','响应','resp'];
const NON_VALUE_HEAD=['单位','unit','名称','name','cas','定性','qual','compound','化合物','ion','离子'];
function _lineAt(t,p){ const e=t.indexOf('\n',p); const s=t.lastIndexOf('\n',p-1)+1; return t.slice(s,e<0?t.length:e); }
// 只取物质名之后的行尾，避免把「2-庚酮」里的 2 当成数值
function _lineTail(t,p){ const e=t.indexOf('\n',p); return t.slice(p,e<0?t.length:e); }
// 公式字母 -> 表头候选词（rr 置于 area 前，避免「响应比」误判为「响应」）
// 注意顺序：istd 必须排在 area 之前——「内标峰面积」含「峰面积」，若 area 在前会
// 先抢占到该列，导致 areaPos 错位（误把内标峰面积 A₁ 当成化合物峰面积 A 填入响应值栏）。
const COL_SPECS = {
  rt:   ['保留时间','retention','rt','出峰时间','time'],
  // 注意：istd 只匹配「内标峰面积」类数值列，不再含裸 'istd'/'内标'（那是内标名列，非数值）
  istd: ['istd响应','istd 响应','内标峰面积','内标响应','istd面积','istd 面积'],
  rr:   ['响应比','相对响应','相对峰面积','ratio','rr'],
  conc: ['最终浓度','浓度','含量','conc','content'],
  area: ['峰面积','响应值','响应','面积','area','peak','response'],
};
// 合并 'ISTD 响应' 这类被空白拆开的两词表头为单个单元格
function _splitHeader(line){ const cells=line.trim().split(/\s+|\|/).map(c=>c.trim()).filter(Boolean);
  const out=[]; let i=0;
  while(i<cells.length){ const c=cells[i], cl=c.toLowerCase();
    if((cl==='istd'||cl==='内标') && i+1<cells.length){ const nxt=cells[i+1].toLowerCase();
      if(['响应','response','面积'].some(k=>nxt.indexOf(k)>=0)){ out.push(c+' '+cells[i+1]); i+=2; continue; } }
    out.push(c); i++; }
  return out; }
// 返回 {pos:{role:0-based表头单元格索引}, ncols:表头单元格数}；与数据行按尾部对齐取数
function _detectColumns(text){ const lines=(text||'').split('\n').slice(0,80);
  for(const line of lines){ const low=line.toLowerCase();
    if(low.indexOf('化合物')<0 && low.indexOf('compound')<0 && low.indexOf('物质')<0) continue;
    if(!['rt','响应','面积','浓度','峰','ratio','retention'].some(lb=>low.indexOf(lb)>=0)) continue;
    const cells=_splitHeader(line);
    if(cells.length<3) continue;
    let nameI=cells.findIndex(c=>['化合物','compound','物质'].includes(c.toLowerCase())); if(nameI<0) nameI=0;
    const pos={};
    for(let j=0;j<cells.length;j++){ if(j===nameI) continue; const c=cells[j].toLowerCase();
      if(NON_VALUE_HEAD.some(u=>c.indexOf(u)>=0)) continue;
      for(const f in COL_SPECS){ if(COL_SPECS[f].some(lb=>c.indexOf(lb)>=0)){ if(!(f in pos)) pos[f]=j; break; } } }
    if(Object.keys(pos).length) return {pos:pos, ncols:cells.length};
  }
  return {pos:{}, ncols:0}; }
// 按表头单元格索引从数据行取第 idx 列数值；数据行多出的 token（化合物名含空格）算到行首偏移
function _valAt(line, idx, ncols){ if(idx==null||ncols<=0) return null;
  const toks=line.split(/\s+/).filter(Boolean); let off=toks.length-ncols; if(off<0) off=0;
  const ti=idx+off; if(ti>=0&&ti<toks.length){ const s=toks[ti].trim();
    if(/^-?\d+(?:\.\d+)?(?:[eE][-+]?\d+)?$/.test(s)) return parseFloat(s); } return null; }
// 按表义识别各列并提取，映射内标法公式字母：A=峰面积 rr=响应比 c=浓度 RT=保留时间；A₁ 从内标列取
function parseReportText(text, istName){
  if(!text) return [];
  const {idx,cands}=_nameIdx(); const {n:norm,m:pmap}=normalizeMap(text);
  if(!norm) return [];
  const {pos:cols,ncols}=_detectColumns(text);
  const areaPos=cols.area, rrPos=cols.rr, concPos=cols.conc, rtPos=cols.rt, istdPos=cols.istd;
  let istdArea=null;
  const seen={}; const taken=[]; const rows=[];
  const overlap=(a,b)=>!(b[1]<=a[0]||b[0]>=a[1]);
  // 词边界判断（基于原文 text 与 pos_map 映射，避免「芳樟醇」误嵌在「反式-呋喃型氧化芳樟醇」中）。
  // 仅拒字母/数字/中文紧邻（连字符、括号、逗号等已被归一化剥离，属同名内部分隔符，不能算非边界，否则会误杀「β-月桂烯」等行首名）
  const _atWordBoundary=(txt,o_s,o_e)=>{ if(o_s>0){ const pr=txt[o_s-1]; if((pr>='a'&&pr<='z')||(pr>='0'&&pr<='9')||(pr>='\u4e00'&&pr<='\u9fff')) return false; }
    const nx=(o_e+1<txt.length)?txt[o_e+1]:''; if(nx && ((nx>='a'&&nx<='z')||(nx>='0'&&nx<='9')||(nx>='\u4e00'&&nx<='\u9fff'))) return false; return true; };
    for(let i=0;i<cands.length;i++){ const nk=cands[i]; const c=idx[nk]; if(seen[c.en]) continue;
    // 找第一个「不与已占用区间重叠」且「处于词边界」的位置（词边界避免「芳樟醇」误嵌在「反式-呋喃型氧化芳樟醇」中）
    let st=0,pos=-1;
    for(;;){ const p=norm.indexOf(nk,st); if(p<0) break;
      const os=pmap[p], oe=pmap[p+nk.length-1];
      if(_atWordBoundary(text,os,oe) && !taken.some(iv=>overlap([p,p+nk.length],iv))){ pos=p; break; } st=p+1; }
    if(pos<0) continue;
    taken.push([pos,pos+nk.length]); seen[c.en]=1;
    const s=pmap[pos+nk.length-1]+1;
    // 取该物质所在的完整数据行（含全部数值列）
    const ls=text.lastIndexOf('\n',s)+1; let e=text.indexOf('\n',s); const line=text.slice(ls, e<0?undefined:e);
    const rt=_valAt(line,rtPos,ncols);
    const resp=_valAt(line,areaPos,ncols);   // A 化合物峰面积（响应）
    const rr=_valAt(line,rrPos,ncols);       // rr 响应比
    const conc0=_valAt(line,concPos,ncols);  // c 浓度（报告自带）
    rows.push(Object.assign({match:'文本匹配',score:0.9,query:c.en,rt:rt,resp:resp,rr:rr,conc0:conc0},c));
    if(istdArea==null && istdPos!=null) istdArea=_valAt(line,istdPos,ncols); }
  rows.istdArea=istdArea;
  if(istdArea!=null) for(const r of rows) r.istdArea=istdArea;
  return rows; }

// 统一 PDF 文本抽取（pdf.js 走主线程 fake worker，file:// 下同样可用）
// 内联 CMaps 以支持中文 / 各类 CID 字体 PDF 的离线文本抽取
class OfflineCMapFactory {
  constructor(options){ this.isCompressed = true; }
  async fetch({name}){
    const fn = name + '.bcmap';
    const key = CMAP_DATA && (CMAP_DATA[name] || CMAP_DATA[fn]);
    if(!key) throw new Error('CMAP 缺失: ' + name);
    const bin = atob(key); const arr = new Uint8Array(bin.length);
    for(let i=0;i<bin.length;i++) arr[i] = bin.charCodeAt(i);
    return { cMapData: arr, compressionType: 1 };
  }
}
function _mtxMul(m1,m2){ return [m1[0]*m2[0]+m1[2]*m2[1], m1[1]*m2[0]+m1[3]*m2[1], m1[0]*m2[2]+m1[2]*m2[3], m1[1]*m2[2]+m1[3]*m2[3], m1[0]*m2[4]+m1[2]*m2[5]+m1[4], m1[1]*m2[4]+m1[3]*m2[5]+m1[5]]; }
async function extractPdfText(file){ if(typeof pdfjsLib==='undefined') throw new Error('PDF 解析库未加载');
  const buf=await file.arrayBuffer();
  const opts={ data:new Uint8Array(buf), cMapPacked:true };
  if(typeof OfflineCMapFactory==='function') opts.CMapReaderFactory=OfflineCMapFactory;
  const pdf=await pdfjsLib.getDocument(opts).promise; let t='';
  // 旋转感知：把每页按「显示方向」归类成行（解决 90° 旋转报告被 pdf.js 抽到一列、列错位的问题）
  for(let i=1;i<=pdf.numPages;i++){
    const pg=await pdf.getPage(i);
    const vp=pg.getViewport({scale:1, rotation:pg.rotate||0});   // 含页面旋转的显示坐标系
    const c=await pg.getTextContent();
    const items=[];
    for(const it of c.items){
      const str=(it.str||'').trim(); if(!str) continue;
      const tf=(it.transform&&it.transform.length===6)?it.transform:[1,0,0,1,(it.x||0),(it.y||0)];
      const tx=_mtxMul(vp.transform, tf);   // 显示坐标(x右, y下)
      items.push({x:tx[4], y:tx[5], str});
    }
    if(!items.length){ t+='\n'; continue; }
    const ys=[...new Set(items.map(o=>Math.round(o.y)))].sort((a,b)=>a-b);
    const rowTol=Math.max(2.5, (ys[ys.length-1]-ys[0])/Math.max(1,ys.length)*0.7);
    const rows=[]; let cur=null, curY=null;
    // 显示坐标 y 向下：升序(顶部优先)=阅读顺序；行内按 x 升序(左→右)
    for(const o of items.slice().sort((p,q)=> (p.y-q.y) || (p.x-q.x))){
      if(curY==null || Math.abs(o.y-curY)>rowTol){ cur=[]; rows.push(cur); curY=o.y; }
      cur.push(o);
    }
    const leftX=Math.min(...items.map(o=>o.x));
    let pageText='';
    for(const r of rows){
      r.sort((p,q)=>p.x-q.x);
      const indent=(r[0].x-leftX)>10 ? ' ' : '';   // 化合物名列(右移)视为缩进行，复刻 pdftotext -layout 缩进
      pageText+=indent+r.map(o=>o.str).join(' ')+'\n';
    }
    t+=pageText;
  }
  return t; }
// 浏览器端「岛津多处理组定量报告」解析（离线无 pdftotext，pdf.js 旋转感知抽取后在此复用与后端一致的逻辑）
function _gNum(x){ x=(x||'').replace(/,/g,'').trim(); if(['ND','nd','N.D.','N/A','—','-','–',''].includes(x)) return null; const v=parseFloat(x); return isNaN(v)?null:v; }
const _SAMPLE_RE=/^([A-Za-z0-9_.\-]+\.d)\b/;
function parse_gcms_pdf_multi_js(text){
  if(!text) return {multi:true, compounds:[], groups:[], samples:[], schema:'shimadzu_quant_multi'};
  const lines=text.split('\n');
  const sample_to_name={}, samples=[], ist_area_by_group={}; let ist_name_detected=null;
  const piv={}; let cur=null, frags=[], header_open=false, in_quant=false;
  for(const line of lines){
    const s=line.replace(/\s+$/,''); const st=s.trim();
    if(!st) continue;
    if(st.includes('定量结果') && !st.includes('数据文件')){ in_quant=true; continue; }
    const m0=_SAMPLE_RE.exec(line);
    if(!in_quant && m0){ const samp=m0[1]; const toks=line.trim().split(/\s+/);
      if(toks.length>=2 && samp===toks[0] && !(samp in sample_to_name)){ sample_to_name[samp]=toks[1]; samples.push(samp); }
      continue; }
    if(!in_quant) continue;
    if(st.startsWith('数据文件') && st.includes('化合物') && st.includes('ISTD')){ if(header_open && frags.length){ cur=frags.join(' ').trim(); frags=[]; } header_open=false; continue; }
    const m=_SAMPLE_RE.exec(line);
    if(m){ const samp=m[1]; if(samples.indexOf(samp)<0) samples.push(samp);
      if(header_open && frags.length){ cur=frags.join(' ').trim(); frags=[]; header_open=false; }
      const toks=line.trim().split(/\s+/); let istCand=null, nums=[null,null,null,null,null];
      if(toks.indexOf('Sample')>=0){ const i=toks.indexOf('Sample'); if(i>=1 && ist_name_detected==null){ const cand=toks[i-1]; if(cand && cand!=='化合物' && cand!=='ISTD') istCand=cand; }
        for(let k=0;k<5 && i+1+k<toks.length;k++) nums[k]=_gNum(toks[i+1+k]); }
      if(istCand && ist_name_detected==null) ist_name_detected=istCand;
      const mcomp=match(cur||samp); const en=mcomp.en||cur||samp;
      let entry=piv[en]; if(!entry){ entry={en:en, cn:mcomp.cn, cat:mcomp.cat||'其他', thr:mcomp.thr, cas:mcomp.cas, source:mcomp.source, match:mcomp.match, odor:mcomp.odor, by_group:{}}; piv[en]=entry; }
      const grp=sample_to_name[samp]||samp;
      if(!(grp in entry.by_group) || !entry.by_group[grp]){ entry.by_group[grp]={rt:nums[0], resp:nums[1], istd_resp:nums[2], rr:nums[3], final_conc:nums[4], sample:samp}; }
      if(nums[2]!=null && !(grp in ist_area_by_group)) ist_area_by_group[grp]=nums[2];
      continue; }
    if(line[0]===' '||line[0]==='\t'){ if(header_open) frags.push(st); continue; }
    if(header_open && frags.length) cur=frags.join(' ').trim();
    frags=[st]; header_open=true;
  }
  const compounds=Object.values(piv);
  const groups=samples.map(s=>sample_to_name[s]||s);
  if(!compounds.length) return {multi:true, compounds:[], groups, samples, schema:'shimadzu_quant_multi'};
  const rows=[]; const sample_by_group={}; samples.forEach(s=> sample_by_group[sample_to_name[s]||s]=s);
  compounds.forEach(c=> groups.forEach(g=>{ const cell=c.by_group[g]; if(!cell) return;
    const row={en:c.en, cn:c.cn, cat:c.cat, thr:c.thr, cas:c.cas, source:c.source, match:c.match, odor:c.odor, group:g, sample:sample_by_group[g]||g};
    row.rt=cell.rt; row.resp=cell.resp; row.istd_resp=cell.istd_resp; row.rr=cell.rr; row.final_conc=cell.final_conc; rows.push(row); }));
  return {multi:true, schema:'shimadzu_quant_multi', software:'Shimadzu GCMSsolution 定量分析完成报告', groups, samples, ist_name:ist_name_detected, ist_area_by_group, n_compounds:compounds.length, compounds, rows};
}
function doReport(){ const f=document.getElementById('reportfile').files[0]; if(!f) return; const name=f.name.toLowerCase(); const rd=new FileReader();
  if(name.endsWith('.csv')||name.endsWith('.txt')||name.endsWith('.md')||name.endsWith('.markdown')){ document.getElementById('xlsxWarn').style.display='none';
    rd.onload=e=>{ const txt=(e.target.result||''); const lines=txt.split(/\r?\n/).map(s=>s.trim()).filter(Boolean); const rows=[];
      lines.forEach(l=>{ const p=l.split(/[,，\t]/); const m=match(p[0].trim()); const v=parseFloat(p[1]); if(!isNaN(v)) m.rr=v; rows.push(m); });
      lastReport=rows; applyReportConc(); renderReport(); }; rd.readAsText(f,'utf-8'); return; }
  if(name.endsWith('.pdf')){ document.getElementById('xlsxWarn').style.display='none';
    extractPdfText(f).then(t=>{
      const multi=parse_gcms_pdf_multi_js(t);
      if(multi && multi.compounds && multi.groups && multi.groups.length>=2){
        lastMulti=multi; lastReport=[];
        const ie=document.getElementById('r-ist'); if(ie) ie.value=multi.ist_name||ie.value;
        recalcGroupCompare();
      } else {
        lastReport=parseReportText(t, document.getElementById('r-ist').value.trim()); applyReportConc(); renderReport();
      }
    }).catch(err=>{ alert('PDF 解析失败：'+err.message); }); return; }
  if(name.endsWith('.docx')){ if(typeof mammoth==='undefined'){ document.getElementById('xlsxWarn').style.display='block'; alert('Word 解析库未加载（离线环境无法解析 .docx，请另存 CSV）'); return; }
    document.getElementById('xlsxWarn').style.display='none'; rd.onload=async e=>{ try{ const res=await mammoth.extractRawText({arrayBuffer:e.target.result}); lastReport=parseReportText(res.value, document.getElementById('r-ist').value.trim()); applyReportConc(); renderReport(); }catch(err){ alert('Word 解析失败：'+err.message); } }; rd.readAsArrayBuffer(f); return; }
  if(typeof XLSX==='undefined'){ document.getElementById('xlsxWarn').style.display='block'; return; }
  document.getElementById('xlsxWarn').style.display='none'; rd.onload=e=>{ try{ const rows=parseXlsxRows(e.target.result); if(!rows){ alert('未在 Excel 中找到"化合物"表头'); return; }
    lastReport=rows; applyReportConc(); renderReport(); }catch(err){ alert('Excel 解析失败：'+err.message); } }; rd.readAsArrayBuffer(f);
}
document.getElementById('reportfile').addEventListener('change',doReport);

// ---------- ④ 多报告对比 ----------
let _samples=[];
function onMultiFile(e){ const f=e.target.files[0]; if(!f) return; _samples.push({file:f,name:''}); renderSampleList(); e.target.value=''; }
function renderSampleList(){ const box=document.getElementById('sample-list');
  if(!_samples.length){ box.innerHTML='<div class="field" style="flex-basis:100%"><span class="hint">尚未添加报告，点击上方「选择文件」逐个添加。</span></div>'; return; }
  box.innerHTML=_samples.map((s,i)=>'<div class="field" style="flex-basis:100%"><span>样品 '+(i+1)+' 名称 <small style="color:var(--mut-2)">（'+(s.file.name||'')+'）</small></span><div style="display:flex;gap:8px;align-items:center"><input type="text" placeholder="如 样品A / 发酵乳1号" value="'+s.name+'" oninput="_samples['+i+'].name=this.value"><button class="ghost" style="flex:0 0 auto;padding:8px 12px" onclick="removeSample('+i+')">删除</button></div></div>').join(''); }
function removeSample(i){ _samples.splice(i,1); renderSampleList(); }
function clearMulti(){ _samples=[]; renderSampleList(); document.getElementById('multi-out').innerHTML=''; }
let _multiSamples=null;
async function doMulti(){ if(_samples.length<2){ alert('请至少添加 2 个报告进行对比'); return; }
  // 逐个读文件
  const samples=[];
  for(const s of _samples){
    const f=s.file; const name=s.name||f.name||('样品'+(samples.length+1));
    const rows=await readReportFile(f); if(rows&&rows.length) samples.push({name,rows});
  }
  if(!samples.length){ document.getElementById('multi-out').innerHTML='<p class="hint">未解析到任何物质，请检查文件内容。</p>'; return; }
  // 统一内标参数（若填写）按公式 C=内标浓度×内标加量×响应比÷样品量 算浓度(μg/L)
  const cis=document.getElementById('m-cis').value.trim();
  const vis=document.getElementById('m-vis').value.trim();
  const ms=document.getElementById('m-ms').value.trim();
  const cisu=document.getElementById('m-cisu').value.trim()||'mg/L';
  const visu=document.getElementById('m-visu').value.trim()||'μL';
  const msu=document.getElementById('m-msu').value.trim()||'g';
  const ok = cis!==''&&vis!==''&&ms!==''&&!isNaN(parseFloat(cis))&&!isNaN(parseFloat(vis))&&!isNaN(parseFloat(ms))&&parseFloat(ms)!==0;
  samples.forEach(s=>{
    s.rows.forEach(r=>{
      if(ok){
        const resp=(typeof r.rr==='number'&&r.rr!=null)?r.rr:null;
        const c=resp!=null?conc(cis,vis,resp,ms,cisu,visu,msu):null;
        r.conc=c; r.conc_unit='μg/L'; r.conc_ugkg=c;
      } else if(typeof r.conc0==='number' && r.conc0!=null){ r.conc=r.conc0; r.conc_ugkg=r.conc0; r.conc_unit=(r.conc_unit||'μg/L'); }
      else { r.conc=null; r.conc_ugkg=null; }
    });
    computeOAV(s.rows);
    s.keyCount = s.rows.length;
    s.keyFrag = s.rows.filter(r=>r.oav_flag==='关键致香').length;
    s.maxOAV = s.rows.reduce((m,r)=>Math.max(m, (r.oav&&!isNaN(r.oav))?r.oav:0), 0);
  });
  renderMulti(samples);
}
function readReportFile(f){ return new Promise(resolve=>{ const name=f.name.toLowerCase(); const rd=new FileReader();
  if(name.endsWith('.csv')||name.endsWith('.txt')||name.endsWith('.md')||name.endsWith('.markdown')){ rd.onload=e=>{ const rows=[]; (e.target.result||'').split(/\r?\n/).map(s=>s.trim()).filter(Boolean).forEach(l=>{ const p=l.split(/[,，\t]/); const m=match(p[0].trim()); const v=parseFloat(p[1]); if(!isNaN(v)) m.rr=v; rows.push(m); }); resolve(rows); }; rd.readAsText(f,'utf-8'); return; }
  if(name.endsWith('.pdf')){ extractPdfText(f).then(t=>resolve(parseReportText(t, document.getElementById('m-ist').value.trim()))).catch(err=>{ alert('PDF 解析失败：'+err.message); resolve([]); }); return; }
  if(name.endsWith('.docx')){ if(typeof mammoth==='undefined'){ alert('Word 解析库未加载（离线环境无法解析 .docx，请另存 CSV）'); resolve([]); return; }
    rd.onload=async e=>{ try{ const res=await mammoth.extractRawText({arrayBuffer:e.target.result}); resolve(parseReportText(res.value, document.getElementById('m-ist').value.trim())); }catch(err){ alert('Word 解析失败：'+err.message); resolve([]); } }; rd.readAsArrayBuffer(f); return; }
  if(typeof XLSX==='undefined'){ alert('Excel 解析库未加载'); resolve([]); return; }
  rd.onload=e=>{ try{ resolve(parseXlsxRows(e.target.result)||[]); }catch(err){ alert('Excel 解析失败：'+err.message); resolve([]); } }; rd.readAsArrayBuffer(f);
}); }
function renderMulti(samples){ _multiSamples=samples; const fmtOAV=v=>(v==null||isNaN(v))?'—':v.toPrecision(3).replace(/\.?0+$/,'');
  const map={}; samples.forEach(s=>{ s.rows.forEach(r=>{ const en=r.en; if(!en) return; if(!map[en]) map[en]={cn:r.cn,cat:r.cat,odor:r.odor,bySample:{}}; map[en].bySample[s.name]={conc:r.conc,rr:r.rr,oav:r.oav,roav:r.roav,flag:r.oav_flag}; }); });
  const ensAll=Object.keys(map); const catMap={}; ensAll.forEach(en=>{ const c=map[en].cat||'其他'; catMap[c]=(catMap[c]||0)+1; });
  const catOptions=Object.entries(catMap).sort((a,b)=>b[1]-a[1]);
  const catOrderKeys=Object.keys(CAT); const catRank=c=>{ const i=catOrderKeys.indexOf(c); return i>=0?i:999; };
  const bestOAV=en=>Object.values(map[en].bySample).reduce((m,x)=>Math.max(m,(x&&x.oav!=null&&!isNaN(x.oav))?x.oav:0),0);
  const sortEns=list=>list.slice().sort((a,b)=>{ const ca=map[a].cat||'其他',cb=map[b].cat||'其他'; if(ca!==cb) return catRank(ca)-catRank(cb); return bestOAV(b)-bestOAV(a); });
  const nS=samples.length;
  // 未填内标参数时 conc 为 null，回退展示响应比，避免整列空白
  const hasConc=samples.some(s=>s.rows.some(r=>r.conc!=null&&!isNaN(r.conc)));
  const valLabel=hasConc?'浓度(μg/L) / OAV':'响应比 / OAV';
  const ctrlSample = samples.find(s=>/ck|对照|control|空白/i.test(s.name)) || samples[nS-1];
  const head='<th class="multi-del-col">✕</th><th>物质</th><th>中文名</th><th>类别</th>'+samples.map(s=>'<th>'+escapeHtml(s.name)+'<br><small>'+valLabel+'</small></th>').join('')+'<th>相对于对照<br><small>'+(hasConc?'浓度比':'响应比')+'</small></th>'+'<th>气味活性</th><th>样品检出</th><th>风味描述</th>';
  function buildRows(filterCat){ const ens=filterCat?ensAll.filter(en=>(map[en].cat||'其他')===filterCat):ensAll; const sorted=sortEns(ens);
    const del=mcDelSet();
    if(!sorted.length) return '<tr><td colspan="'+(nS+8)+'" style="text-align:center;color:var(--mut)">该类别下无物质</td></tr>';
    return sorted.map(en=>{ if(del.has(en)) return ''; const m=map[en]; const cnt=Object.keys(m.bySample).length; let act='—',actCls='f-na';
      Object.values(m.bySample).forEach(x=>{ if(x&&x.flag==='关键致香'){act='关键致香';actCls='f-key';} }); if(act==='—') Object.values(m.bySample).forEach(x=>{ if(x&&x.flag==='潜在贡献'){act='潜在贡献';actCls='f-pot';} });
      const cells=samples.map(s=>{ const x=m.bySample[s.name]; if(!x) return '<td class="m-none">—</td>';
        const hasC=x.conc!=null&&!isNaN(x.conc); const hasR=x.rr!=null&&!isNaN(x.rr);
        if(!hasC&&!hasR) return '<td class="m-none">—</td>';
        const cu=Number(hasC?x.conc:x.rr).toPrecision(4).replace(/\.?0+$/,''); const tag=hasC?'':'(响应比) ';
        const fls=x.flag==='关键致香'?'f-key':(x.flag==='潜在贡献'?'f-pot':'f-na');
        return '<td class="oavflag '+fls+' num-dark">'+cu+'<br><small>'+tag+'OAV '+fmtOAV(x.oav)+'</small></td>'; }).join('');
      // 相对于对照：各样品（浓度或响应比，与展示口径一致）÷ 对照样品，与单 PDF「相对对照」并列设计一致
      let relCell='—';
      if(ctrlSample){ const ccx=m.bySample[ctrlSample.name];
        if(ccx){ const cBase=hasConc?(ccx.conc!=null&&!isNaN(ccx.conc)?ccx.conc:null):(ccx.rr!=null&&!isNaN(ccx.rr)?ccx.rr:null);
          if(cBase!=null&&cBase>0){ const parts=[];
            samples.forEach(s=>{ if(s===ctrlSample) return; const x=m.bySample[s.name]; if(!x) return;
              const v=hasConc?(x.conc!=null&&!isNaN(x.conc)?x.conc:null):(x.rr!=null&&!isNaN(x.rr)?x.rr:null);
              if(v!=null&&v>0){ const fc=v/cBase; const cls=fc>=2?'fc-up':(fc<=0.5?'fc-down':''); const arr=fc>=2?'▲':(fc<=0.5?'▼':'');
                parts.push('<span class="'+cls+'">'+escapeHtml(s.name)+' '+arr+fc.toFixed(1)+'×</span>'); } });
            relCell=parts.join('<br>')||'—'; } } }
      const od=(m.odor&&m.odor.trim())?m.odor:'—';
      const cnDisp=(m.cn && m.cn!=='(未收录)' && m.cn.trim()!=='')?m.cn:en;
      return '<tr><td class="multi-del-col"><button class="multi-del-btn" data-en="'+escapeHtml(en)+'" title="删除该物质">✕</button></td><td><b>'+en+'</b></td><td class="cn">'+cnDisp+'</td><td>'+badge(m.cat)+'</td>'+cells+'<td style="text-align:left">'+relCell+'</td><td class="oavflag '+actCls+'">'+act+'</td><td style="text-align:center">'+cnt+'/'+nS+'</td><td style="min-width:160px">'+od+'</td></tr>'; }).join(''); }
  let filterBar='<div class="multi-toolbar"><label>类别筛选：<select id="m-cat-filter" onchange="renderMultiMatrix()"><option value="">全部类别</option>';
  catOptions.forEach(([c,n])=>{ filterBar+='<option value="'+c+'">'+c+'（'+n+'）</option>'; }); filterBar+='</select></label><span class="hint">按类别顺序分组 · 同类数值由大到小</span><button class="ghost" onclick="exportMultiCSV()">⬇ 导出对比报告(CSV)</button></div>';
  document.getElementById('multi-out').innerHTML='<p class="hint">共解析 '+samples.length+' 个样品、'+ensAll.length+' 种风味物质参与对比（按类别分组、同类 OAV 由大到小排序，可点选类别筛选）。</p>'+filterBar+'<div class="scroll" id="multi-matrix"><table><thead><tr>'+head+'</tr></thead><tbody id="multi-matrix-rows">'+buildRows('')+'</tbody></table></div>' + '<div class="viz-wrap" id="viz-multi"></div>';
  window.__buildMultiRows=buildRows;
  const mm=document.getElementById('multi-matrix');
  if(mm && !mm._wired){ mm._wired=true; mm.addEventListener('click', function(e){ const b=e.target.closest('.multi-del-btn'); if(b) mcDeleteMulti(b.getAttribute('data-en')); }); }
  if(window.NatureViz) NatureViz.build('multi', _multiSamples);
}
function renderMultiMatrix(){ const sel=document.getElementById('m-cat-filter'); const f=sel?sel.value:''; const rows=window.__buildMultiRows?window.__buildMultiRows(f):''; const box=document.getElementById('multi-matrix-rows'); if(box) box.innerHTML=rows; }
function exportMultiCSV(){ const samples=_multiSamples||[]; if(!samples.length) return; const map={};
  samples.forEach(s=>s.rows.forEach(r=>{ if(!r.en) return; if(!map[r.en]) map[r.en]={cn:r.cn,cat:r.cat,odor:r.odor,bySample:{}}; map[r.en].bySample[s.name]={conc:r.conc,rr:r.rr,oav:r.oav,roav:r.roav,flag:r.oav_flag}; }));
  const hasConc=samples.some(s=>s.rows.some(r=>r.conc!=null&&!isNaN(r.conc)));
  const vLabel=hasConc?'_浓度(μg/L)':'_响应比';
  const ens=Object.keys(map); const ctrlName=(samples.find(s=>/ck|对照|control|空白/i.test(s.name))||samples[samples.length-1]);
  const head=['物质','中文名','类别'].concat(samples.map(s=>s.name+vLabel)).concat(samples.map(s=>s.name+'_OAV')).concat(samples.map(s=>s.name+'_ROAV')).concat(['相对于对照('+(ctrlName?ctrlName.name:'')+(hasConc?'_浓度比':'_响应比')+')','气味活性','样品检出数','风味描述']);
  const lines=[head.map(csvCell).join(',')]; ens.forEach(en=>{ const m=map[en]; const cnt=Object.keys(m.bySample).length; let act='—';
    Object.values(m.bySample).forEach(x=>{ if(x&&x.flag==='关键致香')act='关键致香'; }); if(act==='—') Object.values(m.bySample).forEach(x=>{ if(x&&x.flag==='潜在贡献')act='潜在贡献'; });
    const row=[en,m.cn||'',m.cat||'其他']; samples.forEach(s=>{ const x=m.bySample[s.name]; const hasC=x&&x.conc!=null&&!isNaN(x.conc); const hasR=x&&x.rr!=null&&!isNaN(x.rr);
      row.push(hasC?Number(x.conc).toPrecision(4).replace(/\.?0+$/,''):(hasR?Number(x.rr).toPrecision(4).replace(/\.?0+$/,''):'—')); });
    samples.forEach(s=>{ const x=m.bySample[s.name]; row.push((x&&x.oav!=null&&!isNaN(x.oav))?x.oav.toPrecision(3).replace(/\.?0+$/,'') : '—'); });
    samples.forEach(s=>{ const x=m.bySample[s.name]; row.push((x&&x.roav!=null&&!isNaN(x.roav))?x.roav.toFixed(2):'—'); });
    let relParts=[]; if(ctrlName){ const ccx=m.bySample[ctrlName.name];
      if(ccx){ const cBase=hasConc?(ccx.conc!=null&&!isNaN(ccx.conc)?ccx.conc:null):(ccx.rr!=null&&!isNaN(ccx.rr)?ccx.rr:null);
        if(cBase!=null&&cBase>0){ samples.forEach(s=>{ if(s===ctrlName) return; const x=m.bySample[s.name]; if(!x) return;
          const v=hasConc?(x.conc!=null&&!isNaN(x.conc)?x.conc:null):(x.rr!=null&&!isNaN(x.rr)?x.rr:null);
          if(v!=null&&v>0) relParts.push(s.name+'='+(v/cBase).toFixed(2)); }); } } }
    row.push(relParts.join(';')||'');
    row.push(act); row.push(cnt+'/'+samples.length); row.push((m.odor&&m.odor.trim())?m.odor:'—'); lines.push(row.map(csvCell).join(',')); });
  const blob=new Blob(['\ufeff'+lines.join('\r\n')],{type:'text/csv;charset=utf-8'}); const a=document.createElement('a'); a.href=URL.createObjectURL(blob); a.download='风味物质对比报告.csv'; a.click();
}
function csvCell(v){ v=(v==null)?'':String(v); if(/[",\r\n]/.test(v)) v='"'+v.replace(/"/g,'""')+'"'; return v; }

// ---------- ⑤ 浏览 ----------
let _allDB=null;
function populateBcat(){
  const sel=document.getElementById('bcat'); if(!sel||!_allDB) return;
  if(sel.dataset.filled) return;
  const cnt={}; _allDB.forEach(c=>{const k=c.cat||'其他'; cnt[k]=(cnt[k]||0)+1;});
  const opts=Object.entries(cnt).sort((a,b)=>b[1]-a[1]);
  sel.innerHTML='<option value="">全部类别</option>'+opts.map(([k,n])=>'<option value="'+k+'">'+k+'（'+n+'）</option>').join('');
  sel.dataset.filled='1';
}
async function browseDB(){ if(!_allDB) _allDB=DB_RAW; populateBcat(); const cat=document.getElementById('bcat').value; const key=normalize(document.getElementById('bkey').value);
  let rows=_allDB.filter(c=>(!cat||c.cat===cat)&&(!key||normalize(c.en).includes(key)||normalize(c.cn).includes(key)||(c.syn||[]).some(s=>normalize(s).includes(key))));
  rows=rows.map(c=>Object.assign({match:'库内',query:''},c));
  rows=sortByCompleteness(rows);  // 信息完整（有风味描述/阈值/来源）优先展示
  const nFull=rows.filter(r=>completeness(r)>=4).length;
  document.getElementById('bcount').textContent='共 '+rows.length+' 种（信息完整 '+nFull+' 种优先展示）';
  document.getElementById('browse-out').innerHTML='<div class="scroll">'+tableHTML(rows)+'</div>';
}

// ---------- tab ----------
document.querySelectorAll('.tab').forEach(t=>t.onclick=()=>{ document.querySelectorAll('.tab').forEach(x=>x.classList.remove('active')); document.querySelectorAll('.panel').forEach(x=>{x.classList.remove('active'); x.style.display='none';}); t.classList.add('active'); const p=document.getElementById('p-'+t.dataset.t); if(p){p.classList.add('active'); p.style.display='block';} if(t.dataset.t==='browse') browseDB(); });
</script>
<script>
/* ===== 处理组对比（单 PDF 多处理组自动识别）—— 渲染与对比逻辑（与在线版一致） ===== */
__GC_MULTI_JS__
</script>
<script>
/* ===== 处理组对比：内标参数变动实时重算 ===== */
(function(){
  const ids=['r-cis','r-vis','r-ms','r-cisu','r-visu','r-msu','r-ist'];
  ids.forEach(id=>{ const el=document.getElementById(id); if(el) el.addEventListener('input', ()=>{ if(lastMulti && lastMulti.compounds) recalcGroupCompare(); else if(lastReport && lastReport.length){ applyReportConc(); renderReport(); } }); });
})();
</script>
<script>
/* ===== Nature 风格可视化（热力图 / PCA）+ 物质管理 ===== */
__NATURE_VIZ_JS__
</script>
</body>
</html>
"""

# 用在线版 app.py 的基础 <style> 与 tabs+面板骨架替换离线版自带的分叉副本（单一来源，UI 保持一致）
if BASE_STYLE and BASE_SKELETON:
    HTML = re.sub(r'<style>.*?__GC_MULTI_CSS__.*?</style>',
                  '<style>' + BASE_STYLE + '</style>', HTML, count=1, flags=re.S)
    HTML = re.sub(r'<div class="tabs">.*?</div>\s*</div>\s*<!-- ===== 内联解析库',
                  BASE_SKELETON + '\n</div>\n<!-- ===== 内联解析库', HTML, count=1, flags=re.S)
    # 离线表格容器沿用 .scroll（与在线 .wrap 视觉一致）：补一条等价样式（app.py 仅定义 .wrap）
    HTML = HTML.replace('</style>',
        '.scroll{max-height:66vh;overflow:auto;border:1px solid var(--line);border-radius:var(--radius);background:#fff;box-shadow:var(--shadow-sm)}\n'
        '.scroll thead th:first-child{border-top-left-radius:var(--radius)}\n'
        '.scroll thead th:last-child{border-top-right-radius:var(--radius)}\n'
        '.scroll::-webkit-scrollbar{height:10px;width:10px}\n'
        '.scroll::-webkit-scrollbar-thumb{background:#d8dee6;border-radius:8px;border:2px solid #fff}\n'
        '.scroll::-webkit-scrollbar-thumb:hover{background:#c2ccd6}\n'
        '</style>', 1)

html = (HTML
        .replace("__DATA__", data)
        .replace("__CAT_COLOR__", json.dumps(CAT_COLOR, ensure_ascii=False))
        .replace("__N__", str(COMPOUNDS_N))
        .replace("__USE_XLSX__", USE_XLSX)
        .replace("__USE_PDF__", USE_PDF)
        .replace("__USE_PDF_WORKER__", USE_PDF_WORKER)
        .replace("__USE_CMAPS__", CMAP_JS)
        .replace("__USE_MAMMOTH__", USE_MAMMOTH)
        .replace("__GC_MULTI_JS__", GC_MULTI_JS)
        .replace("__GC_MULTI_CSS__", GC_MULTI_CSS)
        .replace("__GC_MULTI_HTML__", GC_MULTI_HTML)
        .replace("__NATURE_VIZ_JS__", NATURE_VIZ_JS))

out = os.path.join(HERE, "风味物质检索_离线版.html")
with open(out, "w", encoding="utf-8") as f:
    f.write(html)
print("生成:", out, "数据库条目:", COMPOUNDS_N, "HTML 大小:", round(len(html)/1024/1024,2), "MB",
      "| SheetJS内联:", bool(XLSX_JS), "| pdf.js内联:", bool(PDF_JS))
