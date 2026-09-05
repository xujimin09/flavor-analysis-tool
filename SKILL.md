---
name: flavor-analysis
description: 风味物质 GC-MS 检索与分析工具。当用户需要解析 GC-MS 定量报告或文献（xlsx/pdf/docx/txt）、按内标法计算浓度与 OAV/ROAV、做处理组对比、关键致香物分析、风味热力图可视化时使用。支持单物质检索、批量检索、报告解析、处理组对比（单 PDF 多组自动识别）、多报告批量对比、完全离线单文件 HTML。触发词：风味物质、GC-MS、内标法、OAV、ROAV、响应比、处理组、致香物、热力图、香气成分。
---

# 风味物质分析工具（Flavor Analysis Tool）

本 Skill 封装了一套中文风味物质分析工具，核心能力是把 GC-MS 报告/文献自动变成可解释的风味结论。

## 何时使用

- 用户上传/提及 GC-MS 报告、风味物质清单、香气成分表。
- 需要计算浓度、OAV/ROAV、响应比 rr、相对对照倍数。
- 需要「处理组对比」「关键致香物」「热力图」等可视化。
- 需要完全离线、可单文件分发的分析页面。

## 仓库关键文件

| 文件 | 作用 |
|------|------|
| `app.py` | Flask 主程序，提供 Web 界面与 `/api/*` 接口 |
| `flavor_core.py` | 核心：物质匹配 `match_compound`、报告解析 `parse_gcms_excel`、富化 `enrich`、自动分类 `auto_classify`；启动时加载数据库 |
| `build_standalone.py` | 把数据库与前端打包成自包含单文件 HTML |
| `vendor/` | 离线 HTML 内联的解析库（pdf.js / SheetJS / mammoth / cmaps），**构建离线版时必须保留** |
| `*.json` / `*.csv` | 化合物数据库与衍生数据，运行时由 `flavor_core` 加载 |

> 运行 Flask 只需 `app.py + flavor_core.py + 数据文件 + vendor/`；构建离线版还需要 `build_standalone.py`。

## 常用操作

### 1. 启动 Web 服务
```bash
pip install -r requirements.txt
python app.py          # 监听 http://localhost:8080
```

### 2. 生成离线单文件 HTML（无需服务器、可分发）
```bash
python build_standalone.py
# 产出 风味物质检索_离线版.html（自包含，双击即用）
```
注意：`build_standalone.py` 会从 `vendor/` 内联第三方解析库，并内联整个化合物数据库，因此产物较大（数 MB）。

### 3. 编程调用（HTTP API）
- `POST /api/search` → `{"names":["Butanoic acid","丁酸"]}` 批量匹配
- `POST /api/upload` → multipart `file` + `ist_name`，单份报告解析（返回 `rows` 或 `{multi:true, groups, compounds,...}`）
- `POST /api/multi_upload` → 多份报告批量对比
- `GET /api/all` → 完整化合物库（前端检索用）

## 领域要点（内标法）

- 响应比 `rr = A ÷ A₁`（A=化合物峰面积，A₁=内标峰面积），已含内标信息，浓度公式不再单独除以 A₁。
- 浓度 `C = 内标浓度 × 内标加量 × rr ÷ 样品量`，结果折算到 μg/L。
- `OAV = C ÷ 阈值`（>1 具气味活性）；`ROAV = OAV ÷ 最大OAV × 100`（≥10 关键致香，≥1 潜在贡献）。
- 处理组对比自动识别「对照/CK/control/空白」为对照组，其余组相对其算上调(≥2×)/下调(≤0.5×)。

## 注意事项

- 修改前端/配色逻辑时，注意 `build_standalone.py` 会从 `app.py` 自动抽取 `GC_MULTI_JS` / `NATURE_VIZ_JS` 片段同步到离线版，改完务必重新运行 `build_standalone.py`。
- 配色方案集中在 `NV_PALETTES`（`app.py` 内 `NATURE_VIZ_JS` 区块），增删配色后离线版自动同步。
- 不要删除 `vendor/` 或数据文件，否则离线构建与运行时加载都会失败。
