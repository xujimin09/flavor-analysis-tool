# 风味物质检索与分析工具（Flavor Compound Analysis Tool）

一个面向 GC-MS 风味分析的中文工具：上传报告/文献即可自动识别风味物质，按**内标法**计算浓度与 OAV/ROAV，支持处理组对比、关键致香物分析与风味热力图可视化。提供 **Flask Web 服务**与**完全离线单文件 HTML**两种用法。

## 功能特性

| 模块 | 说明 |
|------|------|
| 单物质检索 | 输入中/英文名，返回类别、阈值、气味描述、来源 |
| 批量检索 | 粘贴/上传物质名单，批量匹配并聚合 |
| GC-MS 报告解析 | 上传 `.xlsx / .pdf / .docx / .txt`，自动识别化合物、内标、响应比，按内标法算浓度/OAV/ROAV |
| 处理组对比 | 单个 PDF 自动识别多个处理组，横向对比共有/特有物质与关键致香物 |
| 多报告批量对比 | 逐个上传多份报告，统一内标法计算并对比 |
| 风味热力图 | 物质 × 处理组归一化热力图，多套配色，支持导出 PNG / SVG |

- 内置 ~1900 种风味化合物数据库（中英文名、CAS、阈值、气味描述、来源、类别）。
- 离线版 HTML 内联数据库与解析库（pdf.js / SheetJS / mammoth），**无需联网、无需服务器**，双击即用。

## 目录结构

```
flavor_tool/
├── app.py                 # Flask 服务（Web 主程序）
├── flavor_core.py         # 核心：物质匹配、报告解析、数据库加载
├── build_standalone.py    # 把数据库+前端打包成离线单文件 HTML
├── 风味物质检索_离线版.html  # 由 build_standalone.py 生成（见下方命令）
├── requirements.txt
├── SKILL.md               # CodeBuddy Skill 描述（见下文）
├── vendor/                # 离线 HTML 内联的第三方解析库（必须保留）
│   ├── xlsx.full.min.js
│   ├── pdf.min.js / pdf.worker.min.js
│   ├── mammoth.browser.min.js
│   └── cmaps/             # pdf.js 中文字体映射
└── *.json / *.csv         # 化合物数据库与衍生数据（运行时加载）
```

## 环境要求

- Python 3.10+
- 依赖：`flask`、`openpyxl`、`pdfplumber`

## 快速开始（Web 服务）

```bash
pip install -r requirements.txt
python app.py
# 打开 http://localhost:8080
```

## 生成离线单文件 HTML

```bash
python build_standalone.py
# 生成 风味物质检索_离线版.html（自包含，可直接发给他人/双击打开）
```

> 构建依赖 `vendor/` 目录与数据库文件，请勿删除。

## 作为 CodeBuddy Skill 使用

本仓库根目录的 `SKILL.md` 即一个 CodeBuddy Skill。安装方式：

```bash
# 将本仓库作为 skill 放到 CodeBuddy 的 skills 目录
cp -r . ~/.codebuddy/skills/flavor-analysis
# 或直接 symlink
ln -s "$(pwd)" ~/.codebuddy/skills/flavor-analysis
```

之后在对话中提到「风味物质 / GC-MS 报告 / 内标法 / 热力图」等需求时，Agent 会自动加载该 Skill 并按 `SKILL.md` 指导操作。

## HTTP API（供二次开发）

| 端点 | 说明 |
|------|------|
| `POST /api/search` | `{"names":[...]}` 批量匹配物质 |
| `POST /api/batch_upload` | 上传名单文件（txt/csv/xlsx） |
| `POST /api/upload` | 上传单份报告（xlsx/pdf/docx/txt），返回解析结果或 `{multi:true}` 处理组对比数据 |
| `POST /api/multi_upload` | 上传多份报告做批量对比 |
| `GET  /api/all` | 返回完整化合物库 |
| `GET  /api/info` | 返回库规模等元信息 |

## License

本项目仅供科研与教学使用。数据库整理自公开文献，版权归原作者所有。
