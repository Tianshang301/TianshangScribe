# TianshangScribe

> [English](../README.md)

<p align="center">
  <a href="https://glama.ai/mcp/servers/Tianshang301/TianshangScribe">
    <img alt="TianshangScribe MCP server"
         src="https://glama.ai/mcp/servers/Tianshang301/TianshangScribe/badges/card.svg"
         width="380">
  </a>
</p>

[![PyPI](https://img.shields.io/badge/pypi-tianshang--scribe-blue)](https://pypi.org/project/tianshang-scribe/)
[![CI](https://github.com/Tianshang301/TianshangScribe/actions/workflows/ci.yml/badge.svg)](https://github.com/Tianshang301/TianshangScribe/actions)
[![License](https://img.shields.io/badge/license-Apache%202.0-green)](LICENSE)
[![TianshangScribe MCP server](https://glama.ai/mcp/servers/Tianshang301/TianshangScribe/badges/score.svg)](https://glama.ai/mcp/servers/Tianshang301/TianshangScribe)

跨平台命令行 Office 文档处理工具。支持 Word、Excel、PowerPoint 的创建、编辑、模板填充、格式转换，内置 LaTeX 风格排版标记、数学公式渲染引擎与 MCP Server（AI Agent 集成）。


> **警告：接口不稳定——预期将发生破坏性更新**
>
> 本项目处于 0.x 预发布阶段。CLI 选项、MCP 工具签名、模板语法与输出格式**尚未冻结**，
> 可能随时发生不兼容变更。
> **兼容期承诺**：任何破坏性变更都会至少提前一个版本在 CHANGELOG 中预告，并附迁移指南。
> 生产环境请锁定具体版本，升级前先查阅 CHANGELOG。

## 安装

```bash
pip install tianshang-scribe

# 或从源码安装：
git clone https://github.com/Tianshang301/TianshangScribe.git
cd TianshangScribe
pip install -e ".[dev]"
```

### Linux 部署

**Docker**（推荐，用于 Streamable HTTP MCP Server）：
```bash
git clone https://github.com/Tianshang301/TianshangScribe.git
cd TianshangScribe
docker compose up -d
# Streamable HTTP MCP Server 监听 http://localhost:8080/mcp
# （传输 / 认证 / 限流均可通过 TIANSHANG_SCRIBE_* 环境变量覆盖）
```

**.deb 包**（Debian / Ubuntu）：
```bash
# 从 GitHub Releases 下载
sudo dpkg -i tianshang-scribe_0.7.1_all.deb
tianshang-scribe --help
```

**pipx**（隔离 CLI 安装）：
```bash
pipx install tianshang-scribe
tianshang-scribe --help
```

依赖：Python 3.10+ · python-docx · openpyxl · python-pptx · typer · rich · lxml

## 快速开始

```bash
# 创建 Word 文档
tianshang-scribe -w --create -a "Hello World" -o hello.docx

# 替换文本（--regex 启用正则）
tianshang-scribe input.docx -r "旧文本" --replace-new "新文本" -o output.docx

# LaTeX 标记 + 嵌套
tianshang-scribe -w --create --latex-style \
  -s "font=Times New Roman,size=14" \
  -a "\bfseries{\itshape{加粗斜体}} \fontsize{24}{标题} \color{FF0000}{红色}" \
  -o styled.docx

# 数学公式 —— 自动转为 Word 原生 OMML
tianshang-scribe -w --create \
  --math "x = \frac{-b \pm \sqrt{b^2 - 4ac}}{2a}" \
  --math "\sum_{i=0}^{n} i^2" \
  -o formulas.docx

# 模板填充（JSON / CSV / YAML → {{placeholder}}）
tianshang-scribe template.docx -t data.json -o filled.docx

# 转换为 PDF（office2pdf ~2MB，或 LibreOffice 回退）
tianshang-scribe input.docx --topdf -o output.pdf

# MCP Server — stdio 模式（Claude Code / Cursor）
python -m tianshang_scribe.mcp.server

# MCP Server — SSE 模式（Dify / Coze / FastGPT）
python -m tianshang_scribe.mcp.server --transport sse --port 8080

# Excel：导入 CSV、排序、导出 JSON
tianshang-scribe -e --create --from-csv data.csv --sort "A1:A10 asc" --to-json -o out.json

# Excel：写入公式、设置密码
tianshang-scribe budget.xlsx --formula "B10 =SUM(B2:B9)" --protect "p@ss" -o protected.xlsx
```

## 全局选项

| 参数 | 说明 |
|------|------|
| `input_file` | 输入文档路径（与 `--create` 互斥） |
| `-w` `--word` | 处理 Word 文档 |
| `-e` `--excel` | 处理 Excel 工作簿 |
| `-p` `--ppt` | 处理 PowerPoint 演示文稿 |
| `-o` `--output` | 输出文件路径 |
| `--force` | 允许覆盖已有文件 |
| `--topdf` | 输出为 PDF |
| `--stdin` | 从标准输入读取 |
| `--stdout` | 输出到标准输出 |

不指定 `-w/-e/-p` 时，根据输入文件扩展名自动推断类型。

## 通用操作

| 参数 | 说明 | 示例 |
|------|------|------|
| `-cr` `--create` | 创建空白文档 | `--create -w` |
| `-a` `--add` | 添加文本 | `-a "Hello"` |
| `--column` | 指定 `--add` 目标列 | `--column 2` |
| `-r` `--replace` | 查找替换 | `-r "foo" --replace-new "bar"` |
| `-d` `--delete` | 删除内容 | `-d "关键词"` |
| `-cl` `--clear` | 清除内容/格式/链接 | `--clear formats` |
| `-m` `--modify` | 修改内容 | `-m "旧值" --modify-new "新值"` |
| `-s` `--style` | 设置样式 | `-s "font=Times,size=14,bold"` |
| `-t` `--template` | 模板填充 | `-t data.json` |
| `-x` `--extract` | 提取数据（`math`/`latex` 等） | `-x latex` |
| `--meta` | 设置属性 | `--meta "title=报告,author=张三"` |
| `--latex-style` | 启用 LaTeX 标记解析 | |
| `--math` | 添加数学公式（Word） | `--math "\frac{a}{b}"` |
| `--math-style` | 公式解析方言（office/mathtype） | `--math-style mathtype` |
| `--math-font` | OMML 公式渲染字体（默认 Cambria Math） | `--math-font "Times New Roman"` |
| `--math-mtef` | 以 MathType OLE 对象（MTEF）嵌入 | `--math "\frac{a}{b}" --math-mtef` |
| `--heading` | 添加标题（Word） | `--heading "level:1 text:引言"` |
| `--regex` | 正则模式 | 配合 `--replace` `--delete` 使用 |
| `--merge` | 合并文件 | `--merge "a.docx,b.docx"` |
| `--split` | 拆分文档（仅 Excel：`--split by-sheet`） | `--split by-sheet` |
| `--comment` | 添加批注（Word）/演讲者备注（PPT） | `--comment "2 批注文字"` |
| `--add-table` | 添加表格（Word） | `--add-table "H1,H2\|a1,a2"` |
| `--chart-add` | 创建图表（Excel） | `--chart-add "type=bar data=B1:C10"` |
| `--batch` | 批量模式 | `--batch` |
| `--files` | 批量 glob 通配符 | `--files "reports/*.docx"` |
| `--schedule-db` | 调度 SQLite 数据库路径 | `--schedule-db ~/.tianshang-scribe/schedules.db` |
| `--schedule-add` | 注册调度 | `--schedule-add "daily\|0 9 * * *\|echo hi"` |
| `--schedule-rm` | 删除调度 | `--schedule-rm daily` |
| `--schedule-list` | 列出调度 | `--schedule-list` |
| `--schedule-run` | 立即运行调度 | `--schedule-run daily` |
| `--schedule-run-all` | 运行所有到期调度 | `--schedule-run-all` |
| `--run-script` | 沙箱执行脚本 | `--run-script build.py` |
| `--stdin` | 从标准输入读取 | |
| `--stdout` | 输出到标准输出 | |

## Word 专属操作

| 参数 | 说明 | 示例 |
|------|------|------|
| `--heading` | 添加标题 | `--heading "level:1 text:引言"` |
| `--math` | 数学公式 | `--math "\frac{a}{b}"` |
| `--latex-style` | LaTeX 标记 | |
| `--toc` | 生成目录 | `--toc` |
| `--section-break` | 插入分节符 | `--section-break` |
| `--header` | 设置页眉 | `--header "第一章"` |
| `--footer` | 设置页脚 | `--footer "第 X 页"` |
| `--watermark` | 文字水印 | `--watermark "草稿"` |
| `--tomd` | 转为 Markdown | `--tomd` |
| `--tohtml` | 转为 HTML | `--tohtml` |

## Excel 专属操作

| 参数 | 说明 | 示例 |
|------|------|------|
| `--sheet-add` | 添加工作表 | `--sheet-add "Q1"` |
| `--sheet-delete` | 删除工作表 | `--sheet-delete "Sheet2"` |
| `--sheet-rename` | 重命名工作表 | `--sheet-rename "旧名 新名"` |
| `--column-width` | 设置列宽 | `--column-width "2=20"` |
| `--row-height` | 设置行高 | `--row-height "3=30"` |
| `--formula` | 设置公式 | `--formula "A1 =SUM(B1:B10)"` |
| `--from-csv` | 导入 CSV | `--from-csv data.csv` |
| `--sort` | 排序 | `--sort "A1:A10 asc"` |
| `--chart-add` | 添加图表 | `--chart-add "type=bar data=B1:C10"` |
| `--protect` | 设置密码 | `--protect "p@ss"` |
| `--unprotect` | 解除密码 | `--unprotect` |
| `--to-csv` | 导出 CSV | |
| `--to-json` | 导出 JSON | |
| `--to-html` | 导出 HTML | |

## LaTeX 样式标记

在 `--add` 中嵌入下列标记，启用 `--latex-style` 后自动解析。支持嵌套。

| 标记 | 效果 |
|------|------|
| `\bfseries{text}` | 加粗 |
| `\itshape{text}` | 斜体 |
| `\scshape{text}` | 小型大写 |
| `\underline{text}` | 下划线 |
| `\rmfamily{text}` | 衬线（Roman） |
| `\sffamily{text}` | 无衬线（Sans-serif） |
| `\ttfamily{text}` | 等宽（Monospace） |
| `\fontfamily{Arial}{text}` | 指定字体 |
| `\fontsize{18}{text}` | 字号（pt） |
| `\color{FF0000}{text}` | 颜色（十六进制） |
| `\centering{...}` | 居中 **†** |
| `\raggedright{...}` | 左对齐 **†** |
| `\raggedleft{...}` | 右对齐 **†** |
| `\linespread{1.5}{...}` | 行距 **†** |
| `\indent{...}` / `\noindent{...}` | 缩进 **†** |
| `\heading{2}{标题}` | 插入标题 |
| `\newpage` | 分页符 |
| `\includegraphics{path}` | 插入图片 |

**†** 段落级格式（创建新段落）。

### 字体配置

| 命令 | 效果 |
|------|------|
| `\setmainfont{Name}` | 西文默认字体 |
| `\setCJKmainfont{Name}` | CJK 默认字体 |
| `\setsansfont{Name}` | 无衬线字体 |
| `\setCJKsansfont{Name}` | CJK 无衬线字体 |
| `\setmonofont{Name}` | 等宽字体 |
| `\setCJKmonofont{Name}` | CJK 等宽字体 |

Word OOXML 原生分离 `w:ascii`（西文）与 `w:eastAsia`（CJK）字体，中西文混排自动切换。

## 数学公式

通过 `--math` 支持完整的 LaTeX 数学公式，内部转换为 Word 原生 OMML。转换器为自研递归下降解析器（表达式→项→因子→原子），基于不可变嵌套 Token 树（分式/根号/大运算符/上下标/重音/样式/定界符 Token），经 O(1) 命令分发表、预编译正则与零拷贝参数切片驱动。使用 `--math-font "Times New Roman"` 可让公式以 MathType 风格衬线字体渲染，替代 Word 默认的 Cambria Math（通过 `<m:mathPr><m:mathFont>`）。`--math-style mathtype` 切换到 MathType 兼容的 LaTeX 解析方言。使用 `--math-mtef` 可将公式以真正的 MathType OLE 对象（MTEF 二进制）嵌入——供老版 MathType（6.x 及更早）编辑，与 `--extract math` 读取的是同一格式。输出跨版本逐字节稳定（由黄金快照回归套件保障）。

### 公式语法

| 类别 | 命令 |
|------|------|
| 分式 | `\frac{分子}{分母}` |
| 根号 | `\sqrt{内容}` `\sqrt[n]{内容}` |
| 上下标 | `x^{2}` `x_{i}` `x_{i}^{n}` |
| 求和/积分 | `\sum` `\int` `\oint` `\prod` `\coprod` `\bigcup` `\bigcap` `\bigvee` `\bigwedge` |
| 极限 | `\lim_{x \to 0}` `\max` `\min` `\sup` `\inf` |
| 标准函数 | `\sin` `\cos` `\tan` `\cot` `\sec` `\csc` `\log` `\ln` `\det` `\Pr` `\gcd` `\deg` `\dim` `\hom` `\ker` `\arg` |
| 希腊字母 | `\alpha` `\beta` `\gamma` … `\Gamma` `\Delta` `\Theta` … |
| 符号 | `\pm` `\times` `\div` `\cdot` `\infty` `\partial` `\nabla` `\forall` `\exists` … |
| 关系符 | `\leq` `\geq` `\neq` `\approx` `\equiv` `\propto` `\subset` `\supset` `\in` … |
| 箭头 | `\to` `\rightarrow` `\leftarrow` `\mapsto` `\uparrow` … |
| 重音 | `\hat{x}` `\bar{x}` `\tilde{x}` `\dot{x}` `\ddot{x}` `\vec{x}` `\widehat{x}` `\widetilde{x}` … |
| 括号 | `\left( \right)` `\left[ \right]` `\left\{ \right\}` |
| 数学字体 | `\mathrm{abc}` `\mathbf{abc}` `\mathit{abc}` `\mathcal{ABC}` `\mathbb{ABC}` `\mathsf{abc}` `\mathtt{abc}` |

### 数学字体规范

遵循主流数学期刊（AMS、Elsevier、Springer）标准：

| 内容 | 样式 | 示例 |
|------|------|------|
| 单字母变量 | *斜体* | `a` `b` `x` `y` |
| 数字 | **正体** | `0` `1` `2` … |
| 标准函数 | **正体** | `\sin` `\cos` `\log` |
| 希腊小写 | *斜体* | `\alpha` `\beta` `\gamma` |
| 希腊大写 | **正体** | `\Gamma` `\Delta` `\Theta` |

### 数学命令自动识别

`--add` 文本中的以下命令即使未被 `$...$` 包裹也会自动转为数学公式：

- 带参数：`\frac` `\sqrt` `\sum` `\int` `\prod` `\lim`
- 重音：`\hat{x}` `\bar{x}` `\vec{x}` 等
- 无参运算：`\sin` `\cos` `\tan` `\log` `\ln` 等
- `H_{2}O` 和 `m^{2}` 自动转为 Unicode 上下标（H₂O / m²）

## 样式语法

`--style` 采用逗号分隔的键值对：

```bash
--style "font=Times New Roman,size=14,bold,italic,color=FF0000,align=center"
```

| 键 | 别名 | 值 | 说明 |
|----|------|-----|------|
| `font` | `font_name`、`font-family` | 字体名 | 西文字体 |
| `cjk-font` | `cjk_font_name`、`cjk-font-family` | 字体名 | CJK 字体 |
| `size` | `font_size`、`font-size` | pt 值 | 字号 |
| `bold` | | 标志 | 加粗 |
| `italic` | | 标志 | 斜体 |
| `underline` | | 标志 | 下划线 |
| `color` | `font_color`、`font-color` | `FF0000` | 十六进制颜色 |
| `align` | `alignment` | `left`/`center`/`right`/`justify` | 对齐 |

布尔键（`bold` `italic` `underline`）出现即为 `True`。

## 模板填充

支持 JSON、CSV、YAML 数据源，填充 `{{placeholder}}` 占位符。嵌套对象以点号展开。列表值支持循环展开。条件块支持显隐控制。

```json
{
  "name": "张三",
  "date": "2026-07-28",
  "user": { "city": "北京" },
  "show": true,
  "paid": false,
  "items": [
    { "product": "部件", "price": "10" },
    { "product": "配件", "price": "20" }
  ]
}
```

```
{{name}}              →  张三
{{user.city}}         →  北京
{{#each items}}       →  循环展开每个元素
  {{product}}: {{price}}
{{/each}}
{{#if show}}          →  show 为真时显示
  机密内容
{{/if}}
{{#if role=admin}}    →  role 等于 "admin" 时显示
  管理仪表盘
{{/if}}
{{#unless paid}}      →  paid 为假时显示
  需付款
{{/unless}}
```

## Excel 功能

| 功能 | CLI 选项 |
|------|---------|
| 工作表管理 | `--sheet-add` `--sheet-delete` `--sheet-rename` |
| 行列尺寸 | `--column-width` `--row-height` |
| 公式 | `--formula "A1 =SUM(B1:B10)"` |
| 数据导入 | `--from-csv` |
| 数据导出 | `--to-csv` `--to-json` `--to-html` |
| 排序 | `--sort "A1:A10 asc"` |
| 图表 | `--chart-add "type=bar data=B1:C10"` |
| 保护 | `--protect` `--unprotect` |

## PPT 功能

| 功能 | 说明 |
|------|------|
| 幻灯片管理 | 增删移（`--slide-add`、`--slide-delete`、`--slide-move`） |
| 版式 | 按名称或索引应用版式（`--layout`） |
| 演讲者备注 | 添加备注（`--notes`） |
| 数学公式 | `$...$` / `$$...$$` 转为原生 OMML |
| 切换效果 | fade、push、wipe 等 17 种（`--transition`） |
| 导出 | 图片序列（`--toimg`）、PDF（`--topdf`） |
| 媒体压缩 | 压缩图片（`--compress-media "1920,80"`） |
| 保护 | 设置/解除密码（`--protect`、`--unprotect`） |

## 退出码

| 码 | 含义 |
|----|------|
| `0` | 成功 |
| `1` | 一般错误 |
| `2` | 参数错误 |
| `3` | 功能未实现 |

## MCP Server

TianshangScribe 内置 MCP（Model Context Protocol）服务端——AI Agent 可创建、编辑、填充模板、格式转换和提取 Office 文档数据。

### 快速接入

**stdio**（Claude Code、Cursor）：
```json
{"mcpServers": {"tianshang-scribe": {
  "command": "python", "args": ["-m", "tianshang_scribe.mcp.server"]
}}}
```

**SSE**（Dify、Coze、FastGPT）：
```bash
python -m tianshang_scribe.mcp.server --transport sse --host 0.0.0.0 --port 8080
```
```json
{"mcpServers": {"tianshang-scribe": {
  "url": "http://localhost:8080/sse", "transport": "sse"
}}}
```

### 工具列表（7 个）

| 工具 | 说明 |
|------|------|
| `create_office_document` | 用结构化内容块创建 .docx / .xlsx / .pptx |
| `edit_office_document` | 替换、删除、修改、样式、添加等操作 |
| `fill_template` | 用数据填充 `{{占位符}}`，支持 `{{#each}}` / `{{#if}}` |
| `convert_document` | 格式转换（docx↔pdf/md/html、xlsx↔csv/json） |
| `extract_document_data` | 提取元数据、全文或文档结构 |
| `validate_template` | 填充前预检查模板占位符与数据是否匹配 |
| `compare_documents` | 两个 .docx 文件的段落级差异对比 |

### 能力矩阵

| 特性 | 详情 |
|------|------|
| **协议** | MCP 2024-11-05 · stdio + SSE · JSON-RPC 2.0 |
| **资源** | `resources/list` + `resources/read` — 文档暴露为可读 URI |
| **提示模板** | 5 个内置工作流（`prompts/list` + `prompts/get`） |
| **进度通知** | PDF 转换等长操作发送 `notifications/progress` |
| **返回值** | 多类型 `content[]`：文本 + 资源（文件 URI、MIME 类型、大小） |
| **Schema** | 所有参数支持 `enum`、`default`、`examples`、约束 |

### 生产就绪（SSE 模式）

```bash
# 带认证启动
TIANSHANG_SCRIBE_AUTH_TOKEN="secret" python -m tianshang_scribe.mcp.server --transport sse --host 0.0.0.0 --port 8080

# 健康检查
curl http://localhost:8080/health
# {"status":"ok","version":"0.7.1","active_sessions":3,"tools_available":7}

# CORS 白名单
python -m tianshang_scribe.mcp.server --transport sse --cors-origins "https://coze.com,https://dify.ai"
```

**端点**：`GET /health` · `GET /sse` · `POST /message?session_id=X`

详细文档：[docs/mcp/README.zh-CN.md](docs/mcp/README.zh-CN.md)。

```bash
python tests/integration/mcp/mcp_stdio_smoke.py     # 9/9 基础测试（stdio）
python tests/integration/mcp/test_sse.py        # 3/3 SSE 传输测试
python tests/integration/mcp/mcp_agent_sim.py      # 11 场景 Agent 模拟测试
```

## 架构

```
src/
└── tianshang_scribe/    # 可 import 包（tianshang_scribe.*）
    ├── cli/               # Typer CLI 入口
    │   ├── main.py        # 命令解析与分发
    │   └── global_opts.py # 文件路径 / 类型推断
    ├── core/              # 文档引擎抽象层
    │   ├── document.py    # DocumentABC 统一接口
    │   ├── word_engine.py # Word 引擎 (python-docx)
    │   ├── excel_engine.py# Excel 引擎 (openpyxl)
    │   └── ppt_engine.py  # PPT 引擎 (python-pptx)
    ├── rendering/         # 样式与公式渲染
    │   ├── styles.py      # TextStyle 数据类
    │   ├── latex_parser.py # LaTeX 标记解析器
    │   ├── math_omml.py   # LaTeX → OMML 数学公式转换
    │   └── template.py    # 模板填充引擎
    ├── transform/         # 格式转换
    │   └── pdf.py         # PDF 导出（office2pdf + LibreOffice）
    ├── mcp/                    # MCP Server（官方 mcp SDK 2.x）
    │   ├── server.py           # build_server + 入口（stdio / SSE / Streamable HTTP）
    │   ├── transport.py        # 传输接线 + ASGI 中间件
    │   ├── schemas.py          # pydantic 模型 + as_dict
    │   ├── auth.py             # Bearer Token 认证
    │   ├── rate_limit.py       # 令牌桶限流
    │   ├── metrics.py          # Prometheus 风格指标
    │   ├── security.py         # 工具只读/破坏性分类
    │   ├── prompts.py          # 5 个提示词工作流
    │   ├── tools/              # 7 个 Agent 工具
    │   │   ├── _registry.py    # 工具注册表（schema 自动派生）
    │   │   ├── create.py / edit.py / template.py / convert.py
    │   │   ├── validate.py / compare.py
    │   └── errors.py           # 结构化错误码 + 修复建议
    └── utils/             # 工具函数
        └── file_utils.py
```

### 技术栈

| 组件 | 核心技术 |
|------|---------|
| CLI | Typer + Rich |
| Word | python-docx |
| Excel | openpyxl |
| PPT | python-pptx |
| 数学公式 | 自研递归下降解析器 → OMML XML（不可变 Token 树 + 命令分发表） |
| 模板 | 自研引擎（{{placeholder}}、{{#each}}、{{#if}}） |
| PDF | office2pdf（~2MB Rust 二进制，零依赖）+ LibreOffice 回退 |
| 质量 | pytest（936 用例）· ruff · mypy |

## 构建 EXE

```bash
pip install pyinstaller
pyinstaller --onefile --name tianshang-scribe --hidden-import openpyxl.cell._writer --hidden-import openpyxl.cell.read_only --hidden-import openpyxl.styles --hidden-import openpyxl.chart --hidden-import openpyxl.comments src/tianshang_scribe/cli/main.py
# dist/tianshang-scribe.exe (~35 MB)
```

## 演示

```bash
python -m demo.generate_demos
# demo/demo_word.docx   — LaTeX + 数学公式 + 目录 + 水印
# demo/demo_excel.xlsx  — CSV 导入 + 公式 + 图表 + 保护
# demo/demo_ppt.pptx    — 幻灯片 + 备注 + 切换效果 + 数学公式
```

CLI 合规测试：

```bash
python demo/test_cli.py
```

## 开发

```bash
git clone https://github.com/Tianshang301/TianshangScribe.git
cd TianshangScribe
pip install -e ".[dev]"

pytest tests/ -v      # 运行测试
ruff check src/tianshang_scribe/ tests/ # 代码检查
mypy src/tianshang_scribe/             # 类型检查
```

## 许可

Apache-2.0
