## AGENTS.md

> **项目名称**：天殇·书契（TianshangScribe）  
> **定位**：跨平台命令行 Office 文档处理工具，支持 Word、Excel、PowerPoint 的创建、编辑、模板填充、格式转换，融入 LaTeX 风格排版标记。内置原生 OMML 数学公式渲染引擎与 MCP Server（AI Agent 集成）。  
> **面向用户**：开发者、运维人员、自动化脚本编写者、AI Agent 开发者。  
> **开发原则**：Unix 哲学（小而专、可组合）、管道友好、声明式排版、高性能。

---

### 一、项目愿景与命名

- **中文名**：天殇·书契  
  “书契”出自《周易·系辞下》，指文书、契约，涵盖文字（Word）、簿记（Excel）、展示（PPT）三大 Office 文件类型。
- **英文名**：`TianshangScribe`  
  `Scribe` 意为“文书、抄写员”，直接暗示文档处理功能，符合 CLI 工具定位，简洁易读。
- **包名建议**：`tianshang-scribe`（npm/pip/cargo 通用），命令入口 `tianshang-scribe` 或 `scribe`。

---

### 二、核心功能矩阵

| 功能域 | 描述 | 优先级 |
|--------|------|--------|
| **基础 CRUD** | 创建空白文档 / 打开文档 / 保存 / 另存 | P0 ✅ |
| **内容编辑** | 添加文本、表格、图片、分页符；删除指定元素；修改现有内容 | P0 ✅ |
| **样式与排版** | 字体、字号、颜色、加粗、斜体、对齐、行距、段落缩进、标题层级、LaTeX 风格标记、中西文分离字体 | P0 ✅ |
| **数学公式** | LaTeX → OMML 原生渲染，110+ 符号，\frac \sqrt \sum \int，AMS/Springer 字体规范，`--math-font` 可配置渲染字体（默认 Cambria Math，支持 Times New Roman 等 MathType 风格衬线），`--math-mtef` 以 MathType OLE 对象（MTEF）嵌入 | P0 ✅ |
| **模板填充** | 用 JSON/CSV/YAML 数据源填充占位符 `{{key}}`，保留原样式，支持 `{{#each}}` 循环、`{{#if}}` 条件 | P0 ✅ |
| **格式转换** | Word ↔ PDF/Markdown/HTML；Excel ↔ CSV/JSON/HTML；PPT → PDF/图片序列<br>PDF: office2pdf（主，~2MB）+ LibreOffice（回退） | P0 ✅ |
| **MCP Server** | stdio + SSE + Streamable HTTP 三传输，官方 mcp SDK 2.x，7 tools（create/edit/fill/convert/extract/validate/compare），认证/限流/指标，Tool Search（SEP-1821，`tools/list` 支持 `query` 参数），Dify/Coze 可接入 | P1 ✅ |
| **合并与拆分** | 多个文档合并为一个（Word/Excel/PPT 均支持）；拆分：当前仅 Excel 支持 `--split by-sheet` 按工作表拆分，Word 按页 / PPT 按幻灯片拆分待实现 | P1 ✅（合并 ✅；拆分仅 Excel） |
| **保护与元数据** | 设置/解除密码保护；读写作者、标题等文档属性 | P1 ✅ |
| **高级结构操作** | Excel 工作表增删重命名；PPT 幻灯片增删移动；Word 目录生成、分节 | P1 ✅ |
| **批量与管道** | 支持 stdin/stdout、退出码规范 | P1 ✅ |
| **批注与修订** | Word 添加批注（`--comment`）✅；PPT 演讲者备注（`--notes`）✅；开启修订（track changes）待实现 | P2 🚧 |
| **图表与媒体** | Excel 创建图表；PPT 压缩媒体、动画设置 | P2 ✅ |
| **批量增强** | 通配符递归、`--batch` 模式 | P2 ✅ |

---

### 三、命令行接口设计 (CLI)

#### 1. 调用范式
```
tianshang-scribe [全局选项] <输入文件> [操作选项...] [-o 输出文件]
```

- 若不提供 `-w/-e/-p`，工具根据输入文件扩展名自动推断文档类型。
- 支持同时指定多个操作，按顺序执行（例如先替换文本，再转换 PDF）。
- 默认不覆盖原文件，需用 `--force` 或指定不同输出名。

#### 2. 全局选项

| 短选项 | 长选项 | 说明 |
|--------|--------|------|
| `-w` | `--word` | 处理 Word 文档 |
| `-e` | `--excel` | 处理 Excel 工作簿 |
| `-p` | `--ppt` | 处理 PowerPoint 演示文稿 |
| | `--topdf` | 输出为 PDF（对所有类型有效） |
| `-o` | `--output` | 输出文件路径（若省略则打印到标准输出或另存为 `<原文件名>-out.<ext>`） |
| | `--force` | 允许覆盖已有文件 |
| | `--stdin` | 从标准输入读取文档内容（如管道传入） |
| | `--stdout` | 将结果输出到标准输出（二进制数据慎用） |

#### 3. 通用操作参数（所有文档类型）

| 短选项 | 长选项 | 参数示例 | 说明 |
|--------|--------|----------|------|
| `-cr` | `--create` | `--create` | 创建空白文档（可后接 `-w/-e/-p` 指定类型） |
| `-a` | `--add` | `--add "文本内容"` | 添加内容（文本、表格、图片等，结合子参数细化） |
| | `--column` | `--column 2` | 指定 `--add` 的目标列（Excel） |
| `-d` | `--delete` | `--delete "段落关键词"` | 删除指定对象（支持位置、索引或关键字） |
| `-cl` | `--clear` | `--clear content` | 清除内容(`content`)、格式(`formats`)、超链接(`links`)等 |
| `-m` | `--modify` | `--modify "old" --modify-new "new"` | 修改内容（旧值→新值） |
| `-r` | `--replace` | `--replace "search" --replace-new "replacement"` | 查找替换文本（支持 `--regex` 启用正则） |
| | `--regex` | | 配合 `--replace` `--delete` 使用，启用正则表达式 |
| `-x` | `--extract` | `--extract text` | 提取 text/tables/images/structure/metadata/math/latex（images 需 `-o <目录>`；math 把 MathType 公式转为 OMML 写回，latex 输出公式的 LaTeX） |
| `-t` | `--template` | `--template data.json` | 用数据文件填充模板中的 `{{placeholder}}` |
| `-s` | `--style` | `--style "font=Times,size=12,bold,cjk-font=SimSun"` | 设置全局样式或当前段落默认样式（逗号分隔键值对，支持中西文分离字体） |
| | `--merge` | `--merge "a.docx,b.docx"` | 合并多个文档（逗号分隔；不支持通配符；Word/Excel/PPT 均支持 `merge_workbooks`） |
| | `--split` | `--split by-sheet` | 拆分文档（当前仅 Excel 支持：`--split by-sheet` 按工作表拆分；Word 按页 / PPT 按幻灯片拆分待实现） |
| | `--meta` | `--meta title="报告" author="张三"` | 读写文档属性 |
| | `--protect` | `--protect password123` | 设置打开密码或编辑限制 |
| | `--unprotect` | `--unprotect password123` | 解除密码保护 |
| | `--comment` | `--comment "index text"` | 添加批注（Word）/演讲者备注（PPT，追加到备注文本区，与 `--notes` 重叠；格式 `index text`，index 为幻灯片/位置的整型序号） |
| | `--add-table` | `--add-table "H1,H2\|a1,a2"` | 添加 Word 表格；`@file.csv` 从 CSV 读取（内联 `\|` 分行、`,` 分列） |
| | `--batch` | `--batch` | 批量模式：逐文件执行、失败不中断、末尾汇总 |
| | `--files` | `--files "reports/*.docx"` | 批量 glob 通配符（隐含 `--batch`） |
| | `--schedule-db` | `--schedule-db ~/.tianshang-scribe/schedules.db` | 指定调度 SQLite 数据库路径（默认 `~/.tianshang-scribe/schedules.db`） |
| | `--schedule-add` | `--schedule-add "daily\|0 9 * * *\|echo hi"` | 注册调度：`名称\|cron表达式\|命令` |
| | `--schedule-rm` | `--schedule-rm daily` | 按名称删除调度 |
| | `--schedule-list` | `--schedule-list` | 列出已注册调度 |
| | `--schedule-run` | `--schedule-run daily` | 立即按名称运行调度（尊重依赖链） |
| | `--schedule-run-all` | `--schedule-run-all` | 运行所有 cron 窗口满足且依赖就绪的调度 |
| | `--run-script` | `--run-script build.py` | 在沙箱中执行 Python 脚本（import 白名单 + 超时限制） |

#### 4. 文档专属操作

**Word 特有：**
- `--heading` ：添加指定级别的标题（格式：`"level:1 text:标题"`）
- `--math "\frac{a}{b}"` ：添加 LaTeX 数学公式（自动转为原生 OMML）
- `--math-style office|mathtype` ：公式 LaTeX 解析方言（默认 `office` 标准 LaTeX；`mathtype` 兼容 MathType 方言，`\text{}` 保留空格、`~` 视为非断空格）
- `--math-font "Times New Roman"` ：Word OMML 公式渲染字体（默认 `Cambria Math`；如 `Times New Roman`/`Times`/`STIX Two Math`/`Latin Modern Math`，替代 Word 默认数学字体，实现 MathType 风格衬线排版）
- `--math-mtef` ：将公式以 MathType OLE 对象（MTEF 二进制嵌入 `word/embeddings/oleObject*.bin`）插入 Word，供老版 MathType（6.x 及更早）编辑；默认生成 Word 原生 OMML
- `--latex-style` ：启用 LaTeX 风格排版标记解析
- `--toc` ：生成目录
- `--section-break` ：插入分节符
- `--header` / `--footer` ：设置页眉页脚
- `--watermark "机密"` ：添加文字水印
- `--tomd` ：输出为 Markdown
- `--tohtml` ：输出为 HTML
- `--add-table "H1,H2\|a1,a2"` ：添加表格（内联或 `@file.csv`）

**Excel 特有：**
- `--sheet-add "SheetName"` ：添加工作表
- `--sheet-delete "SheetName"` ：删除工作表
- `--sheet-rename "OldName" "NewName"` ：重命名工作表
- `--column-width 2=20` ：设置第 2 列宽度（索引从 1 开始）
- `--row-height 3=30` ：设置第 3 行行高
- `--formula A1 "=SUM(B1:B10)"` ：设置公式
- `--sheet "Sheet2"` ：指定后续 Excel 操作（写入/公式/排序/图表/导入导出等）的目标工作表，默认活动工作表
- `--from-csv data.csv` ：从 CSV 导入数据
- `--to-csv` ：导出为 CSV
- `--to-json` ：导出为 JSON
- `--to-html` ：导出为 HTML 表格
- `--sort A1:A10 asc` ：排序
- `--chart-add type=bar data=B1:C10` ：创建图表
- `--freeze "A2"` ：冻结窗格（冻结 `A2` 上方的行与左侧的列）
- `--number-format "A1:A10=0.00%"` ：设置数字格式（支持 `0.00%` / `yyyy-mm-dd` / `#,##0`）
- `--conditional-format "B2:B100=color_scale"` ：条件格式（`color_scale` / `data_bar` / `cell_is:operator:formula`）
- `--data-validation "C2:C50=list:yes,no"` ：数据验证（`list` / `whole` / `decimal` / `date` / `text_length`，`whole`/`decimal` 可用 `min:max`）

> **反向转换**：`.md/.markdown/.html` 作为输入时自动转换为 Word（`tianshang-scribe doc.md -o out.docx`）；`.json`（对象数组/数组）作为输入时自动导入为 Excel。

**PPT 特有：**
- `--slide-add` ：添加幻灯片
- `--slide-delete 3` ：删除第 3 张幻灯片
- `--slide-move 2 4` ：移动幻灯片位置
- `--layout "Title and Content"` ：应用版式
- `--notes "演讲提示文字"` ：添加演讲者备注
- `--toimg output_dir/` ：导出为图片序列
- `--transition "fade"` ：设置幻灯片切换效果
- `--compress-media "1920,80"` ：压缩媒体（最大边长,JPEG 质量）
- `--ppt-table "H1,H2\|a1,a2"` ：在末张幻灯片插入表格（首行为表头）
- `--ppt-chart "bar\|S1,S2\|Cat1,1,2\|Cat2,3,4"` ：在末张幻灯片插入图表（`type\|系列名\|分类行...`）

#### 5. LaTeX 风格排版标记（Word / PPT 文本内容）

通过 `--add` 或 `--modify` 输入时，若启用 `--latex-style`（或默认智能检测），工具解析以下标记：

| 标记语法 | 效果 | 示例 |
|----------|------|------|
| `\bfseries{text}` | 加粗 | `\bfseries{重要}` |
| `\itshape{text}` | 斜体 | `\itshape{强调}` |
| `\scshape{text}` | 小型大写 | |
| `\rmfamily{text}` | 衬线字体 | |
| `\sffamily{text}` | 无衬线字体 | |
| `\ttfamily{text}` | 等宽字体 | |
| `\fontfamily{Arial}{text}` | 指定字体 | |
| `\fontsize{14}{text}` | 字号（pt） | `\fontsize{18}{大标题}` |
| `\color{FF0000}{text}` | 文本颜色（十六进制） | |
| `\underline{text}` | 下划线 | |
| `\centering{...}` | 居中段落 | |
| `\raggedright{...}` | 左对齐 | |
| `\raggedleft{...}` | 右对齐 | |
| `\linespread{1.5}{...}` | 行距倍数 | |
| `\indent{...}` / `\noindent{...}` | 首行缩进/无缩进 | |
| `\heading{2}{小标题}` | 插入 2 级标题 | |
| `\newpage` | 分页符（可独立使用） | |
| `\includegraphics{path}` | 插入图片 | |

**嵌套规则**：支持三层嵌套（如 `\centering{\fontsize{18}{\bfseries{标题}}}`）。解析器从左到右、深度优先遍历，合并样式。段落级命令（`\centering`、`\raggedright` 等）递归解析内部 LaTeX 标记。

#### 6. 交互式文件界面（`open` 子命令）

`tianshang-scribe open <文件> [--latex-style] [-w|-e|-p]` 打开文档并进入交互式 REPL，可持续对内存中的文档进行操作：

```
doc.docx> add "Hello \bfseries{World}"
doc.docx> heading 2 "报告标题"
doc.docx> table "H1,H2|a1,a2"          # 或 table @data.csv
doc.docx> replace "旧" "新"
doc.docx> extract tables
doc.docx> save                         # 显式保存（默认写回原文件）
doc.docx> quit                         # 有未保存修改时先询问
```

- **命令集**：`add` `heading` `table` `math` `replace` `delete` `style` `extract` `info` `path` `save` `help` `quit`
- 交互会话在**单线程**持有文档对象，命令实时修改内存；`save` 才落盘，`quit` 时若有未保存修改会询问
- 双 app 分发：`open` 走 `open_app`（`src/tianshang_scribe/cli/repl.py` 的 `InteractiveSession`），其余走一键式 `app`；两者共享 `--latex-style`/`-w/-e/-p` 选项定义（`src/tianshang_scribe/cli/main.py` 中的共享常量），保证 CLI 参数同步
- 一键式命令为普通 Typer 命令，选项可在位置参数 `input_file` 之前或之后（如 `tianshang-scribe file.docx --add "hi"`）

---

### 四、技术架构

#### 1. 语言与核心库
- **推荐语言**：Python 3.10+（生态丰富，开发效率高）
- **文档处理库**：
  - `python-docx` : Word 读/写/样式
  - `openpyxl` : Excel 读/写/公式/样式
  - `python-pptx` : PowerPoint 操作
- **CLI 框架**：`typer` + `rich`（美观终端输出）
- **LaTeX 标记解析**：自定义简单递归下降解析器 + 正则备选
- **模板引擎**：自研占位符替换器（支持 `{{#each}}` 循环、`{{#if}}` 条件）
- **格式转换**：
  - Word/MD/HTML 转换可集成 `pandoc`（调用外部命令或 Python 绑定）
  - PDF 输出：`office2pdf`（主引擎，~2MB Rust 二进制，零运行时依赖）+ LibreOffice headless（高保真回退）
- **MCP Server**：官方 `mcp` SDK 2.x（`MCPServer`），stdio/SSE/Streamable HTTP 三传输，工具 schema 由 `Annotated` 签名自动派生
- **打包分发**：`PyInstaller` 单文件（~35MB EXE）、`pipx` 安装、Docker 镜像

#### 2. 模块划分
```
src/                          # 构建隔离目录
└── tianshang_scribe/         # 可 import 包（tianshang_scribe.*，对标 TianshangCAD 的 tianshangcad）
    ├── cli/               # 命令解析与分发
    │   ├── main.py        # 入口，定义 Typer 应用（一键式 + open 双 app 分发）
    │   ├── repl.py        # 交互式文件界面（InteractiveSession REPL）
    │   └── global_opts.py # 全局选项处理、路径推断、类型解析
    ├── core/              # 文档模型抽象层
    │   ├── document.py    # 统一接口 DocumentABC + DocumentType
    │   ├── word_engine.py # Word 引擎
    │   ├── excel_engine.py# Excel 引擎
    │   ├── ppt_engine.py  # PPT 引擎
    │   ├── scheduler.py   # Cron 调度器 + 依赖链（--schedule-*）
    │   └── script_runner.py# 沙箱脚本执行（--run-script，import 白名单 + 超时）
    ├── rendering/         # 模板填充、样式解析、数学公式
    │   ├── template.py    # 模板引擎（{{key}}, {{#each}}, {{#if}}）
    │   ├── latex_parser.py# LaTeX 标记递归下降解析器（20 命令，三层嵌套）
    │   ├── math_omml.py   # LaTeX → OMML 数学公式转换器（110+ 符号）
    │   ├── styles.py      # TextStyle 数据类（中西文字体分离）
    │   └── mtef/          # MathType 兼容（MTEF/OLE ↔ LaTeX → OMML）
    │       ├── ole_util.py    # OLE 复合文档（CFB）解析器
    │       ├── mtef_reader.py # MTEF 二进制读取器（→ LaTeX）
    │       ├── mtef_writer.py # LaTeX → MTEF 二进制写入器
    │       ├── cfb_writer.py  # MTEF → OLE 复合文件（make_ole）
    │       └── symbols.py     # LaTeX ↔ mtcode+typeface 双向符号表（281 项）
    ├── transform/         # 格式转换器
    │   ├── pdf.py         # PDF/MD/HTML 转换（office2pdf + LibreOffice）
    │   └── reverse.py     # 反向转换（HTML/Markdown→Word）
    ├── mcp/               # MCP Server（AI Agent 集成，官方 mcp SDK 2.x）
    │   ├── server.py      # build_server + 入口（stdio / SSE / Streamable HTTP）
    │   ├── transport.py   # 传输接线 + ASGI 中间件（认证/CORS/限流/指标/RBAC）
    │   ├── schemas.py     # pydantic 模型 + as_dict 规范化
    │   ├── auth.py        # Bearer Token 认证
    │   ├── rate_limit.py  # 令牌桶限流
    │   ├── metrics.py     # Prometheus 风格指标
    │   ├── security.py    # 工具只读/破坏性分类 + RBAC 角色矩阵
    │   ├── tool_search.py # Tool Search（SEP-1821）：tools/list 支持 query 参数评分搜索
    │   ├── prompts.py     # 5 个提示词工作流（prompts/list）
    │   ├── tools/          # 7 个 Agent 工具实现
    │   │   ├── _registry.py# 工具注册表（schema 自动派生）
    │   │   ├── create.py  # create_office_document
    │   │   ├── edit.py    # edit_office_document
    │   │   ├── template.py# fill_template
    │   │   ├── convert.py # convert_document + extract_document_data
    │   │   ├── validate.py# validate_template
    │   │   └── compare.py # compare_documents（快照 snapshot/list/restore）
    │   └── errors.py      # 结构化错误码 + 修复建议
    └── utils/             # 公共工具
        ├── config.py      # pydantic-settings 集中配置（TIANSHANG_SCRIBE_* env + .env）
        ├── logging.py     # structlog 结构化日志（console/JSON，uvicorn 统一格式）
        ├── store.py       # SQLite 持久化（ScheduleStore：调度 + 运行历史）
        └── file_utils.py  # check_overwrite / ensure_parent_dir
```

#### 3. 管道与组合原则
- 支持从 `stdin` 读取文档二进制数据，向 `stdout` 输出结果，以便链式调用。
- 退出码：`0` 成功，`1` 一般错误，`2` 参数错误，`3` 文档损坏等，遵循 shell 规范。
- 错误信息输出到 `stderr`，保持输出纯净。

---

### 五、开发路线图

#### Phase 1：MVP（最小可行产品）✅ 已完成
- 完成 CLI 框架搭建，实现 `-cr`、`-a`（纯文本添加）、`-r`、`-s`（基础字体样式）、`--topdf` 转换。
- 支持 Word 和 Excel，PPT 仅支持创建和简单文本。
- LaTeX 标记解析核心（加粗、斜体、字号、字体族、颜色）。
- 模板填充基础版（纯文本替换，无循环）。
- 输出 `--help` 完整文档。

#### Phase 2：功能完备 ✅ 已完成
- ✅ 实现所有通用操作参数（`-d`, `-cl`, `-m`, `-x`, `--merge`, `--split`, `--meta`, `--protect`, `--comment`）。
- ✅ 完善文档专属操作（标题、表格、工作表管理、幻灯片操作）。
- ✅ Excel 导入导出 CSV/JSON，PPT 转图片。
- ✅ 完整 LaTeX 标记集，支持段落级格式、三层嵌套。
- ✅ 数学公式 OMML 渲染（110+ 符号，`\frac \sqrt \sum \int`）。
- ✅ MCP Server（官方 mcp SDK 2.x：stdio/SSE/Streamable HTTP 三传输，7 tools，认证/限流/指标中间件，Tool Search SEP-1821）。
- ✅ PDF 引擎：office2pdf（~2MB）+ LibreOffice 回退。
- ✅ 模板引擎 `{{#each}}` 循环 + `{{#if}}`/`{{#unless}}` 条件。
- ✅ 单二进制分发（PyInstaller ~35MB EXE）。
- ✅ 发布到 PyPI / GitHub Releases（v0.2.0）。
- ✅ 批量模式（`--batch` / `--files` glob 通配符，逐文件执行、失败不中断）。
- ✅ reverse 格式转换（HTML/Markdown→Word、JSON→Excel，input_file 自动识别）。
- ✅ `--extract` 全量（text / tables / images / structure / metadata）。
- ✅ `--add-table`（内联或 `@file.csv`）与 PPT `--compress-media` 媒体压缩。
- ✅ 调度与沙箱：cron 调度器 + SQLite store（`--schedule-*`）、沙箱脚本执行（`--run-script`，import 白名单 + 超时）。
- ✅ 文档快照与 RBAC：`compare_documents`（snapshot/list_snapshots/restore）+ 角色矩阵（viewer/editor/owner）。

#### Phase 3：优化与生态
- 性能优化：大文件流式读写、多线程转换。
- 插件机制：支持用户自定义扩展操作（Python 脚本）。
- 远程文件支持（S3、HTTP）。
- VSCode 扩展或 GitHub Action 集成。
- Coze / Dify 工具市场上架。
- 文档与交互教程。

---

### 六、开发约定与规范

1. **命令命名**：一律使用英文小写短选项，长选项采用全小写+连字符；子参数键值对采用 `key=value` 格式。
2. **错误处理**：所有异常应捕获并转化为用户友好的错误信息，包含错误码和修复建议。
3. **测试要求**：每个操作需覆盖正常、异常、边界用例；CI 覆盖 Windows、macOS、Linux（9-matrix: 3 OS × 3 Python）。
4. **文档同步**：AGENTS.md 作为本项目的 AI 代理指导文档，必须与代码同步更新，描述最新的 CLI 语法和功能。
5. **许可协议**：Apache 2.0，需注明依赖库许可证合规。

---

*本文档将作为项目开发的唯一事实来源，所有贡献者及 AI 助手在生成代码、回答问题、修改功能时需严格遵循此设计。*

### 版本记录

| 版本 | 日期 | 核心变更 |
|------|------|---------|
| v0.1.0 | 2026-07 | MVP：Word/Excel/PPT 基础 CRUD、LaTeX 标记、模板填充 |
| v0.1.1 | 2026-07 | PyPI 发布、EXE 打包、CI 9-matrix、双语 README |
| v0.2.0 | 2026-07 | SSE MCP Server、office2pdf PDF 引擎、数学公式 OMML、`{{#if}}` 条件模板 |
| v0.3.0 | 2026-08 | 工程基线对齐：mcp 包迁入 `src/tianshang_scribe/mcp/`、官方 mcp SDK 2.x（stdio/SSE/Streamable HTTP）、工具 schema 自动派生（7 tools）、认证/限流/指标中间件、严格工具链（ruff/mypy/pytest -W error）、治理文档五件套；CLI 双 app 分发（一键式 + `open` 交互 REPL）、选项顺序修复、reverse 转换、`--extract` 全量、`--add-table`、`--batch`、`--compress-media`；MCP 工具描述三句模板（副作用/只读/替代披露）与注解一致性加固 |
| v0.4.0 | 2026-08 | 生产可观测性 + 工程里程碑（M1-M3）：SEP-1821 Tool Search（`tools/list` 支持 `query` 评分搜索）、pydantic-settings 集中配置（`SCRIBE_*` env + `.env`）、HTTP 认证 401/403 区分、structlog 结构化日志（console/JSON，uvicorn 统一格式）、Docker 多阶段加固（HEALTHCHECK、非 root、Streamable HTTP 默认）、修复 Excel `{{#each}}` 首项丢失缺陷；M1 覆盖率门禁 70→80（830 用例，92.93%）；M2 cron 调度器 + 沙箱脚本执行 + SQLite store（`--schedule-*` / `--run-script`，56 用例）；M3 `compare_documents` 文档快照（snapshot/list_snapshots/restore，Discriminated-Union action 分发）+ RBAC 角色矩阵（viewer/editor/owner，`X-Scribe-Role` 中间件） |
| v0.5.0 | 2026-08 | **破坏性变更**：import 包名 `src` → `tianshang_scribe`（`src/tianshang_scribe/`，对标 TianshangCAD 的 `src/tianshangcad/`），`src/` 仅作构建隔离目录；`from src.x` → `from tianshang_scribe.x`，`python -m src.mcp.server` → `python -m tianshang_scribe.mcp.server`，coverage source 改 `tianshang_scribe`；console scripts（`tianshang-scribe`/`scribe-mcp`）不变，迁移说明见 `MIGRATION.md` |
| v0.6.0 | 2026-08 | **破坏性变更**：环境变量前缀 `SCRIBE_*` → `TIANSHANG_SCRIBE_*`（pydantic-settings `env_prefix`，`config.py` 一处 + 全仓引用同步）；MCP Server 命令 `scribe-mcp` → `tianshang-scribe-server`（对标 TianshangCAD 的 `tianshangcad-server`，`python -m tianshang_scribe.mcp.server` 不变）；Dockerfile/docker-compose env 同步；Prometheus 指标名统一 `tianshang_scribe_` 前缀（`scribe_operation_duration_seconds`→`tianshang_scribe_operation_duration_seconds`、`scribe_operations_total`→`tianshang_scribe_operations_total`）；docker-compose 命名卷 `scribe_output`→`tianshang_scribe_output`；迁移说明见 `MIGRATION.md` |
| v0.7.0 | 2026-08 | MathType MTEF 写入路径：`--math-mtef` 以真实 MathType OLE 对象（MTEF 二进制）嵌入 Word、`--math-font` 可配置 OMML 渲染字体（默认 Cambria Math）；数学公式转换器重构为递归下降解析器（表达式→项→因子→原子）+ 黄金快照回归套件 |
| v0.7.1 | 2026-08 | 修复 `--comment` 在 PPT 上的索引解析崩溃（原将 slide_index 传为字符串导致 `TypeError`）；文档与实现对齐（`--split` 仅 Excel 支持、`--merge` 逗号分隔不支持通配符、`--comment` PPT 写备注区、ROADMAP `tools_available` 7） |
| v0.8.0 | 2026-08 | **Excel/PPT 引擎缺陷修复与能力补全（MINOR）**：PPT `merge_workbooks` 改为关系感知的真实幻灯片深拷贝（含图片/媒体/图表，修复此前仅生成空白幻灯片）；PPT `add_text`/`add_styled_content` 修复多段文本/公式重叠定位并支持 `slide_index` 追加到已有幻灯片；PPT 修改保护改用合规 SHA-512+盐+10万次迭代的 ECMA-376 敏捷加密（原明文密码无效）；PPT `to_images` 改为「先转 PDF 全页再逐页栅格化」（PyMuPDF/pdftoppm），修复 LibreOffice 仅导出首屏；Excel `sort` 支持多列键与混合类型安全排序（整行保真，不再 `TypeError`）；Excel 新增 `--sheet` 选项选择目标工作表；Excel 新增冻结窗格 `--freeze`、数字格式 `--number-format`、条件格式 `--conditional-format`、数据验证 `--data-validation`、边框/填充 `set_range_style`、图表类型扩展（area/scatter/doughnut）、超链接/命名范围/自动列宽；PPT 新增精确文本框 `add_textbox`、表格 `add_table`、图表 `add_chart`、图片 `add_picture`、形状 `add_shape`、`replace_text` 跨 run 保留格式，以及 CLI `--ppt-table`/`--ppt-chart`；`SERVER_VERSION` 同步至 0.8.0 |
| v0.8.0 | 2026-08 | **MCP Server 能力对齐（P0）**：7 个工具不变，但 `create_office_document`/`edit_office_document` 现已暴露上述 Excel/PPT 引擎能力——`ContentBlock`/`EditOperation` 新增可选字段（`slide_index`/`slide_layout`/`notes`/`transition`/`sheet_name`/`cell`/`formula`/`chart_type`/`chart_data_range`/`chart_data`/`rows`/`number_format`/`conditional_format`/`data_validation`/`freeze`/`hyperlink`/`named_range`）；修复 `create_office_document` 在 PPT 表格上因 `add_table` 签名不匹配导致的崩溃，PPT 多段内容现堆叠至同一张幻灯片；新增 `src/tianshang_scribe/mcp/tools/_parse.py` 统一解析数字格式/条件格式/数据验证/PPT 图表/幻灯片索引 |