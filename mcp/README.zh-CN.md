# TianshangScribe MCP Server

> [English](./README.md)

MCP（Model Context Protocol）服务端，用于 Office 文档处理。让 AI Agent 能**创建**、**编辑**、**模板填充**、**格式转换**和**数据提取** Word、Excel、PowerPoint 文档，原生支持 LaTeX 样式标记和数学公式渲染。

## 快速开始

```bash
git clone https://github.com/Tianshang301/TianshangScribe.git
cd TianshangScribe
pip install -e ".[dev]"

python mcp/test_server.py
```

**Claude Code / Cursor 配置**：

```json
{
  "mcpServers": {
    "tianshang-scribe": {
      "command": "python",
      "args": ["-m", "mcp.server"],
      "cwd": "/path/to/TianshangScribe"
    }
  }
}
```

## 工具列表

### 1. `create_office_document`

用结构化内容创建 Word、Excel、PPT 文档。

```
用户：   "生成一份 Q3 财务报告"
Agent：  create_office_document(format="docx", content=[...])
```

**参数**

| 参数 | 类型 | 必填 | 说明 |
|------|------|------|------|
| `format` | `"docx"` \| `"xlsx"` \| `"pptx"` | 是 | 输出文档格式 |
| `content` | `ContentBlock[]` | 是 | 有序内容块数组 |
| `style` | `string` | 否 | 全局样式：`"font=SimSun,size=12,bold"` |
| `template_data` | `object` | 否 | 键值对，填充 `{{placeholder}}` |
| `metadata` | `object` | 否 | 文档属性：`{"title": "报告", "author": "AI"}` |
| `output_path` | `string` | 否 | 输出路径（省略自动生成） |
| `options` | `object` | 否 | `{"dry_run": true, "backup": true}` |

**ContentBlock 类型**

| `type` | 字段 | 说明 |
|--------|------|------|
| `paragraph` | `text`, `style` | 格式化文本。支持 `\bfseries{}`、`$...$` 等 |
| `heading` | `text`, `level` (1-6), `style` | 标题 |
| `formula` | `text` | LaTeX 数学公式，转为原生 OMML |
| `table` | `rows`（二维数组） | 数据表格 |
| `image` | `path` | 插入图片 |
| `page_break` | — | 分页符 |

**示例**

```json
{
  "format": "docx",
  "content": [
    {"type": "heading", "text": "执行摘要", "level": 1},
    {"type": "paragraph", "text": "\\bfseries{营收：}\\color{0000FF}{1250 万元}，同比增长 \\itshape{15.3%}。"},
    {"type": "formula", "text": "\\sum_{i=1}^{n} x_i = \\frac{n(n+1)}{2}"},
    {"type": "table", "rows": [["Q1", "350万"], ["Q2", "420万"], ["Q3", "480万"]]}
  ],
  "metadata": {"title": "Q3 报告", "author": "AI Agent"}
}
```

### 2. `edit_office_document`

对已有文档执行一系列编辑操作。

**操作类型**

| `action` | 关键字段 | 说明 |
|----------|---------|------|
| `replace` | `old_text`, `new_text`, `regex` | 查找替换 |
| `delete` | `target`, `regex` | 删除内容 |
| `modify` | `old_text`, `new_text` | 修改内容（非正则） |
| `style` | `style`, `apply_all` | 设置全局样式 |
| `add` | `text`, `column` | 添加文本（支持 Excel 列） |
| `clear` | — | 清除单元格内容 |

```json
{
  "input_path": "contract.docx",
  "operations": [
    {"action": "replace", "old_text": "甲方", "new_text": "乙方"},
    {"action": "style", "style": "font=SimSun,size=14", "apply_all": true}
  ]
}
```

### 3. `fill_template`

用结构化数据填充模板中的占位符。支持嵌套键（`{{user.name}}`）和循环（`{{#each items}}...{{/each}}`）。

```json
{
  "template_path": "invitation.docx",
  "data": {
    "name": "张三",
    "event": "AI 峰会",
    "date": "2026-09-15"
  }
}
```

### 4. `convert_document`

格式转换。

| 源格式 | 目标格式 | 支持 |
|--------|---------|------|
| `docx` | `pdf`, `md`, `html` | 全部 |
| `xlsx` | `pdf`, `csv`, `json`, `html` | 全部 |
| `pptx` | `pdf` | 是 |

```json
{
  "input_path": "report.docx",
  "target_format": "pdf",
  "output_path": "report.pdf"
}
```

### 5. `extract_document_data`

从文档中提取数据。

| `mode` | 返回内容 |
|--------|---------|
| `metadata` | 作者、标题、主题、关键词 |
| `text` | 全文文本 + 块计数 |
| `structure` | 段落/节（Word）、工作表（Excel）、幻灯片（PPT） |

```json
{
  "input_path": "report.docx",
  "mode": "text"
}
```

## LaTeX 标记参考

所有 `text` 字段支持 LaTeX 风格标记：

| 标记 | 效果 |
|------|------|
| `\bfseries{文字}` | **加粗** |
| `\itshape{文字}` | *斜体* |
| `\underline{文字}` | 下划线 |
| `\scshape{文字}` | 小型大写 |
| `\color{FF0000}{文字}` | 红色文字 |
| `\fontsize{24}{文字}` | 字号（磅） |
| `\fontfamily{SimHei}{文字}` | 指定字体 |
| `\heading{N}{标题}` | N 级标题 |
| `\newpage` | 分页符 |
| `\centering{...}` | 居中 |
| `$E=mc^2$` | 行内公式 |
| `$$x=1$$` | 行间公式 |

## 错误处理

所有工具返回结构化错误：

```json
{
  "success": false,
  "error_code": 1002,
  "error_message": "文档受密码保护。",
  "suggested_fix": "请提供密码或先解除保护。",
  "retryable": true
}
```

**错误码**

| 码 | 名称 | 可重试 |
|----|------|--------|
| 0 | 成功 | — |
| 1001 | 文档未找到 | 否 |
| 1002 | 文档被锁定 | 是 |
| 1003 | 不支持的格式 | 否 |
| 1004 | 模板错误 | 是 |
| 1005 | 转换失败 | 是 |
| 1006 | 参数无效 | 否 |
| 9999 | 内部错误 | 否 |

## Dry Run 与备份

所有工具支持 `options`：

```json
{
  "options": {
    "dry_run": true,
    "backup": true,
    "deterministic_id": "uuid"
  }
}
```

- `dry_run`：预览变更，不写入文件
- `backup`：修改前创建 `.bak` 备份
- `deterministic_id`：可追踪的操作 ID

## 架构

```
MCP Client (Claude Code / Cursor / Agent)
    ↕ stdio JSON-RPC 2.0
mcp/server.py          ← 协议分发
    ↕
mcp/tools/
    ├── create.py      ← WordEngine / ExcelEngine / PptEngine
    ├── edit.py        ← replace_text / set_style / clear_content
    ├── template.py    ← TemplateEngine
    └── convert.py     ← pdf.py / export 方法
    ↕
src/core/              ← 文档引擎（python-docx / openpyxl / python-pptx）
```

- **零外部 MCP 依赖** — 纯 stdio JSON-RPC 2.0 实现
- **传输层**：stdio（Phase 1），SSE 计划中（Phase 2）
- **协议**：MCP 2024-11-05

## 许可

Apache-2.0
