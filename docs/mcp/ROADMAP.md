# TianshangScribe MCP Server · 完善路线图

> 基于当前实现（官方 `mcp` SDK 2.x、7 tools、stdio/SSE/Streamable HTTP 三传输）的系统性改进方案。

---

## 一、协议合规性：补齐 MCP 标准能力

### 1.1 `resources` — 文档可读性暴露

**当前**：Agent 需通过 `extract_document_data` 工具调用获取纯文本摘要。

**目标**：将处理过的文档暴露为 MCP Resource，Agent 可直接在上下文中"阅读"。

| 端点 | 说明 |
|------|------|
| `resources/list` | 列出当前 session 可访问的所有文档资源 |
| `resources/read` | 按 URI 读取文档内容（文本/JSON/图片） |

**URI 设计**：
```
file:///tmp/report.docx                    # 文件系统路径
scribe://session/{session_id}/content      # 当前操作的文档内容
scribe://session/{session_id}/structure    # 当前文档的结构摘要
scribe://session/{session_id}/preview      # 当前文档的 PNG 预览
```

**MIME types**：`text/plain` | `application/json` | `image/png`

**实现要点**：
- 每次 `create` / `edit` / `fill` 成功后自动注册 resource URI
- `resources/read` 触发 `extract_document_data`，但返回 MCP 标准 resource 格式
- Session 过期或文档删除时取消注册

**Agent 体验变化**：
> 改进前：Agent 用 `extract_document_data` → 得到 JSON 摘要 → 自行拼凑上下文
> 改进后：Agent 调用 `resources/read scribe://session/x/content` → 直接在对话中看到全文

---

### 1.2 `prompts` — 内置工作流模板

**目标**：提供预置 Prompt，降低 Agent 的调用链复杂度。

Agent 调用 `prompts/list` → `prompts/get {id}` → 获得完整 system prompt + 推荐 tool 序列。

| Prompt ID | 场景 | 嵌入上下文 |
|-----------|------|-----------|
| `generate_report` | 根据数据生成带目录的 Word 报告 | `{data_source, template_hint}` |
| `batch_fill_templates` | 用 CSV 数据批量填充邀请函模板 | `{csv_path, template_path}` |
| `convert_and_archive` | 将 Excel 转为 PDF 并添加水印 | `{input_glob, watermark_text}` |
| `extract_and_analyze` | 提取文档结构并分析关键指标 | `{document_path}` |
| `create_presentation` | 根据大纲生成 PPT 演示文稿 | `{outline, theme}` |

**实现要点**：
- 每个 prompt 定义为静态 data class，包含 `description`、`arguments`、`messages`
- `prompts/list` 返回 5 个 prompt 元数据
- `prompts/get` 返回完整 messages 数组供 Agent 注入上下文

---

### 1.3 `sampling` — Server 反向请求 LLM 协助

**场景**：模板条件复杂、内容语义不明确时，Server 主动向 Agent 请求 LLM 协助。

```json
{
  "jsonrpc": "2.0",
  "method": "sampling/createMessage",
  "params": {
    "messages": [{
      "role": "user",
      "content": "请分析以下模板中的条件占位符逻辑..."
    }],
    "maxTokens": 500
  }
}
```

**适用场景**：
- 模板中的 `{{#if condition}}` 条件表达式无法直接求值
- 文档内容质量检查（"这段文字是否有语法错误？"）
- 智能格式推荐（"这个表格数据用什么图表类型最合适？"）

---

## 二、生产就绪：安全加固与可观测性

### 2.1 SSE 安全加固

| 机制 | 实现方式 | 配置 |
|------|---------|------|
| Bearer Token 认证 | `Authorization: Bearer <token>` 头验证 | `--auth-token $TIANSHANG_SCRIBE_TOKEN` |
| CORS 白名单 | 替换当前通配 `*` 为可配置 origin | `--cors-origins "https://coze.com,https://dify.ai"` |
| 请求签名 | HMAC-SHA256 对 `POST /message` body 签名 | `--sign-secret $TIANSHANG_SCRIBE_SECRET` |
| IP 白名单 | 按来源 IP 过滤 | `--allow-from "10.0.0.0/8,192.168.1.0/24"` |

### 2.2 限流与资源隔离

```
MAX_CONCURRENT_PER_SESSION = 3
MAX_FILE_SIZE_MB = 50

TIMEOUTS = {
    "pdf_convert": 120,   # PDF 转换
    "word_edit": 30,       # 文档编辑
    "extract": 10,         # 提取操作
}

MAX_MEMORY_MB = 512
```

### 2.3 健康检查与 Metrics

**Health Endpoint**：`GET /health`
```json
{
  "status": "ok",
  "version": "0.2.0",
  "uptime_seconds": 86400,
  "active_sessions": 3,
  "tools_available": 5,
  "pdf_engine": "office2pdf"
}
```

**Prometheus Metrics**：
| Metric | 类型 | 说明 |
|--------|------|------|
| `tianshang_scribe_tools_total{tool,format}` | Counter | 各工具调用次数 |
| `tianshang_scribe_duration_seconds{tool}` | Histogram | 操作耗时分布 |
| `tianshang_scribe_errors_total{code}` | Counter | 错误码分布 |
| `tianshang_scribe_file_size_bytes{format}` | Summary | 处理的文件大小 |
| `tianshang_scribe_active_sessions` | Gauge | 当前活跃 session 数 |

**Structured Logging**：JSON 格式，每条日志含 `timestamp`、`session_id`、`tool`、`format`、`duration_ms`、`file_size_bytes`、`success`。

### 2.4 优雅关闭与 Session 持久化

| 需求 | 方案 |
|------|------|
| Graceful Shutdown | 捕获 SIGTERM → 拒绝新连接 → 等待活跃操作完成（最多 30s）→ 关闭 |
| Session 持久化 | 支持 `--session-backend redis://localhost:6379`，多实例共享 session |
| 连接重连 | SSE 连接断开后，客户端携带 `session_id` 重新连接即可恢复 |

---

## 三、Agent 体验优化

### 3.1 进度通知（Progress Notifications）

```json
{
  "jsonrpc": "2.0",
  "method": "notifications/progress",
  "params": {
    "progressToken": "uuid-abc",
    "progress": 65,
    "total": 100,
    "message": "Rendering page 13/20..."
  }
}
```

### 3.2 批量操作原子性

```json
{
  "arguments": {
    "operations": [...],
    "options": {"atomic": true, "backup": true}
  }
}
```

- `atomic: true`：操作前创建 `.bak` 副本，任一步骤失败回滚
- 扩展 `batch` tool：跨文档原子操作（`[edit A, convert A→PDF, edit B]`）

### 3.3 Schema 增强

```python
"format": {
    "type": "string",
    "enum": ["docx", "xlsx", "pptx"],
    "default": "docx",
    "description": (
        "Document format to create.\n"
        "- 'docx': Word document (reports, letters, contracts)\n"
        "- 'xlsx': Excel workbook (spreadsheets, tables, charts)\n"
        "- 'pptx': PowerPoint presentation (slides, decks)"
    ),
    "examples": ["docx", "xlsx"]
}
```

### 3.4 返回值多类型增强

```json
{
  "content": [
    {"type": "text", "text": "Document created successfully."},
    {"type": "resource", "resource": {
      "uri": "file:///output/report.docx",
      "mimeType": "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
      "title": "report.docx",
      "size": 38400
    }},
    {"type": "image", "data": "iVBORw0KGgo...", "mimeType": "image/png"}
  ]
}
```

---

## 四、功能扩展：新工具建议

### 4.1 新增工具清单

| 工具 | 优先级 | 输入 | 输出 | 说明 |
|------|--------|------|------|------|
| `validate_template` | **高** | `template_path`, `data` | 缺失占位符清单、类型不匹配警告 | 模板验证 |
| `compare_documents` | 中 | `path_a`, `path_b` | 差异摘要（增/删/改段落） | 文档差异对比 |
| `generate_preview` | 中 | `input_path`, `page=1` | `image/png` base64 | 文档首页缩略图 |
| `merge_documents` | 中 | `paths[]`, `output_path` | 合并后文档路径 | 多文档合并 |
| `add_watermark` | 低 | `input_path`, `text`, `output_path` | — | 添加/移除水印 |
| `protect_document` | 低 | `input_path`, `password`, `output_path` | — | 加密/解密 |

### 4.2 Excel 创建增强

| `type` | 说明 |
|--------|------|
| `sheet` | 定义工作表：名称、列宽、行高、冻结窗格 |
| `formula` | 写入公式：`"=SUM(B2:B10)"` |
| `chart` | 创建图表：`{type: "bar", data_range: "A1:B10", position: "D1"}` |
| `conditional_format` | 条件格式：`{rule: ">100", style: "color=FF0000"}` |
| `pivot_table` | 数据透视表 |

### 4.3 PPT 创建增强

| `type` | 说明 |
|--------|------|
| `slide` | 单页定义：`{layout, notes, transition}` |
| `theme` | 全局参数：`{colors, fonts, background}` |

---

## 五、实施路线图

### Phase 1（1-2 周）— 协议补齐 + Quick Win

```
resources/list + resources/read       # 文档可读性
validate_template 工具                 # 最高性价比新工具
Schema 增强（enum + examples）        # Agent 调用准确率
返回值多类型（text + resource uri）    # Quick Win
```

### Phase 2（2-3 周）— 生产加固

```
Bearer Token 认证 + CORS 白名单       # SSE 安全
GET /health + Prometheus metrics      # 可观测性
Progress Notifications                # 长操作反馈
Structured Logging (JSON)             # 日志标准化
Session 持久化 (Redis)                # 多实例部署
```

### Phase 3（3-4 周）— 高级能力

```
prompts/list + prompts/get            # 工作流模板
Batch 原子操作                         # 跨文档事务
compare_documents + generate_preview  # 文档处理
sampling/createMessage                # 反向 LLM 协助
Excel/PPT create Block 类型扩展        # 富内容创建
```

---

## 六、立即可做（≤30 分钟）

在 `mcp/tools/create.py` 的返回值中，将纯文本改为多类型 content，增加 `resource` 类型指向生成的文件。

### 建议推进顺序

1. **立刻**：返回值多类型 + Schema 增强（2 个文件，~30 行改动）
2. **本周**：`GET /health` + Bearer Token（3 个文件，~80 行）
3. **下周**：`resources/list` + `resources/read`（2 个文件，~150 行）
4. **本月**：`validate_template` 工具 + Prometheus metrics

---

*最后更新：2026-07-31 · 对应版本 v0.2.0*
