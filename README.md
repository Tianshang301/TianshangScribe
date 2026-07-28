# 天殇·书契（TianshangScribe）

跨平台命令行 Office 文档处理工具，支持 Word、Excel、PowerPoint 的创建、编辑、模板填充、格式转换，融入 LaTeX 风格排版标记。

## 安装

```bash
pip install tianshang-scribe

# 含 PDF 转换支持
pip install tianshang-scribe[pdf]

# 开发依赖
pip install tianshang-scribe[dev]
```

## 快速开始

```bash
# 创建空白 Word 文档并添加文本
tianshang-scribe -w --create --add "Hello World" -o hello.docx

# 打开已有文件，替换文本
tianshang-scribe input.docx --replace "旧文本" --replace-new "新文本" -o output.docx

# 样式 + LaTeX 标记
tianshang-scribe -w --create \
  --style "font=Times,size=14,bold" \
  --add "普通文本 \bfseries{加粗} \itshape{斜体}" \
  --latex-style \
  -o styled.docx

# 添加数学公式
tianshang-scribe -w --create \
  --math "E=mc^2" \
  --math "\frac{-b \pm \sqrt{b^2 - 4ac}}{2a}" \
  -o formulas.docx

# 模板填充
tianshang-scribe template.docx --template data.json -o filled.docx

# 转换为 PDF
tianshang-scribe input.docx --topdf -o output.pdf
```

## 全局选项

| 短选项 | 长选项 | 说明 |
|--------|--------|------|
| `-w` | `--word` | 处理 Word 文档 |
| `-e` | `--excel` | 处理 Excel 工作簿 |
| `-p` | `--ppt` | 处理 PowerPoint 演示文稿 |
| `-o` | `--output` | 输出文件路径 |
| `--force` | | 允许覆盖已有文件 |
| `--topdf` | | 输出为 PDF |

若不指定 `-w/-e/-p`，工具根据输入文件扩展名自动推断文档类型。

## 通用操作

| 短选项 | 长选项 | 参数示例 | 说明 |
|--------|--------|----------|------|
| `-cr` | `--create` | `--create` | 创建空白文档（需配合 `-w/-e/-p`） |
| `-a` | `--add` | `--add "内容"` | 添加内容 |
| `-s` | `--style` | `--style "font=Times,size=12,bold"` | 设置默认样式 |
| `-r` | `--replace` | `--replace "旧值" --replace-new "新值"` | 查找替换（加 `--regex` 启用正则） |
| `-d` | `--delete` | `--delete "关键词"` | 删除指定内容 |
| `-m` | `--modify` | `--modify "旧值" --modify-new "新值"` | 修改指定对象 |
| `-t` | `--template` | `--template data.json` | 数据填充模板 `{{placeholder}}` |
| `--meta` | | `--meta title="报告",author="张三"` | 设置文档属性 |

## LaTeX 风格标记

通过 `--latex-style` 启用，可在 `--add` 中嵌入以下标记：

| 标记语法 | 效果 |
|----------|------|
| `\bfseries{text}` | 加粗 |
| `\itshape{text}` | 斜体 |
| `\underline{text}` | 下划线 |
| `\rmfamily{text}` | 衬线字体 |
| `\sffamily{text}` | 无衬线字体 |
| `\ttfamily{text}` | 等宽字体 |
| `\fontfamily{Arial}{text}` | 指定字体 |
| `\fontsize{18}{text}` | 字号（pt） |
| `\color{FF0000}{text}` | 文本颜色（十六进制） |
| `\heading{2}{标题}` | 插入标题 |
| `\newpage` | 分页符 |
| `\includegraphics{path}` | 插入图片 |

支持嵌套：`\bfseries{\itshape{bold italic}}`

## 数学公式（Word）

通过 `--math` 直接插入 LaTeX 数学公式，内部转换为 Word 原生 OMML 数学公式：

```bash
tianshang-scribe -w --create \
  --math "x = \frac{-b \pm \sqrt{b^2 - 4ac}}{2a}" \
  --math "\sum_{i=0}^{n} i^2" \
  --math "\int_{0}^{\infty} e^{-x} dx" \
  -o math.docx
```

支持：上标/下标、分式、根号、求和/积分/连乘、希腊字母、数学符号、重音、绝对值等。

## Word 专属操作

| 选项 | 参数 | 说明 |
|------|------|------|
| `--heading` | `level:1 text:标题` | 添加指定级别的标题 |
| `--math` | `\frac{a}{b}` | 添加 LaTeX 数学公式 |
| `--latex-style` | | 启用 LaTeX 风格标记解析 |
| `--tomd` | | 输出为 Markdown |
| `--tohtml` | | 输出为 HTML |

## Excel 专属操作

| 选项 | 参数 | 说明 |
|------|------|------|
| `--sheet-add` | `SheetName` | 添加工作表 |
| `--sheet-delete` | `SheetName` | 删除工作表 |
| `--sheet-rename` | `Old New` | 重命名工作表 |
| `--column-width` | `col=width` | 设置列宽 |
| `--row-height` | `row=height` | 设置行高 |
| `--formula` | `A1 "=SUM(B1:B10)"` | 设置公式 |
| `--to-csv` | | 导出为 CSV |
| `--to-json` | | 导出为 JSON |
| `--from-csv` | `data.csv` | 从 CSV 导入 |

## PPT 专属操作

| 选项 | 参数 | 说明 |
|------|------|------|
| `--slide-add` | | 添加幻灯片 |
| `--slide-delete` | `3` | 删除指定幻灯片 |
| `--slide-move` | `2 4` | 移动幻灯片位置 |
| `--toimg` | `output_dir/` | 导出为图片序列 |

## 样式语法

`--style` 使用逗号分隔的键值对：

```bash
# 完整语法
--style "font=Times New Roman,size=14,bold,italic,color=FF0000,align=center"

# 布尔简写（出现即 True）
--style "font=Arial,bold,underline"

# 别名
--style "font-family=Times,font-size=18pt"
```

支持的键：

| 键 | 别名 | 类型 | 说明 |
|----|------|------|------|
| `font` | `font_name`, `font-family` | string | 字体名称 |
| `size` | `font_size`, `font-size` | int | 字号（pt） |
| `bold` | | bool | 加粗 |
| `italic` | | bool | 斜体 |
| `underline` | `underlined` | bool | 下划线 |
| `color` | `font_color`, `font-color` | hex | 颜色（`FF0000`） |
| `align` | `alignment` | string | 对齐：`left`/`center`/`right`/`justify` |

## 模板填充

支持 JSON、CSV、YAML 数据源，填充文档中的 `{{placeholder}}`：

```json
{
  "name": "张三",
  "date": "2026-07-28",
  "user": {
    "city": "北京"
  }
}
```

嵌套对象以点号展开：`{{user.city}}` 替换为 `北京`。

## 退出码

| 码 | 含义 |
|----|------|
| `0` | 成功 |
| `1` | 一般错误 |
| `2` | 参数错误 |
| `3` | 功能未实现 |

## 技术栈

- **语言**：Python 3.12（向下兼容 3.10+）
- **CLI**：Typer + Rich
- **Word**：python-docx
- **Excel**：openpyxl
- **PPT**：python-pptx
- **数学公式**：自研 LaTeX → OMML 转换器
- **质量**：pytest（99 测试）、ruff、mypy

## 开发

```bash
git clone https://github.com/yourname/tianshang-scribe.git
cd tianshang-scribe
pip install -e ".[dev]"

# 运行测试
pytest tests/ -v

# 代码检查
ruff check src/ tests/
mypy src/
```
