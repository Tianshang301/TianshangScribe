# TianshangScribe

> [English](../README.md)

跨平台命令行 Office 文档处理工具。支持 Word、Excel、PowerPoint 的创建、编辑、模板填充、格式转换，内置 LaTeX 风格排版标记与数学公式渲染引擎。

## 安装

```bash
git clone https://github.com/Tianshang301/TianshangScribe.git
cd TianshangScribe
pip install -e ".[dev]"
```

依赖：Python 3.10+ · python-docx · openpyxl · python-pptx · typer · rich · jinja2 · lxml

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

# 转换为 PDF
tianshang-scribe input.docx --topdf -o output.pdf

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
| `-r` `--replace` | 查找替换 | `-r "foo" --replace-new "bar"` |
| `-d` `--delete` | 删除内容 | `-d "关键词"` |
| `-m` `--modify` | 修改内容 | `-m "旧值" --modify-new "新值"` |
| `-s` `--style` | 设置样式 | `-s "font=Times,size=14,bold"` |
| `-t` `--template` | 模板填充 | `-t data.json` |
| `-x` `--extract` | 提取数据 | `-x metadata` |
| `--meta` | 设置属性 | `--meta "title=报告,author=张三"` |
| `--latex-style` | 启用 LaTeX 标记解析 | |
| `--math` | 添加数学公式（Word） | `--math "\frac{a}{b}"` |
| `--heading` | 添加标题（Word） | `--heading "level:1 text:引言"` |
| `--regex` | 正则模式 | 配合 `--replace` `--delete` 使用 |
| `--merge` | 合并文件 | `--merge "a.docx,b.docx"` |

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
| `--clear` | 清除内容 | `--clear` |
| `--to-csv` | 导出 CSV | |
| `--to-json` | 导出 JSON | |
| `--to-html` | 导出 HTML | |

## LaTeX 样式标记

在 `--add` 中嵌入下列标记，启用 `--latex-style` 后自动解析。支持嵌套。

| 标记 | 效果 |
|------|------|
| `\bfseries{text}` | 加粗 |
| `\itshape{text}` | 斜体 |
| `\underline{text}` | 下划线 |
| `\rmfamily{text}` | 衬线（Roman） |
| `\sffamily{text}` | 无衬线（Sans-serif） |
| `\ttfamily{text}` | 等宽（Monospace） |
| `\fontfamily{Arial}{text}` | 指定字体 |
| `\fontsize{18}{text}` | 字号（pt） |
| `\color{FF0000}{text}` | 颜色（十六进制） |
| `\centering{...}` | 居中 |
| `\raggedright{...}` | 左对齐 |
| `\raggedleft{...}` | 右对齐 |
| `\heading{2}{标题}` | 插入标题 |
| `\newpage` | 分页符 |
| `\includegraphics{path}` | 插入图片 |

### 字体配置

| 命令 | 效果 |
|------|------|
| `\setmainfont{Name}` | 西文默认字体 |
| `\setCJKmainfont{Name}` | CJK 默认字体 |
| `\setsansfont{Name}` | 无衬线字体 |
| `\setmonofont{Name}` | 等宽字体 |

Word OOXML 原生分离 `w:ascii`（西文）与 `w:eastAsia`（CJK）字体，中西文混排自动切换。

## 数学公式

通过 `--math` 支持完整的 LaTeX 数学公式，内部转换为 Word 原生 OMML。

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
| 重音 | `\hat{x}` `\bar{x}` `\tilde{x}` `\dot{x}` `\ddot{x}` `\vec{x}` `\widehat{x}` … |
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
| `font` | `font_name` | 字体名 | 西文字体 |
| `cjk-font` | `cjk_font_name` | 字体名 | CJK 字体 |
| `size` | `font_size` | pt 值 | 字号 |
| `bold` | | 标志 | 加粗 |
| `italic` | | 标志 | 斜体 |
| `underline` | | 标志 | 下划线 |
| `color` | `font_color` | `FF0000` | 十六进制颜色 |
| `align` | `alignment` | `left`/`center`/`right`/`justify` | 对齐 |

布尔键（`bold` `italic` `underline`）出现即为 `True`。

## 模板填充

支持 JSON、CSV、YAML 数据源，填充 `{{placeholder}}` 占位符。嵌套对象以点号展开。

```json
{
  "name": "张三",
  "date": "2026-07-28",
  "user": { "city": "北京" }
}
```

```
{{name}}         →  张三
{{user.city}}    →  北京
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
| 清除内容 | `--clear` |

## PPT 功能

| 功能 | 说明 |
|------|------|
| 幻灯片管理 | 增删移幻灯片 |
| 版式 | 应用幻灯片版式 |
| 演讲者备注 | 添加备注 |
| 导出 | 另存为图片序列（`--toimg`） |

## 退出码

| 码 | 含义 |
|----|------|
| `0` | 成功 |
| `1` | 一般错误 |
| `2` | 参数错误 |
| `3` | 功能未实现 |

## 架构

```
src/
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
│   └── pdf.py         # PDF 导出（LibreOffice）
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
| 数学公式 | 自研递归下降解析器 → OMML XML |
| 模板 | Jinja2 + docxtpl |
| PDF | LibreOffice headless |
| 质量 | pytest（147 用例）· ruff · mypy |

## 开发

```bash
git clone https://github.com/Tianshang301/TianshangScribe.git
cd TianshangScribe
pip install -e ".[dev]"

pytest tests/ -v      # 运行测试
ruff check src/ tests/ # 代码检查
mypy src/             # 类型检查
```

## 许可

Apache-2.0
