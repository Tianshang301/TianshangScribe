"""TianshangScribe — Comprehensive Demo Generator

Generates three output files exercising all major features:
  - demo_word.docx   : LaTeX markup, math formulas, TOC, header/footer, watermark
  - demo_excel.xlsx  : CSV import, formula, sort, chart, comment, protection
  - demo_ppt.pptx    : slides, layouts, notes, transitions
"""
from __future__ import annotations

import tempfile
from pathlib import Path

from src.core.word_engine import WordEngine
from src.core.excel_engine import ExcelEngine
from src.core.ppt_engine import PptEngine


def build_word_demo(out_dir: str) -> None:
    """Word: LaTeX styles + math formulas + TOC + header/footer + watermark."""
    print('--- Building Word demo ---')
    engine = WordEngine()
    engine.create()
    engine.set_style('font=Times New Roman,size=12')
    engine.add_latex_content(r'\setCJKmainfont{SimSun}')

    # Title
    engine.add_latex_content(r'\heading{1}{TianshangScribe — Word 综合演示}')
    engine.add_latex_content(r'\centering{TianshangScribe Comprehensive Demo}')
    engine.add_latex_content(r'\newpage')
    engine.add_latex_content(r'\heading{2}{目录}')
    engine.add_toc()
    engine.add_latex_content(r'\newpage')

    # Chapter 1: Text styles
    engine.add_latex_content(r'\heading{2}{文本样式 Text Styles}')
    engine.add_latex_content(
        r'\bfseries{加粗 Bold} and \itshape{斜体 Italic} '
        r'and \underline{下划线 Underline}'
    )
    engine.add_latex_content(
        r'\scshape{Small Caps Text} and '
        r'\fontfamily{Courier New}{Monospace} and '
        r'\fontsize{18}{Large Size}'
    )
    engine.add_latex_content(
        r'\color{FF0000}{Red Text} and \color{0000FF}{Blue Text}'
    )
    engine.add_latex_content(r'\centering{居中段落 Centered Paragraph}')
    engine.add_latex_content(r'\raggedright{左对齐段落 Left-Aligned Paragraph}')
    engine.add_latex_content(r'\raggedleft{右对齐段落 Right-Aligned Paragraph}')
    engine.add_latex_content(r'\linespread{2.0}{行距 2.0 倍 Line Spacing 2x}')
    engine.add_latex_content(r'\indent{首行缩进段落 First-line Indented}')

    # Chapter 2: Math formulas
    engine.add_latex_content(r'\newpage')
    engine.add_latex_content(r'\heading{2}{数学公式 Math Formulas}')

    engine.add_latex_content(r'\heading{3}{基础公式 Basic Formulas}')
    engine.add_text('行内公式 Inline: $E = mc^2$')
    engine.add_text('行间公式 Display:')
    engine.add_latex_content(r'$$x = \frac{-b \pm \sqrt{b^2 - 4ac}}{2a}$$')
    engine.add_latex_content(r'$$\sum_{n=1}^{\infty} \frac{1}{n^2} = \frac{\pi^2}{6}$$')
    engine.add_latex_content(
        r'$$\int_{0}^{\infty} e^{-x^2} dx = \frac{\sqrt{\pi}}{2}$$'
    )
    engine.add_latex_content(
        r'$$\lim_{x \to 0} \frac{\sin x}{x} = 1$$'
    )

    engine.add_latex_content(r'\heading{3}{高级公式 Advanced Formulas}')
    engine.add_latex_content(
        r'$$f(x) = \int_{-\infty}^{\infty} \hat{f}(\xi) e^{2\pi i \xi x} d\xi$$'
    )
    engine.add_latex_content(
        r'$$\frac{\partial^2 u}{\partial t^2} = c^2 \nabla^2 u$$'
    )
    engine.add_latex_content(
        r'$$\nabla \cdot \mathbf{E} = \frac{\rho}{\epsilon_0},'
        r'\quad \nabla \times \mathbf{B} = \mu_0 \mathbf{J}$$'
    )

    engine.add_latex_content(r'\heading{3}{数学字体 Math Font Styles}')
    engine.add_latex_content(r'$$\mathrm{normal} \quad \mathbf{bold} \quad '
                             r'\mathit{italic} \quad \mathsf{sans} '
                             r'\quad \mathtt{mono} \quad \mathbb{ABC}$$')
    engine.add_latex_content(
        r'$$\mathcal{A} \subseteq \mathcal{B} \implies '
        r'\mathcal{P}(\mathcal{A}) \subseteq \mathcal{P}(\mathcal{B})$$'
    )

    engine.add_latex_content(r'\heading{3}{符号表 Symbol Table}')
    engine.add_latex_content(
        r'$\alpha \beta \gamma \delta \epsilon \zeta \eta \theta$'
    )
    engine.add_latex_content(
        r'$\Gamma \Delta \Theta \Lambda \Xi \Pi \Sigma \Phi \Psi \Omega$'
    )
    engine.add_latex_content(
        r'$\forall x \in \mathbb{R}, \exists N \in \mathbb{N}: '
        r'x \leq N \lor x \geq -N$'
    )

    # Chapter 3: Mixed content
    engine.add_latex_content(r'\newpage')
    engine.add_latex_content(r'\heading{2}{混合内容 Mixed Content}')
    engine.add_text('')
    engine.add_text(
        '自动识别数学命令：The limit \\lim_{x \\to 0} \\frac{\\sin x}{x} = 1 '
        'is fundamental. No $ needed!'
    )
    engine.add_text('')
    engine.add_text(
        'Unicode 上下标：H₂O (H_{2}O) and area m² (m^{2}). '
        'Linear algebra: x₁ + x₂ = x₃.'
    )

    # Nested styles
    engine.add_latex_content(
        r'\bfseries{\itshape{粗体斜体嵌套 Bold-Italic Nested}}'
    )
    engine.add_latex_content(
        r'\fontfamily{Arial}{\color{FF6600}{'
        r'Arial Orange with \bfseries{Bold inside}}}'
    )

    # Header / Footer / Watermark
    engine.set_header('TianshangScribe 演示文档')
    engine.set_footer('第 PAGE 页 — 天殇·书契')
    engine.add_watermark('天殇·书契')

    path = str(Path(out_dir) / 'demo_word.docx')
    engine.save(path)
    print(f'  Saved: {path}')


def build_excel_demo(out_dir: str) -> None:
    """Excel: CSV import + formula + sort + chart + comment + protection."""
    print('--- Building Excel demo ---')
    engine = ExcelEngine()
    engine.create()
    engine.set_style('font=Calibri,size=11')

    # Sheet 1: Sales data
    engine.add_sheet('Sales')
    # Delete default sheet
    if 'Sheet' in engine.wb.sheetnames:
        del engine.wb['Sheet']

    engine.import_csv(str(Path('demo') / 'test_data.csv'))
    engine.set_column_width(1, 18)
    engine.set_column_width(2, 12)
    engine.set_column_width(3, 12)
    engine.set_column_width(4, 12)
    engine.set_column_width(5, 14)

    # Add Total column
    engine.add_text('Total', column=6)
    engine.set_formula('F2', '=C2*D2')
    # Copy formula down manually
    ws = engine.wb.active
    for row in range(3, 12):
        ws[f'F{row}'] = f'=C{row}*D{row}'

    # Add summary row
    engine.add_text('', column=1)
    engine.add_text('SUM:', column=5)
    engine.set_formula('C12', '=SUM(C2:C11)')
    engine.set_formula('D12', '=SUM(D2:D11)')
    engine.set_formula('F12', '=SUM(F2:F11)')

    # Add comments
    engine.add_comment('A1', 'Product name')
    engine.add_comment('C1', 'Unit price in USD')
    engine.add_comment('F1', 'Total = Price x Quantity')

    # Sheet 2: Sorted data
    engine.add_sheet('Sorted by Price')
    engine.import_csv(str(Path('demo') / 'test_data.csv'))
    engine.sort('C2:C11', 'desc')

    # Sheet 3: Chart
    engine.add_sheet('Chart')
    engine.add_text('Product', column=1)
    engine.add_text('Price', column=2)
    products = ['Widget', 'Gadget', 'Gizmo', 'Thingamajig', 'Doohickey']
    prices = [19.99, 34.50, 12.75, 49.99, 27.30]
    for p, pr in zip(products, prices):
        engine.add_text(p, column=1)
        engine.add_text(str(pr), column=2)
    ws = engine.wb.active
    engine.add_chart('bar', f"'{ws.title}'!A1:B6")

    # Protection
    engine.set_protection('demo123')

    path = str(Path(out_dir) / 'demo_excel.xlsx')
    engine.save(path)
    print(f'  Saved: {path}')


def build_ppt_demo(out_dir: str) -> None:
    """PPT: slides + layouts + notes + transitions."""
    print('--- Building PPT demo ---')
    engine = PptEngine()
    engine.create()
    engine.set_style('font=Calibri,size=18')

    # Slide 0: Title
    engine.add_slide()
    engine.add_latex_content(r'\heading{TianshangScribe}')
    engine.add_notes(0, 'Welcome slide. Introduce the tool.')

    # Slide 1: Overview
    engine.add_slide(layout_index=1)
    engine.add_latex_content(r'\heading{功能概况 Feature Overview}')
    engine.add_latex_content(
        r'\bfseries{Word} — LaTeX 排版 + OMML 数学公式 + 模板填充'
    )
    engine.add_slide()
    engine.add_latex_content(
        r'\bfseries{Excel} — CSV 导入导出 + 公式 + 排序 + 图表'
    )
    engine.add_slide()
    engine.add_latex_content(
        r'\bfseries{PPT} — 幻灯片管理 + 切换效果 + 演讲者备注'
    )
    engine.add_notes(1, 'Overview of all three document types supported.')

    # Slide 4: Word features
    engine.add_slide()
    engine.add_latex_content(r'\heading{Word 功能}')
    engine.add_latex_content(
        r'\sffamily{'
        r'LaTeX 样式标记（加粗、斜体、颜色、字号）\newpage'
        r'数学公式（分式、根号、积分、求和、希腊字母）\newpage'
        r'目录生成、页眉页脚、水印'
        r'}'
    )
    engine.add_notes(4, 'Word features include LaTeX styling and native math formulas.')

    # Slide 5: Excel features
    engine.add_slide()
    engine.add_latex_content(r'\heading{Excel 功能}')
    engine.add_latex_content(
        r'\sffamily{'
        r'CSV 导入导出\newpage'
        r'公式写入与排序\newpage'
        r'图表生成（bar/line/pie）\newpage'
        r'密码保护与批注'
        r'}'
    )
    engine.add_notes(5, 'Excel features include data import, formulas, sorting, and charts.')

    # Slide 6: Math showcase
    engine.add_slide()
    engine.add_latex_content(r'\heading{数学公式示例 Math Examples}')
    engine.add_text(
        r'The quadratic formula is $$x = \frac{-b \pm \sqrt{b^2 - 4ac}}{2a}$$'
    )
    engine.add_notes(6, 'Math formulas rendered as native OMML in PowerPoint.')

    # Transitions
    engine.set_transition('fade')

    path = str(Path(out_dir) / 'demo_ppt.pptx')
    engine.save(path)
    print(f'  Saved: {path}')


def main() -> None:
    out_dir = Path('demo')
    out_dir.mkdir(exist_ok=True)

    build_word_demo(str(out_dir))
    build_excel_demo(str(out_dir))
    build_ppt_demo(str(out_dir))

    print('\n=== All demos generated ===')
    print(f'  Word:  demo/demo_word.docx')
    print(f'  Excel: demo/demo_excel.xlsx')
    print(f'  PPT:   demo/demo_ppt.pptx')


if __name__ == '__main__':
    main()
