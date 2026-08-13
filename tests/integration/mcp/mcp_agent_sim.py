"""Agent simulation: test all 7 MCP tools with realistic workflows.

Run:  python tests/integration/mcp/mcp_agent_sim.py
"""

import json
import sys
import tempfile
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent.parent.parent))

from src.core.word_engine import WordEngine
from tests.integration.mcp._mcp_client import MCPClient

OUT = Path(tempfile.mkdtemp(prefix='mcp_agent_test_'))


def main() -> None:
    client = MCPClient()
    tool = client.tool
    call = client.call

    print('=' * 60)
    print('Agent Simulation - TianshangScribe MCP')
    print(f'Output: {OUT}')
    print('=' * 60)

    # 1. Init
    print('\n1. Initialize')
    r = client.initialize()
    name = r['result']['serverInfo']['name']
    print(f'   Server: {name} v{r["result"]["serverInfo"]["version"]}')

    # 2. List tools
    print('\n2. List tools')
    r = call('tools/list')
    for t in r['result']['tools']:
        print(f'   [OK] {t["name"]}')

    # 3. Create Word
    print('\n3. [Agent] "Generate Q3 financial report"')
    result = tool(
        'create_office_document',
        {
            'format': 'docx',
            'style': 'font=Times New Roman,size=12',
            'content': [
                {'type': 'heading', 'text': 'Q3 Financial Report', 'level': 1},
                {'type': 'heading', 'text': 'Revenue', 'level': 2},
                {
                    'type': 'paragraph',
                    'text': r'\bfseries{Total:} \$12.5M. Growth: \itshape{15.3\%}.',
                },
                {'type': 'formula', 'text': r'\sum_{i=1}^{n} x_i = \frac{n(n+1)}{2}'},
                {'type': 'table', 'rows': [['Q1', '$3.5M'], ['Q2', '$4.2M'], ['Q3', '$4.8M']]},
            ],
            'metadata': {'title': 'Q3 Report', 'author': 'Agent'},
            'output_path': str(OUT / 'report.docx'),
        },
    )
    d = result.get('data', {})
    print(f'   [OK] {d.get("paragraphs", 0)} paras, {d.get("sections", 0)} sections')

    # 4. Edit
    print('\n4. [Agent] "Change 15.3% growth to 16%"')
    result = tool(
        'edit_office_document',
        {
            'input_path': str(OUT / 'report.docx'),
            'operations': [{'action': 'replace', 'old_text': '15.3', 'new_text': '16.0'}],
            'output_path': str(OUT / 'report.docx'),
        },
    )
    print(f'   [OK] {result["data"]["total_changes"]} change(s)')

    # 5. Create Excel
    print('\n5. [Agent] "Create sales spreadsheet"')
    result = tool(
        'create_office_document',
        {
            'format': 'xlsx',
            'content': [
                {'type': 'paragraph', 'text': 'Product'},
                {'type': 'paragraph', 'text': 'Widget'},
                {'type': 'paragraph', 'text': 'Gadget'},
                {'type': 'paragraph', 'text': 'Gizmo'},
            ],
            'output_path': str(OUT / 'sales.xlsx'),
        },
    )
    print(f'   [OK] {result["data"].get("sheets", 0)} sheet(s)')

    # 6. Fill template
    print('\n6. [Agent] "Send invitations to 3 people"')
    e = WordEngine()
    e.create()
    e.add_text(
        'Dear {{name}},\n\nYou are invited to {{event}} on {{date}}.\n\nBest regards,\n{{sender}}'
    )
    tpl = str(OUT / 'invite_template.docx')
    e.save(tpl)

    for i, person in enumerate(
        [
            {'name': 'Dr. Smith', 'event': 'AI Summit', 'date': '2026-09-15', 'sender': 'Team'},
            {'name': 'Prof. Lee', 'event': 'ML Workshop', 'date': '2026-10-01', 'sender': 'Team'},
            {'name': 'Ms. Chen', 'event': 'DataConf', 'date': '2026-11-20', 'sender': 'Team'},
        ]
    ):
        result = tool(
            'fill_template',
            {
                'template_path': tpl,
                'data': person,
                'output_path': str(OUT / f'invite_{i}.docx'),
            },
        )
        ok = '[OK]' if result['success'] else '[FAIL]'
        print(f'   {ok} {person["name"]} @ {person["event"]}')

    # 7. Convert PDF
    print('\n7. [Agent] "Convert report to PDF"')
    result = tool(
        'convert_document',
        {
            'input_path': str(OUT / 'report.docx'),
            'target_format': 'pdf',
            'output_path': str(OUT / 'report.pdf'),
        },
    )
    exists = Path(OUT / 'report.pdf').exists()
    print('   [OK] PDF exists' if exists else '   [WARN] Needs LibreOffice')

    # 8. Export JSON
    print('\n8. [Agent] "Export sales as JSON"')
    result = tool(
        'convert_document',
        {
            'input_path': str(OUT / 'sales.xlsx'),
            'target_format': 'json',
            'output_path': str(OUT / 'sales.json'),
        },
    )
    if Path(OUT / 'sales.json').exists():
        data = json.loads(Path(OUT / 'sales.json').read_text(encoding='utf-8'))
        print(f'   [OK] {len(data)} records')

    # 9. Extract
    print('\n9. [Agent] "Show document structure"')
    result = tool(
        'extract_document_data',
        {
            'input_path': str(OUT / 'report.docx'),
            'mode': 'structure',
        },
    )
    print(f'   Paras: {result["data"]["paragraphs"]}')

    print('\n10. [Agent] "Extract report text"')
    result = tool(
        'extract_document_data',
        {
            'input_path': str(OUT / 'report.docx'),
            'mode': 'text',
        },
    )
    print(f'   Blocks: {result["data"]["text_blocks"]}')

    # 11. Dry run
    print('\n11. [Agent] Pre-flight: dry run PPT')
    result = tool(
        'create_office_document',
        {
            'format': 'pptx',
            'content': [{'type': 'heading', 'text': 'Strategy', 'level': 1}],
            'options': {'dry_run': True},
        },
    )
    print(f'   [OK] Planned: {result["data"]["planned_items"]}')

    client.close()

    # Summary
    print('\n' + '=' * 60)
    print(f'All operations complete. Output: {OUT}')
    for f in sorted(OUT.glob('*'), key=lambda x: x.name):
        print(f'  {f.name} ({f.stat().st_size:,} bytes)')
    print('=' * 60)


if __name__ == '__main__':
    main()
