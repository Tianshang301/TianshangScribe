"""Quick test for TianshangScribe MCP Server via stdio JSON-RPC.

Run:  python tests/integration/mcp/mcp_stdio_smoke.py
"""

import json
import sys
import tempfile
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent.parent.parent))

from tests.integration.mcp._mcp_client import MCPClient


def main() -> None:
    client = MCPClient()

    print('1. Initialize...')
    r = client.initialize()
    print(f'   Server: {r["result"]["serverInfo"]["name"]} v{r["result"]["serverInfo"]["version"]}')
    assert 'tools' in r['result']['capabilities']

    print('2. List tools...')
    r = client.call('tools/list')
    tools = r['result']['tools']
    print(f'   {len(tools)} tools: {[t["name"] for t in tools]}')
    assert len(tools) == 12

    print('3. Create Word document...')
    with tempfile.TemporaryDirectory() as tmp:
        path = Path(tmp) / 'test.docx'
        r = client.call(
            'tools/call',
            {
                'name': 'create_office_document',
                'arguments': {
                    'format': 'docx',
                    'content': [
                        {'type': 'heading', 'text': 'Test Report', 'level': 1},
                        {'type': 'paragraph', 'text': '\\bfseries{Important}: this is a test.'},
                        {'type': 'formula', 'text': '\\frac{a}{b}'},
                    ],
                    'output_path': str(path),
                },
            },
        )
        text = r['result']['content'][0]['text']
        data = json.loads(text)
        print(f'   Success: {data["success"]}, Path: {data.get("data", {}).get("output_path", "")}')
        assert data['success']
        assert path.exists()

        print('4. Edit document...')
        r = client.call(
            'tools/call',
            {
                'name': 'edit_office_document',
                'arguments': {
                    'input_path': str(path),
                    'operations': [{'action': 'replace', 'old_text': 'Test', 'new_text': 'QA'}],
                    'output_path': str(path),
                },
            },
        )
        text = r['result']['content'][0]['text']
        data = json.loads(text)
        print(
            f'   Success: {data["success"]}, Changes: {data.get("data", {}).get("total_changes", 0)}'
        )
        assert data['success']

        print('5. Create Excel...')
        xpath = Path(tmp) / 'test.xlsx'
        r = client.call(
            'tools/call',
            {
                'name': 'create_office_document',
                'arguments': {
                    'format': 'xlsx',
                    'content': [
                        {'type': 'paragraph', 'text': 'Product'},
                        {'type': 'paragraph', 'text': 'Widget'},
                        {'type': 'paragraph', 'text': 'Gadget'},
                    ],
                    'output_path': str(xpath),
                },
            },
        )
        data = json.loads(r['result']['content'][0]['text'])
        print(f'   Success: {data["success"]}, Sheets: {data.get("data", {}).get("sheets", 0)}')
        assert data['success']

        print('6. Fill template...')
        from tianshang_scribe.core.word_engine import WordEngine

        e = WordEngine()
        e.create()
        e.add_text('Hello {{name}} from {{city}}')
        tpl = Path(tmp) / 'template.docx'
        e.save(str(tpl))
        filled = Path(tmp) / 'filled.docx'
        r = client.call(
            'tools/call',
            {
                'name': 'fill_template',
                'arguments': {
                    'template_path': str(tpl),
                    'data': {'name': 'Alice', 'city': 'NYC'},
                    'output_path': str(filled),
                },
            },
        )
        data = json.loads(r['result']['content'][0]['text'])
        filled_count = data.get('data', {}).get('placeholders_filled', 0)
        print(f'   Success: {data["success"]}, Filled: {filled_count}')
        assert data['success']

        print('7. Extract data...')
        r = client.call(
            'tools/call',
            {
                'name': 'extract_document_data',
                'arguments': {
                    'input_path': str(filled),
                    'mode': 'text',
                },
            },
        )
        data = json.loads(r['result']['content'][0]['text'])
        print(f'   Text blocks: {data.get("data", {}).get("text_blocks", 0)}')
        assert 'Alice' in data['data']['text']

        print('8. Validate template...')
        r = client.call(
            'tools/call',
            {
                'name': 'validate_template',
                'arguments': {
                    'template_path': str(path),
                    'data': {'name': 'Alice', 'age': '30', 'items': []},
                },
            },
        )
        result = json.loads(r['result']['content'][0]['text'])
        print(
            f'   Valid: {result.get("data", {}).get("valid")}, '
            f'Missing: {len(result.get("data", {}).get("missing", []))}'
        )

    print('9. Resource list (no resources registered)...')
    r = client.call('resources/list')
    resources = r['result'].get('resources', [])
    print(f'   Resources: {len(resources)} (expected 0)')

    client.close()
    print()
    print('All 9 tests passed!')


if __name__ == '__main__':
    main()
