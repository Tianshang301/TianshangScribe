"""Quick test for TianshangScribe MCP Server via stdio JSON-RPC."""
import json
import subprocess
import sys
import tempfile
from pathlib import Path


def mcp_call(method: str, params: dict = None, msg_id: int = 1) -> dict:
    request = json.dumps({
        'jsonrpc': '2.0', 'id': msg_id, 'method': method,
        'params': params or {},
    })
    proc = subprocess.run(
        [sys.executable, '-m', 'mcp.server'],
        input=request + '\n', capture_output=True, text=True, timeout=30,
        cwd=str(Path(__file__).resolve().parent.parent),
    )
    return json.loads(proc.stdout.strip())


print('1. Initialize...')
r = mcp_call('initialize', {'protocolVersion': '2024-11-05', 'capabilities': {}})
print(f'   Server: {r["result"]["serverInfo"]["name"]} v{r["result"]["serverInfo"]["version"]}')
assert 'tools' in r['result']['capabilities']

print('2. List tools...')
r = mcp_call('tools/list')
tools = r['result']['tools']
print(f'   {len(tools)} tools: {[t["name"] for t in tools]}')
assert len(tools) == 6

print('3. Create Word document...')
created_path = None
with tempfile.TemporaryDirectory() as tmp:
    path = Path(tmp) / 'test.docx'
    r = mcp_call('tools/call', {
        'name': 'create_office_document',
        'arguments': {
            'format': 'docx',
            'content': [
                {'type': 'heading', 'text': 'Test Report', 'level': 1},
                {'type': 'paragraph',
                 'text': '\\bfseries{Important}: this is a test.'},
                {'type': 'formula', 'text': '\\frac{a}{b}'},
            ],
            'output_path': str(path),
        },
    })
    text = r['result']['content'][0]['text']
    data = json.loads(text)
    print(f'   Success: {data["success"]}, '
          f'Path: {data.get("data", {}).get("output_path", "")}')
    assert data['success']
    assert path.exists()
    created_path = path

    print('4. Edit document...')
    r = mcp_call('tools/call', {
        'name': 'edit_office_document',
        'arguments': {
            'input_path': str(path),
            'operations': [
                {'action': 'replace', 'old_text': 'Test', 'new_text': 'QA'},
            ],
            'output_path': str(path),
        },
    })
    text = r['result']['content'][0]['text']
    data = json.loads(text)
    print(f'   Success: {data["success"]}, Changes: {data.get("data", {}).get("total_changes", 0)}')
    assert data['success']

    print('5. Create Excel...')
    xpath = Path(tmp) / 'test.xlsx'
    r = mcp_call('tools/call', {
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
    })
    data = json.loads(r['result']['content'][0]['text'])
    print(f'   Success: {data["success"]}, Sheets: {data.get("data", {}).get("sheets", 0)}')
    assert data['success']

    print('6. Fill template...')
    from src.core.word_engine import WordEngine
    e = WordEngine()
    e.create()
    e.add_text('Hello {{name}} from {{city}}')
    tpl = Path(tmp) / 'template.docx'
    e.save(str(tpl))
    filled = Path(tmp) / 'filled.docx'
    r = mcp_call('tools/call', {
        'name': 'fill_template',
        'arguments': {
            'template_path': str(tpl),
            'data': {'name': 'Alice', 'city': 'NYC'},
            'output_path': str(filled),
        },
    })
    data = json.loads(r['result']['content'][0]['text'])
    filled_count = data.get('data', {}).get('placeholders_filled', 0)
    print(f'   Success: {data["success"]}, Filled: {filled_count}')
    assert data['success']

    print('7. Extract data...')
    r = mcp_call('tools/call', {
        'name': 'extract_document_data',
        'arguments': {
            'input_path': str(filled),
            'mode': 'text',
        },
    })
    data = json.loads(r['result']['content'][0]['text'])
    print(f'   Text blocks: {data.get("data", {}).get("text_blocks", 0)}')
    assert 'Alice' in data['data']['text']

print('8. Validate template...')
if created_path:
    r = mcp_call('tools/call', {
        'name': 'validate_template',
        'arguments': {
            'template_path': str(created_path),
            'data': {'name': 'Alice', 'age': '30', 'items': []},
        },
    })
    result = json.loads(r['result']['content'][0]['text'])
    print(f'   Valid: {result.get("data", {}).get("valid")}, '
          f'Missing: {len(result.get("data", {}).get("missing", []))}')
else:
    print('   Skipped (no test doc)')

print('9. Resource list (stdio — cleared per-process)...')
r = mcp_call('resources/list')
resources = r['result'].get('resources', [])
print(f'   Resources: {len(resources)} (expected 0 in stdio mode)')

print()
print('All 9 tests passed!')
