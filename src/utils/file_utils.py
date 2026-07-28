from __future__ import annotations

from pathlib import Path


def resolve_glob(pattern: str) -> list[Path]:
    base = Path('.')
    return sorted(base.glob(pattern))


def safe_read(path: str | Path) -> bytes:
    with open(path, 'rb') as f:
        return f.read()


def safe_write(path: str | Path, data: bytes) -> None:
    Path(path).parent.mkdir(parents=True, exist_ok=True)
    with open(path, 'wb') as f:
        f.write(data)


def ensure_directory(path: str | Path) -> Path:
    p = Path(path)
    p.mkdir(parents=True, exist_ok=True)
    return p


def get_output_name(input_path: str | Path, suffix: str = '-out') -> str:
    p = Path(input_path)
    return str(p.parent / f'{p.stem}{suffix}{p.suffix}')


def parse_key_value(text: str) -> dict[str, str]:
    result: dict[str, str] = {}
    for pair in text.split(','):
        pair = pair.strip()
        if '=' in pair:
            key, _, value = pair.partition('=')
            result[key.strip()] = value.strip()
    return result


def is_binary_content(data: bytes) -> bool:
    return b'\x00' in data[:1024]


EXIT_SUCCESS = 0
EXIT_ERROR = 1
EXIT_INVALID_ARGS = 2
EXIT_CORRUPT = 3
