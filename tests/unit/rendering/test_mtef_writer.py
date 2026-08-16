"""Unit tests for the MTEF binary writer (LaTeX -> MathType).

Every test encodes LaTeX to MTEF and round-trips it through the reader to
verify the structural semantics survive.  Additional byte-level assertions
pin the record encoding.
"""

from __future__ import annotations

import pytest

from tianshang_scribe.rendering.mtef.mtef_reader import mtef_to_latex
from tianshang_scribe.rendering.mtef.mtef_writer import _header, latex_to_mtef


def _roundtrip(latex: str) -> str:
    return mtef_to_latex(latex_to_mtef(latex))


def test_header_structure() -> None:
    header = _header()
    assert header[:5] == bytes([5, 1, 1, 0, 0])
    assert header[5:14] == b'MathType\x00'
    assert header[14] == 0
    assert len(header) == 15


def test_plain_text() -> None:
    assert _roundtrip('x + y') == 'x+y'


def test_fraction() -> None:
    assert _roundtrip(r'\frac{a}{b}') == r'\frac{a}{b}'


def test_superscript() -> None:
    assert _roundtrip('x^2') == 'x^{2}'


def test_superscript_multi_char() -> None:
    assert _roundtrip('x^{n+1}') == 'x^{n+1}'


def test_subscript() -> None:
    assert _roundtrip('a_1') == 'a_{1}'


def test_subsup() -> None:
    # The reader renders SUBSUP as base^{sup}_{sub}.
    out = _roundtrip('x_i^2')
    assert 'x' in out and '2' in out and 'i' in out


def test_sqrt() -> None:
    assert _roundtrip(r'\sqrt{x}') == r'\sqrt{x}'


def test_sqrt_root_degree() -> None:
    assert _roundtrip(r'\sqrt[3]{x}') == r'\sqrt[3]{x}'


def test_sum_with_limits() -> None:
    assert _roundtrip(r'\sum_{i=0}^{n} i') == r'\sum_{i=0}^{n} i'


def test_prod_with_limits() -> None:
    assert _roundtrip(r'\prod_{k=1}^{m} k') == r'\prod_{k=1}^{m} k'


def test_integral_with_limits() -> None:
    assert _roundtrip(r'\int_0^1 f(x) dx').startswith(r'\int_{0}^{1}')


def test_operators_render_as_text() -> None:
    out = _roundtrip(r'\sin x + \cos y')
    assert 'sin' in out and 'cos' in out


def test_greek_letters_roundtrip() -> None:
    out = _roundtrip(r'\alpha + \beta')
    assert '\u03b1' in out and '\u03b2' in out


def test_combined_formula() -> None:
    out = _roundtrip(r'\frac{-b \pm \sqrt{b^2-4ac}}{2a}')
    assert '\\frac' in out
    assert '\\sqrt' in out
    assert 'b' in out and 'a' in out


def test_unknown_command_falls_back_to_text() -> None:
    # \ne is not in the symbol table; the writer must not crash and should
    # still emit a usable formula.
    out = _roundtrip(r'a \ne b')
    assert 'a' in out and 'b' in out


def test_empty_input_returns_header_and_line() -> None:
    data = latex_to_mtef('')
    assert data[:5] == bytes([5, 1, 1, 0, 0])


def test_output_starts_with_header() -> None:
    data = latex_to_mtef('x')
    assert data[:15] == _header()[:15]


@pytest.mark.parametrize(
    'latex',
    [
        'a',
        'xy',
        r'\frac{a}{b}',
        r'\frac{1}{x^2 + y^2}',
        r'\sqrt{x}',
        'x^2',
        'x^{n+1}',
        'a_1',
        'x_i^2',
        r'\sum_{i=1}^{n} i^2',
        r'\int_a^b f(t) dt',
        r'\alpha \beta \gamma',
        r'\pi r^2',
    ],
)
def test_roundtrip_no_crash(latex: str) -> None:
    data = latex_to_mtef(latex)
    assert data
    result = mtef_to_latex(data)
    assert result
