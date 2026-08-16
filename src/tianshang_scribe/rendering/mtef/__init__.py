"""MathType compatibility: OLE/MTEF <-> LaTeX <-> OMML.

Reads MathType equation objects embedded in Word documents
(``word/embeddings/oleObject*.bin``) and converts them to native OMML via a
MTEF binary reader and the existing LaTeX->OMML engine.  The reverse
direction (``mtef_writer`` + ``cfb_writer``) produces MTEF/OLE payloads so a
formula can be embedded back into a document as a MathType object.
"""

from __future__ import annotations

from .cfb_writer import make_ole
from .mtef_reader import MTEFParseError, mtef_to_latex
from .mtef_writer import latex_to_mtef
from .ole_util import extract_native_stream, open_ole

__all__ = [
    'MTEFParseError',
    'extract_native_stream',
    'latex_to_mtef',
    'make_ole',
    'mtef_to_latex',
    'open_ole',
]
