from docx import Document
from docx.oxml import OxmlElement
from docx.oxml.ns import qn

from .constants import W


def _insert_section_break(doc: Document) -> None:
    para = doc.add_paragraph()
    pPr  = para._p.get_or_add_pPr()
    sect = OxmlElement("w:sectPr")
    pgBrk = OxmlElement("w:type")
    pgBrk.set(qn("w:val"), "nextPage")
    sect.append(pgBrk)
    pPr.append(sect)
