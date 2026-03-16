from docx.oxml import OxmlElement
from docx.oxml.ns import qn
import copy

from .constants import W


def insert_citation(doc, ref_id, context):
    before, after = context
    for para in doc.paragraphs:
        full = para.text
        pos  = full.find(before)
        if pos == -1: continue
        insert_at = pos + len(before)
        if after and after[:5] not in full[insert_at:]: continue
        runs = para.runs
        if not runs: continue
        cur = 0
        target_idx, target_off = len(runs)-1, len(runs[-1].text)
        for ri, run in enumerate(runs):
            end = cur + len(run.text)
            if cur <= insert_at <= end:
                target_idx = ri; target_off = insert_at - cur; break
            cur = end
        target_run = runs[target_idx]
        orig_text  = target_run.text
        target_run.text = orig_text[:target_off]
        r_sup    = OxmlElement("w:r")
        orig_rPr = target_run._r.find(f"{{{W}}}rPr")
        new_rPr  = copy.deepcopy(orig_rPr) if orig_rPr is not None \
                   else OxmlElement("w:rPr")
        va = OxmlElement("w:vertAlign")
        va.set(qn("w:val"), "superscript")
        new_rPr.append(va)
        r_sup.append(new_rPr)
        t_sup = OxmlElement("w:t")
        t_sup.text = f"[{ref_id}]"
        r_sup.append(t_sup)
        r_tail = OxmlElement("w:r")
        if orig_rPr is not None:
            r_tail.append(copy.deepcopy(orig_rPr))
        t_tail = OxmlElement("w:t")
        t_tail.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
        t_tail.text = orig_text[target_off:]
        r_tail.append(t_tail)
        target_run._r.addnext(r_tail)
        target_run._r.addnext(r_sup)
        for run in para.runs:
            if run.text and f"[{ref_id}]" in run.text:
                run.text = run.text.replace(f"[{ref_id}]", "")
        return True
    return False
