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
        
        # 获取原 Run 的样式
        # 注意：find 仍然需要 qn
        orig_rPr = target_run._r.find(qn("w:rPr"))

        # --- 第二段：创建上标引用 r_sup ---
        # 修复点：直接使用 "w:r"，不要用 qn
        r_sup = OxmlElement("w:r") 
        
        # 修复点：直接使用 "w:rPr"
        new_rPr = copy.deepcopy(orig_rPr) if orig_rPr is not None else OxmlElement("w:rPr")
        
        # 修复点：直接使用 "w:vertAlign"
        va = OxmlElement("w:vertAlign")
        # 注意：.set() 方法必须使用 qn
        va.set(qn("w:val"), "superscript")
        new_rPr.append(va)
        r_sup.append(new_rPr)
        
        # 修复点：直接使用 "w:t"
        t_sup = OxmlElement("w:t")
        t_sup.text = f"[{ref_id}]"
        r_sup.append(t_sup)

        # --- 第三段：创建 r_tail (插入点之后的文字) ---
        r_tail = OxmlElement("w:r")
        if orig_rPr is not None:
            r_tail.append(copy.deepcopy(orig_rPr))
            
        t_tail = OxmlElement("w:t")
        # 这里可以使用 qn("xml:space") 或者保留你原来的写法
        t_tail.set(qn("xml:space"), "preserve")
        t_tail.text = orig_text[target_off:]
        r_tail.append(t_tail)

        # --- 最后：挂载 ---
        target_run._r.addnext(r_tail)
        target_run._r.addnext(r_sup)
        for run in para.runs:
            if run.text and f"[{ref_id}]" in run.text:
                run.text = run.text.replace(f"[{ref_id}]", "")
        return True
    return False
