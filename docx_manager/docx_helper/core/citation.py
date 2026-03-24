from docx.oxml import OxmlElement
from docx.oxml.ns import qn
import copy

from .constants import W


def insert_citation(doc, ref_id, context):
    before, after = context
    old_ref_text = f"[{ref_id}]"
    
    for para in doc.paragraphs:
        full = para.text
        # 1. 定位插入点
        pos = full.find(before)
        if pos == -1: 
            continue
            
        insert_at = pos + len(before)
        
        # 2. 上下文验证（双向锚定）
        if after and after[:5] not in full[insert_at:]: 
            continue
            
        runs = para.runs
        if not runs: 
            continue
            
        # 3. 寻找插入点落在哪个 Run 
        cur = 0
        target_idx, target_off = len(runs)-1, len(runs[-1].text)
        for ri, run in enumerate(runs):
            end = cur + len(run.text)
            if cur <= insert_at <= end:
                target_idx = ri
                target_off = insert_at - cur
                break
            cur = end
            
        # 4. 处理 target_run (前半部分文字)
        target_run = runs[target_idx]
        orig_text = target_run.text
        target_run.text = orig_text[:target_off]
        
        # 获取原样式
        orig_rPr = target_run._r.find(qn("w:rPr"))

        # 5. 创建上标引用 r_sup
        r_sup = OxmlElement("w:r") 
        new_rPr = copy.deepcopy(orig_rPr) if orig_rPr is not None else OxmlElement("w:rPr")
        
        # 设置上标属性
        va = OxmlElement("w:vertAlign")
        va.set(qn("w:val"), "superscript")
        new_rPr.append(va)
        r_sup.append(new_rPr)
        
        # 设置引用文本
        t_sup = OxmlElement("w:t")
        t_sup.text = old_ref_text
        r_sup.append(t_sup)

        # 6. 创建 r_tail (后半部分文字)
        r_tail = OxmlElement("w:r")
        if orig_rPr is not None:
            r_tail.append(copy.deepcopy(orig_rPr))
            
        t_tail = OxmlElement("w:t")
        t_tail.set(qn("xml:space"), "preserve")
        
        # --- 核心改进：跳过原有的非上标引用文本 ---
        remaining_text = orig_text[target_off:]
        if remaining_text.startswith(old_ref_text):
            # 如果后面紧跟着旧引用，则跳过它的长度
            t_tail.text = remaining_text[len(old_ref_text):]
        else:
            t_tail.text = remaining_text
            
        r_tail.append(t_tail)

        # 7. 链表式挂载节点到 XML
        # 顺序：target_run -> r_sup -> r_tail
        target_run._r.addnext(r_tail)
        target_run._r.addnext(r_sup)
        
        # 注意：不再使用全局 replace 循环，防止误伤新插入的 r_sup
        
        return True
        
    return False
