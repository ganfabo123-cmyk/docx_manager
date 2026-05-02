import re

_MATH_SYMBOLS = set('∑∫∏√±×÷≤≥≠→←⇒∈∉∀∃∂∇∞')
_GREEK_LETTERS = set('αβγδεζηθικλμνξπρστυφχψωΑΒΓΔΕΖΗΘΙΚΛΜΝΞΠΡΣΤΥΦΧΨΩ')
_FORMULA_LABEL_RE = re.compile(r'\(\d+-\d+\)\s*$')


def _is_suspected_formula(content: str) -> bool:
    if '\\' in content:
        return True
    if '^' in content or '_' in content:
        return True
    if any(c in _MATH_SYMBOLS for c in content):
        return True
    if any(c in _GREEK_LETTERS for c in content):
        return True
    if _FORMULA_LABEL_RE.search(content):
        return True
    return False


def _is_suspected_table(content: str) -> bool:
    if '\t' in content or '   ' in content:
        return True
    if '|' in content:
        return True
    lines = [l for l in content.split('\n') if l.strip()]
    if len(lines) > 1:
        counts = [len(l.split()) for l in lines]
        if len(set(counts)) == 1 and counts[0] > 1:
            return True
    return False


def detect(blocks: list[dict]) -> dict:
    result = {"formula": [], "table": [], "image": []}
    for block in blocks:
        btype = block.get("type", "")
        if btype == "formula":
            result["formula"].append(block)
        elif btype == "table":
            result["table"].append(block)
        elif btype == "image":
            result["image"].append(block)
        elif btype == "body":
            content = block.get("content", "")
            if _is_suspected_formula(content):
                result["formula"].append(block)
            elif _is_suspected_table(content):
                result["table"].append(block)
    return result


def confirm(detected: dict) -> dict:
    from llm_router import route_confirm

    result = {
        "formula": [],
        "table":   [],
        "image":   list(detected.get("image", [])),
    }
    suspected = {"formula": [], "table": []}

    for block in detected.get("formula", []):
        if block.get("type") == "body":
            suspected["formula"].append(block)
        else:
            result["formula"].append(block)

    for block in detected.get("table", []):
        if block.get("type") == "body":
            suspected["table"].append(block)
        else:
            result["table"].append(block)

    if suspected["formula"] or suspected["table"]:
        llm_result = route_confirm(suspected)
        result["formula"].extend(llm_result.get("formula", []))
        result["table"].extend(llm_result.get("table", []))

    return result
