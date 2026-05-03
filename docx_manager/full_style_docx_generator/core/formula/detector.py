import re

_LATEX_COMMANDS = re.compile(
    r'\\(?:frac|sum|int|sqrt|alpha|beta|gamma|delta|epsilon|zeta|eta|theta|lambda|mu|nu|xi|pi|rho|sigma|tau|phi|psi|omega'
    r'|Alpha|Beta|Gamma|Delta|Theta|Lambda|Sigma|Phi|Psi|Omega'
    r'|infty|partial|nabla|cdot|times|div|pm|mp|leq|geq|neq|approx|equiv|in|notin|subset|supset|cup|cap'
    r'|left|right|begin|end|text|mathrm|mathbf|overline|hat|vec|dot|ddot|bar|lim|log|ln|sin|cos|tan'
    r'|prod|int|oint|iint|iiint|binom|pmatrix|bmatrix|vmatrix)'
)

_MATH_UNICODE = re.compile(
    r'[∑∫∏√∞±×÷≠≤≥≈≡∈∉⊂⊃∪∩∂∇∝⊕⊗∧∨¬→←↔⟹⟺αβγδεζηθλμνξπρστφψωΑΒΓΔΘΛΣΦΨΩ]'
)

_LATEX_DELIMITERS = re.compile(r'\$\$?.+?\$\$?', re.DOTALL)

_SUPERSCRIPT_SUBSCRIPT = re.compile(r'[a-zA-Z0-9]\s*[\^_]\s*[{a-zA-Z0-9]')

_MATH_EXPRESSION = re.compile(
    r'[a-zA-Z][a-zA-Z0-9]*\s*=\s*[a-zA-Z0-9\(\)\+\-\*/\^]'
)


def is_suspected_formula(content: str) -> bool:
    if not content or not isinstance(content,str) or not content.strip():
        return False
    return bool(
        _LATEX_COMMANDS.search(content)
        or _MATH_UNICODE.search(content)
        or _LATEX_DELIMITERS.search(content)
        or _SUPERSCRIPT_SUBSCRIPT.search(content)
        or _MATH_EXPRESSION.search(content)
    )


def merge_formula_blocks(elements: list) -> list:
    """将 $$ ... $$ 之间的多个片段合并为一个元素，id 取第一个片段的 id。"""
    result = []
    i = 0
    while i < len(elements):
        content = elements[i].get('content', '').strip()
        if content == '$$':
            j = i + 1
            fragments = []
            while j < len(elements):
                if elements[j].get('content', '').strip() == '$$':
                    break
                fragments.append(elements[j])
                j += 1
            if fragments:
                result.append({
                    **fragments[0],
                    'content': '\n'.join(f.get('content', '') for f in fragments),
                })
            i = j + 1
        else:
            result.append(elements[i])
            i += 1
    return result


def detect_formula_blocks(elements: list) -> list:
    return [
        {"id": elem["id"], "content": elem["content"]}
        for elem in elements
        if is_suspected_formula(elem.get("content", ""))
    ]