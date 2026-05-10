import sys
sys.path.insert(0, "D:/PycharmProjects/hit-paper-helper/docx_manager/full_style_docx_generator")

for k in list(sys.modules.keys()):
    if "formula" in k or "converter" in k:
        del sys.modules[k]

from core.formula.converter import _preprocess_latex, latex_to_omath

# The broken LaTeX extracted by LLM (from backfilled_styles.json elem_140)
broken = r"\alphal^{(t)} = \text{clamp}\!\left(\frac{\left\|h{l-1}^{(t)} - \hat{h}{l-1}^{(t-k)}\right\|2}{\left\|\hat{h}{l-1}^{(t-k)}\right\|2 + \epsilon},\ 0,\ 2\right)"

fixed = _preprocess_latex(broken)
print("=== preprocessor output ===")
print(fixed)
print()

# Check key fixes
print("\\alpha_l present:", r"\alpha_l" in fixed)
print("\\alphal gone:    ", r"\alphal" not in fixed)
print("h_{l-1} present: ", r"h_{l-1}" in fixed)
print("h{l-1} gone:     ", "h{l-1}" not in fixed)
print(r"\hat{h}_{l-1} present:", r"\hat{h}_{l-1}" in fixed)
print(r"\hat{h}{l-1} gone:    ", r"\hat{h}{l-1}" not in fixed)

# Full conversion
result = latex_to_omath(broken)
print("\n=== alpha char (U+03B1) in OMML:", chr(0x03B1) in result)
print("=== literal \\alphal in OMML:   ", r"\alphal" in result)

with open("D:/check_omml2.xml", "w", encoding="utf-8") as f:
    f.write(result)
print("\ndone, written to D:/check_omml2.xml")
