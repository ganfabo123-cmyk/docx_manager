from pydantic import BaseModel


class FormulaItem(BaseModel):
    id: str
    text_before: str = ""
    latex_formula: str
    text_after: str = ""
    label: str = ""


class FormulaListResponse(BaseModel):
    formulas: list[FormulaItem]
