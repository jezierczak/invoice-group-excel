from typing import TypedDict

class FormulaDefinition(TypedDict):
    cell: str
    formula: str

class FormulaTemplate(TypedDict):
    column: str
    start_row: int
    end_row: int
    template: str