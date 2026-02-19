import ast
from pathlib import Path

p = (
    Path(__file__).resolve().parent.parent
    / "projects"
    / "check data use"
    / "stock_datagrid.py"
)
with open(p, "r", encoding="utf-8") as f:
    src = f.read()
ast.parse(src)
print("syntax_ok")
