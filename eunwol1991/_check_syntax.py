import ast, sys
p = r'check data use/stock_datagrid.py'
with open(p, 'r', encoding='utf-8') as f:
    src = f.read()
ast.parse(src)
print('syntax_ok')
