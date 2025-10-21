import sys
p = r"C:\Users\jhunj\puython\check data use\Check_File.py"
with open(p, 'rb') as f:
    data = f.read()
text = data.decode('utf-8')
for i, line in enumerate(text.splitlines(), 1):
    if ('浏览' in line or '退' in line or '文件检' in line or '根目' in line or '缺配' in line or '连续' in line):
        print(i, line.encode('utf-8'))
