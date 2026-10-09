# -*- coding: utf-8 -*-
"""统计耀州窑青釉刻花瓶_数据集.xlsx 的数据体量"""
import zipfile, os, re

p = r'd:\project\3\耀州窑青釉刻花瓶_数据集.xlsx'
size = os.path.getsize(p)
print('=' * 50)
print('文件体量')
print('=' * 50)
print(f'文件大小: {size/1024:.2f} KB ({size:,} bytes)')
z = zipfile.ZipFile(p)
names = z.namelist()
print(f'内部 XML 部件数: {len(names)} 个')
comp_size = sum(zi.compress_size for zi in z.infolist())
print(f'压缩后总大小: {comp_size/1024:.2f} KB (压缩比 {comp_size/size*100:.1f}%)')
print()

print('=' * 50)
print('数据记录统计')
print('=' * 50)
total = 0
fields_per_row = 5
for i in range(1, 4):
    x = z.read(f'xl/worksheets/sheet{i}.xml').decode('utf-8')
    rows = re.findall(r'<row r="(\d+)"', x)
    data_rows = [r for r in rows if int(r) >= 3]
    n = len(data_rows)
    total += n
    print(f'表{i}: {n} 条数据记录 × {fields_per_row} 字段 = {n*fields_per_row} 个单元格')

print()
print(f'合计数据记录: {total} 条')
print(f'单元格总数(不含标题/表头): {total*fields_per_row} 个')
print(f'加上标题行3 + 表头行3 = 共 {total*fields_per_row + 6} 个有效单元格')

# 共享字符串
ss = z.read('xl/sharedStrings.xml').decode('utf-8')
strings = re.findall(r'<t[^>]*>(.*?)</t>', ss, re.S)
# 反转义
def unescape(s):
    return s.replace('&lt;', '<').replace('&gt;', '>').replace('&amp;', '&').replace('&quot;', '"').replace('&apos;', "'").replace('&#10;', '\n')
strings = [unescape(s) for s in strings]
total_chars = sum(len(s) for s in strings)
unique_chars = len(set(strings))

print()
print('=' * 50)
print('文本内容体量')
print('=' * 50)
print(f'共享字符串池: {len(strings)} 个 (去重后)')
print(f'文本字符总量: {total_chars:,} 字符')

# 按表统计字符
print()
print('=' * 50)
print('各表字符分布')
print('=' * 50)
import importlib.util
spec = importlib.util.spec_from_file_location("m", r"d:\project\3\build_xlsx.py")
m = importlib.util.module_from_spec(spec)
spec.loader.exec_module(m)

for name, subtitle, rows in m.SHEETS:
    row_chars = sum(sum(len(str(v)) for v in r) for r in rows)
    print(f'{name}:')
    print(f'  数据记录: {len(rows)} 条')
    print(f'  字符总量: {row_chars} 字符')
    print(f'  平均每条: {row_chars/len(rows):.1f} 字符')

# 来源统计
print()
print('=' * 50)
print('来源分布')
print('=' * 50)
from collections import Counter
sources = []
for _, _, rows in m.SHEETS:
    for r in rows:
        sources.append(r[4])
c = Counter(sources)
for src, n in c.most_common(10):
    print(f'  {src}: {n} 条')
