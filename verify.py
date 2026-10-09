# -*- coding: utf-8 -*-
"""验证 xlsx 内部数据是否完整"""
import zipfile, re, os

p = r'd:\project\3\耀州窑青釉刻花瓶_数据集.xlsx'
print('文件存在:', os.path.exists(p))
print('文件大小:', os.path.getsize(p), 'bytes')
print()

z = zipfile.ZipFile(p)
print('内部部件:')
for n in z.namelist():
    print(f'  {n}  ({z.getinfo(n).file_size} bytes)')
print()

# 检查每张表
for i in range(1, 4):
    x = z.read(f'xl/worksheets/sheet{i}.xml').decode('utf-8')
    rows = re.findall(r'<row r="(\d+)"', x)
    cells = re.findall(r'<c r="([A-Z]+\d+)"', x)
    print(f'sheet{i}:')
    print(f'  行数: {len(rows)}')
    print(f'  单元格数: {len(cells)}')
    if rows:
        print(f'  行号范围: {rows[0]} - {rows[-1]}')
    # 显示前几个单元格引用
    print(f'  前10个单元格: {cells[:10]}')
    print()

# 检查共享字符串
ss = z.read('xl/sharedStrings.xml').decode('utf-8')
strings = re.findall(r'<t[^>]*>(.*?)</t>', ss, re.S)
print(f'共享字符串总数: {len(strings)}')
print('前5个字符串:')
for i, s in enumerate(strings[:5]):
    print(f'  [{i}] {s[:80]}')
print()
print('最后5个字符串:')
for i, s in enumerate(strings[-5:]):
    print(f'  [{len(strings)-5+i}] {s[:80]}')
