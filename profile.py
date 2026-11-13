"""一次性：从真实文件里读出"形状"，写成 shape.json。

这是唯一会碰到你原始数据的一步。shape.json 里只有计数（每列类型计数、字符串长度分布、
distinct 值数、空值计数、日期分布、正负号分布），没有任何业务内容。
生成脚本不再需要这个文件。
"""

import datetime
import json
import os
from collections import Counter

import openpyxl

SRC = r'C:\Users\admin\Desktop\data\财务每月报表_业务_2026-07_20261109115102.xlsx'
OUT_DIR = os.path.dirname(os.path.abspath(__file__))
BASE_DATE = datetime.date(2026, 1, 1)


def main():
    ws = openpyxl.load_workbook(SRC, read_only=True).active
    header = [c.value for c in next(ws.iter_rows(min_row=1, max_row=1))]
    spec = [{'name': header[i], 'counts': Counter(), 'str_len': Counter(),
             'distinct_str_len': Counter(), 'num_cell': Counter(), 'num_distinct': Counter(),
             'distinct_sign': Counter(), 'sign': Counter(), 'date_offsets': Counter(),
             '_seen': set()} for i in range(len(header))]

    nrows = 0
    for row in ws.iter_rows(min_row=2):
        nrows += 1
        for i, cell in enumerate(row):
            v = cell.value
            s = spec[i]
            if v is None:
                s['counts']['empty'] += 1
            elif isinstance(v, str):
                s['counts']['str'] += 1
                s['str_len'][len(v)] += 1
                if v not in s['_seen']:
                    s['_seen'].add(v)
                    s['distinct_str_len'][len(v)] += 1
            elif isinstance(v, bool):
                s['counts']['bool'] += 1
            elif isinstance(v, (int, float)):
                kind = 'int' if isinstance(v, int) else 'float'
                s['counts'][kind] += 1
                digits = len(str(abs(int(abs(v)))))
                dp = 0
                if kind == 'float':
                    txt = repr(v)
                    dp = len(txt.split('.')[1]) if '.' in txt else 0
                group = '{},{}'.format(digits, dp)
                s['num_cell'][group] += 1
                if v not in s['_seen']:
                    s['_seen'].add(v)
                    s['num_distinct'][group] += 1
                    s['distinct_sign'][group + ',{}'.format('-' if v < 0 else '+')] += 1
                s['sign']['-' if v < 0 else '+'] += 1
            elif isinstance(v, datetime.datetime):
                s['counts']['date'] += 1
                s['date_offsets'][(v.date() - BASE_DATE).days] += 1

    for s in spec:
        s.pop('_seen', None)

    json.dump({'rows': nrows, 'cols': len(header), 'base_date': BASE_DATE.isoformat(),
               'columns': [dict((k, (dict(v) if isinstance(v, Counter) else v))
                               for k, v in s.items() if k != '_seen') for s in spec]},
              open(os.path.join(OUT_DIR, 'shape.json'), 'w', encoding='utf-8'),
              ensure_ascii=False, indent=1)
    print('shape.json 已生成（{} 行 x {} 列，只有计数）'.format(nrows, len(header)))


if __name__ == '__main__':
    main()
