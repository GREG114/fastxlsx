"""直接读 sample_96125x83.xlsx，用两种方式导出。

输入就是那份文件，用 openpyxl 读（不用 pd.read_excel，那条会把类型读错）。
读文件和构造 DataFrame 单独计时，不算进写入。

用法：python test_export.py [xlsx路径]
"""

import datetime
import json
import os
import sys
import time

import openpyxl
import pandas as pd

import fastxlsx

HERE = os.path.dirname(os.path.abspath(__file__))
OUT = os.path.join(HERE, 'out')
os.makedirs(OUT, exist_ok=True)


def read_sample(path):
    ws = openpyxl.load_workbook(path, read_only=True).active
    header = [c.value for c in next(ws.iter_rows(min_row=1, max_row=1))]
    rows = []
    for row in ws.iter_rows(min_row=2):
        rows.append([c.value for c in row])
    return header, rows


def main():
    path = sys.argv[1] if len(sys.argv) > 1 else os.path.join(HERE, 'sample_96125x83.xlsx')
    t0 = time.time()
    names, data = read_sample(path)
    read_time = time.time() - t0
    print('读样本: {} 行 x {} 列，{:.2f}s'.format(len(data), len(names), read_time))

    t0 = time.time()
    df = pd.DataFrame(data, columns=names)
    build = time.time() - t0
    print('构造 DataFrame: {:.2f}s（两边共用，不计入写入）'.format(build))

    base = os.path.splitext(os.path.basename(path))[0]
    results = {'input': {'file': os.path.basename(path), 'rows': len(data), 'cols': len(names),
                         'read_s': round(read_time, 2)},
               'env': {'python': sys.version.split()[0], 'pandas': pd.__version__,
                       'pandas_default_engine': '由环境里装的库决定，看这条 traceback'},
               'build_df_s': round(build, 2), 'results': []}

    cases = [
        ('pandas.to_excel（老方法）', lambda p: df.to_excel(p, index=False), '{}_old.xlsx'.format(base)),
        ('fastxlsx.write_xlsx（新方法）', lambda p: fastxlsx.write_xlsx(df, p), '{}_new.xlsx'.format(base)),
    ]
    for label, write, fname in cases:
        out = os.path.join(OUT, fname)
        t0 = time.time()
        try:
            write(out)
            row = {'method': label, 'write_s': round(time.time() - t0, 2),
                   'file_mb': round(os.path.getsize(out) / 1048576, 1)}
            print('{method:<32} write {write_s:>7.2f}s  file {file_mb:>5.1f} MB'.format(**row))
        except Exception as e:
            row = {'method': label, 'error': type(e).__name__ + ': ' + str(e)[:120]}
            print('{:<32} 失败 {:.2f}s  {}'.format(label, time.time() - t0, row['error']))
        results['results'].append(row)

    json.dump(results, open(os.path.join(HERE, 'test_result_{}.json'.format(base)), 'w',
                            encoding='utf-8'), ensure_ascii=False, indent=1)


if __name__ == '__main__':
    main()
