"""同一个输入，只换写入器。

输入 = 从 shape.json + 固定 seed 生成的脱敏网格（与 make_sample.py 生成的完全一致）。
每个写入器跑在独立进程里，分别记录：

  1) 构造 DataFrame（或等价内存结构）
  2) 写入文件
  3) 写入前进程内存 / 写入过程中峰值内存（Windows PeakWorkingSetSize）
  4) 输出文件大小

用法：  python run_writers.py            # 跑全部（默认 96125 行）
        python run_writers.py 20000      # 按 shape 比例缩放到 20000 行

注意：这里的数据类型是 object dtype（逐格类型与样本一致）。pandas 的 dtype 会影响写入路径，
真从数据库取数时数值列会是 int64/float64，那一列要单独测，不能和这张表混。
"""

import ctypes
import ctypes.wintypes as wt
import datetime
import json
import os
import random
import subprocess
import sys
import time

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, HERE)
import make_sample  # noqa: E402

SEED = 20261113


class PMC(ctypes.Structure):
    _fields_ = [
        ('cb', wt.DWORD), ('Flags', wt.DWORD),
        ('PeakWorkingSetSize', ctypes.c_size_t), ('WorkingSetSize', ctypes.c_size_t),
        ('QuotaPeakPagedPoolUsage', ctypes.c_size_t), ('QuotaPagedPoolUsage', ctypes.c_size_t),
        ('QuotaPeakPrivatePoolUsage', ctypes.c_size_t), ('QuotaPrivatePoolUsage', ctypes.c_size_t),
        ('PageFaultCount', wt.DWORD), ('WorkingSetSizeEx', ctypes.c_size_t),
        ('PeakWorkingSetSizeEx', ctypes.c_size_t),
    ]


def mem():
    pmc = PMC()
    pmc.cb = ctypes.sizeof(PMC)
    ctypes.windll.kernel32.GetProcessMemoryInfo(ctypes.c_void_p(-1), ctypes.byref(pmc), ctypes.byref(ctypes.wintypes.DWORD(ctypes.sizeof(PMC))))
    return pmc.WorkingSetSize / 1024 / 1024, pmc.PeakWorkingSetSize / 1024 / 1024


def scaled_spec(shape, n_rows):
    f = n_rows / shape['rows']
    out = []
    for c in shape['columns']:
        s = {}
        for k, v in c.items():
            if k == 'name':
                continue
            if k in ('str_len', 'distinct_str_len', 'date_offsets'):
                s[k] = {int(key): max(1, round(n * f)) for key, n in v.items()}
            else:
                s[k] = {key: max(1, round(n * f)) for key, n in v.items()}
        # 四舍五入后每列的格数必须正好等于目标行数，把误差补到最大的那一格上
        total = sum(s['counts'].values())
        big = max(s['counts'], key=lambda t: s['counts'][t])
        s['counts'][big] += n_rows - total
        s['name'] = c['name']
        out.append(s)
    return out


def CounterInt(d):
    from collections import Counter
    return Counter(d)


def build_grid(spec, base_date):
    from collections import Counter
    sp = [{k: (Counter(v) if isinstance(v, dict) else v) for k, v in s.items()} for s in spec]
    n = sum(sp[0]['counts'].values())
    return make_sample.build_grid(n, sp, base_date, random.Random(SEED)), n


def write_with(method, grid, names, path, df=None):
    t0 = time.time()
    if method == 'pandas':
        df.to_excel(path, index=False)
    elif method == 'openpyxl':
        import openpyxl
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = 'Sheet1'
        fmt = wb.add_format({'num_format': 'yyyy-mm-dd'}) if hasattr(wb, 'add_format') else None
        ws.append(names)
        for r, row in enumerate(grid, start=1):
            for c, v in enumerate(row, start=1):
                if v is None:
                    continue
                if isinstance(v, datetime.date):
                    cell = ws.cell(row=r, column=c, value=v)
                    cell.number_format = 'yyyy-mm-dd'
                else:
                    ws.cell(row=r, column=c, value=v)
        wb.save(path)
    elif method in ('xlsxwriter', 'xlsxwriter_cm'):
        import xlsxwriter
        wb = xlsxwriter.Workbook(path, {'constant_memory': method == 'xlsxwriter_cm'})
        ws = wb.add_worksheet('Sheet1')
        fmt = wb.add_format({'num_format': 'yyyy-mm-dd'})
        for c, name in enumerate(names):
            ws.write(0, c, name)
        for r, row in enumerate(grid, start=1):
            for c, v in enumerate(row):
                if v is None:
                    continue
                if isinstance(v, str):
                    ws.write(r, c, v)
                elif isinstance(v, datetime.date):
                    ws.write_datetime(r, c, v, fmt)
                else:
                    ws.write_number(r, c, v)
        wb.close()
    elif method == 'fastxlsx':
        sys.path.insert(0, r'C:\Data\Jask\technology\定期任务\财务报表')
        import fastxlsx
        fastxlsx.write_xlsx(df, path)
    return time.time() - t0


def run_one(method, shape, n_rows, out_dir):
    base_date = datetime.date.fromisoformat(shape['base_date'])
    spec = scaled_spec(shape, n_rows)
    grid, n = build_grid(spec, base_date)
    names = [s['name'] for s in spec]

    t_build = time.time()
    df = None
    if method in ('pandas', 'fastxlsx'):
        import pandas as pd
        df = pd.DataFrame(grid, columns=names)
    build_time = time.time() - t_build

    before, peak_before = mem()
    write_time = write_with(method, grid, names, os.path.join(out_dir, 'x.xlsx'), df)
    after, peak_after = mem()

    size = os.path.getsize(os.path.join(out_dir, 'x.xlsx')) / 1024 / 1024
    import pandas as pd  # noqa: F401  (version line for the report)
    return {
        'method': method, 'rows': n, 'cols': len(names),
        'build_df_s': round(build_time, 2), 'write_s': round(write_time, 2),
        'mem_before_mb': round(before, 1), 'mem_peak_mb': round(max(peak_after, peak_before), 1),
        'file_mb': round(size, 1),
    }


def main():
    n_rows = int(sys.argv[1]) if len(sys.argv) > 1 else None
    shape = json.load(open(os.path.join(HERE, 'shape.json'), encoding='utf-8'))
    if n_rows is None:
        n_rows = shape['rows']
    methods = ['pandas', 'openpyxl', 'xlsxwriter', 'xlsxwriter_cm', 'fastxlsx']
    if os.environ.get('WORKER'):
        print(json.dumps(run_one(sys.argv[2], shape, n_rows, HERE)))
        return
    rows = []
    for m in methods:
        p = subprocess.run([sys.executable, os.path.abspath(__file__), str(n_rows), m],
                           capture_output=True, text=True, env=dict(os.environ, WORKER='1'))
        line = [l for l in p.stdout.splitlines() if l.strip().startswith('{')]
        if not line:
            print(m, 'FAILED:', p.stderr.strip()[:200])
            continue
        rows.append(json.loads(line[0]))
        print('{method:<14} build {build_df_s:>6.2f}s  write {write_s:>6.2f}s  '
              'mem {mem_before_mb:>7.1f} → {mem_peak_mb:>7.1f} MB  file {file_mb:>6.1f} MB'.format(**rows[-1]))
    json.dump(rows, open(os.path.join(HERE, 'results.json'), 'w', encoding='utf-8'), ensure_ascii=False, indent=1)


if __name__ == '__main__':
    main()
