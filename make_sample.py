"""从 shape.json 生成脱敏样本。这一步不读任何业务数据，输入只有那份计数 + 固定 seed。

保住的：行数、列数、每列类型计数、每列字符串长度分布、每列 distinct 值数、空值计数、
日期分布、数字正负号分布（符号挂在值上，不挂在格上）。
字符串按你定的规则：每个格独立随机生成，只保长度，所以没有可复用的重复值。
值的出现频率分布不保。

跑之前先跑一次 profile.py 生成 shape.json。
"""

import datetime
import json
import os
import random
from collections import Counter

import openpyxl
import xlsxwriter

OUT_DIR = os.path.dirname(os.path.abspath(__file__))
SEED = 20261113
ALPHA = 'abcdefghijklmnopqrstuvwxyz'
INT_KEYS = ('str_len', 'distinct_str_len', 'date_offsets')


def load_shape(path):
    d = json.load(open(path, encoding='utf-8'))
    spec = []
    for c in d['columns']:
        s = {}
        for k, v in c.items():
            if k == 'name':
                continue
            if k in INT_KEYS:
                s[k] = Counter({int(key): n for key, n in v.items()})
            else:
                s[k] = Counter(v)
        s['name'] = c['name']
        spec.append(s)
    return d['rows'], spec, datetime.date.fromisoformat(d['base_date'])


def fix(hist, target):
    total = sum(hist.values())
    if total != target:
        big = max(hist, key=lambda k: hist[k])
        hist[big] += target - total


def scale_spec(spec, factor):
    """按比例放大形状。注意：每个（整数位数, 小数位）组能装下的不同值有上限，
    放大后顶到上限就只能重复——这是取值空间限制，不是生成方式的问题。"""
    out, warn = [], []
    for s in spec:
        n = {k: {key: max(1, round(v * factor)) for key, v in d.items()} for k, d in s.items() if k != 'name'}
        n['name'] = s['name']
        total = sum(n['counts'].values())
        fix(n['counts'], total)
        fix(n['str_len'], n['counts'].get('str', 0))
        fix(n['num_cell'], n['counts'].get('int', 0) + n['counts'].get('float', 0))
        fix(n['date_offsets'], n['counts'].get('date', 0))
        fix(n['distinct_sign'], sum(n['num_distinct'].values()))
        for group, cnt in list(n['num_distinct'].items()):
            digits, dp = (int(x) for x in group.split(','))
            cap = (9 * 10 ** (digits - 1) if digits > 1 else 10) * (1 if dp == 0 else max(1, 9 * 10 ** (dp - 1)))
            if cnt > cap:
                warn.append('列 {} 组 {}: distinct {} > 上限 {}'.format(s['name'], group, cnt, cap))
                n['num_distinct'][group] = cap
        out.append(n)
    return out, warn


def build_grid(nrows, spec, base_date, rng):
    grid = [[None] * len(spec) for _ in range(nrows)]
    for i, s in enumerate(spec):
        slots = []
        for t, c in s['counts'].items():
            slots.extend([t] * c)
        assert len(slots) == nrows, '列 {} 槽位数 {} != 行数 {}'.format(s['name'], len(slots), nrows)
        rng.shuffle(slots)
        pos = [r for r, t in enumerate(slots) if t != 'empty']

        # 字符串：每格独立随机，只保长度
        str_pos = [p for p in pos if slots[p] == 'str']
        cur = 0
        for L, count in sorted(s['str_len'].items()):
            for p in str_pos[cur:cur + count]:
                grid[p][i] = ''.join(rng.choice(ALPHA) for _ in range(int(L)))
            cur += count

        # 数字：按（整数位数, 小数位）分组生成 distinct 个值，符号跟值走
        num_pos = [p for p in pos if slots[p] in ('int', 'float')]
        cursor = 0
        for group, dcount in sorted(s['num_distinct'].items()):
            digits, dp = (int(x) for x in group.split(','))
            base = 10 ** (digits - 1) if digits > 1 else 0
            cap = 9 * 10 ** (digits - 1) if digits > 1 else 10
            if dp == 0:
                vals = [base + k for k in range(cap)][:dcount]
            else:
                fracs = [f for f in range(10 ** dp) if f % 10]
                vals = []
                for a in range(cap):
                    for b in range(len(fracs)):
                        vals.append(round(base + a + fracs[b] / 10 ** dp, dp))
                        if len(vals) == dcount:
                            break
                    if len(vals) == dcount:
                        break
            signs = []
            for key, c in sorted(s['distinct_sign'].items()):
                if key.startswith(group + ','):
                    sign = key[len(group) + 1:]
                    signs.extend(-1 if sign == '-' else 1 for _ in range(c))
            targets = num_pos[cursor:cursor + s['num_cell'][group]]
            cursor += s['num_cell'][group]
            if not vals or not signs:
                continue
            for j, p in enumerate(targets):
                v = vals[j % len(vals)]
                if slots[p] == 'float':
                    v = float(v)
                grid[p][i] = signs[j % len(signs)] * v

        # 日期：按真实偏移分布逐格还原
        offsets = []
        for off, c in sorted(s['date_offsets'].items()):
            offsets.extend([off] * c)
        rng.shuffle(offsets)
        for p, off in zip([p for p in pos if slots[p] == 'date'], offsets):
            grid[p][i] = base_date + datetime.timedelta(days=int(off))
    return grid


def write_xlsx(grid, spec, path):
    wb = xlsxwriter.Workbook(path)
    ws = wb.add_worksheet('Sheet1')
    fmt = wb.add_format({'num_format': 'yyyy-mm-dd'})
    for i, s in enumerate(spec):
        ws.write(0, i, s['name'])
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


def verify(path, spec):
    ws = openpyxl.load_workbook(path, read_only=True).active
    chk = [Counter() for _ in spec]
    chkg = [Counter() for _ in spec]
    distinct = [set() for _ in spec]
    lenhist = [Counter() for _ in spec]
    rows = 0
    for row in ws.iter_rows(min_row=2):
        rows += 1
        for i, cell in enumerate(row):
            v = cell.value
            if v is None:
                chk[i]['empty'] += 1
            elif isinstance(v, str):
                chk[i]['str'] += 1
                lenhist[i][len(v)] += 1
                distinct[i].add(v)
            elif isinstance(v, bool):
                chk[i]['bool'] += 1
            elif isinstance(v, (int, float)):
                kind = 'int' if isinstance(v, int) else 'float'
                chk[i][kind] += 1
                digits = len(str(abs(int(abs(v)))))
                dp = 0
                if kind == 'float':
                    txt = repr(v)
                    dp = len(txt.split('.')[1]) if '.' in txt else 0
                chkg[i]['{},{}'.format(digits, dp)] += 1
                distinct[i].add(v)
            elif isinstance(v, datetime.datetime):
                chk[i]['date'] += 1
                distinct[i].add(v)
    bad = 0
    for i, s in enumerate(spec):
        for t, c in s['counts'].items():
            if chk[i][t] != c:
                bad += 1
                print('  列 {} {} 类型 {}: 期望 {} 实际 {}'.format(i, s['name'], t, c, chk[i][t]))
        for group, c in s['num_cell'].items():
            if chkg[i][group] != c:
                bad += 1
                print('  列 {} {} 数字组 {}: 期望 {} 实际 {}'.format(i, s['name'], group, c, chkg[i][group]))
        want = sum(s['num_distinct'].values())
        got = len(distinct[i])
        if want and got != want:
            print('  说明 列 {} {} distinct: 形状里是 {}，这份是 {}（字符串每格随机；数字组可能被取值上限截断）'.format(i, s['name'], want, got))
        if s['str_len'] and lenhist[i] != s['str_len']:
            bad += 1
            print('  列 {} {} 长度分布不一致'.format(i, s['name']))
    print('复核:', '全部一致' if bad == 0 else '{} 项不一致'.format(bad), '(读回 {} 行)'.format(rows))
    return bad


def main():
    shape = os.path.join(OUT_DIR, 'shape.json')
    if not os.path.exists(shape):
        raise SystemExit('没有 shape.json，先跑一次 profile.py（那一步才需要原始文件）')
    nrows, spec, base_date = load_shape(shape)
    print('从 shape.json 生成: {} 行 x {} 列'.format(nrows, len(spec)))
    grid = build_grid(nrows, spec, base_date, random.Random(SEED))
    out = os.path.join(OUT_DIR, 'sample_{}x{}.xlsx'.format(nrows, len(spec)))
    write_xlsx(grid, spec, out)
    print('样本已生成:', out, os.path.getsize(out) // 1024, 'KB')
    verify(out, spec)


if __name__ == '__main__':
    main()
