"""DataFrame -> xlsx（只依赖标准库）。

字符串、ID、日期一律按文本写入（与财务月报现有导出一致），数值仍是数值。
pandas.to_excel 要为每个单元格做一次类型判断和样式匹配；这里直接生成 OOXML 包并流式落盘。
"""
import datetime
import zipfile

CT_NS = 'http://schemas.openxmlformats.org/package/2006/content-types'
PKG_NS = 'http://schemas.openxmlformats.org/package/2006/relationships'
M = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'
R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'

_ESCAPES = {38: '&amp;', 60: '&lt;', 62: '&gt;'}
# Excel 不接受文本里的 ASCII 控制字符，直接丢弃
for _c in range(32):
    if _c not in (9, 10, 13):
        _ESCAPES[_c] = ''


def _colletter(j):
    s = ''
    j += 1
    while j:
        j, r = divmod(j - 1, 26)
        s = chr(65 + r) + s
    return s


def _isblank(v):
    return v is None or v != v or v.__class__.__name__ == 'NaTType'


def _text(v):
    if isinstance(v, str):
        return v
    if isinstance(v, datetime.datetime):
        # 与 Excel 里现在的显示一致：纯日期不补时间
        return v.strftime('%Y-%m-%d %H:%M:%S') if (v.hour or v.minute or v.second) else v.strftime('%Y-%m-%d')
    if isinstance(v, datetime.date):
        return v.strftime('%Y-%m-%d')
    if isinstance(v, bool) or type(v).__name__ == 'bool_':
        return str(v)
    return None


def _number(v):
    if isinstance(v, float) and v.is_integer() and abs(v) < 1e15:
        return str(int(v))
    return str(v)


def write_xlsx(df, path):
    cols = [str(c) for c in df.columns]
    ncols = len(cols)

    order = list(cols)
    sst = {c: i for i, c in enumerate(order)}
    rows = []
    total = 0
    for row in df.itertuples(index=False, name=None):
        out = []
        for v in row:
            if _isblank(v):
                out.append(None)
                continue
            t = _text(v)
            if t is not None:
                total += 1
                if t not in sst:
                    sst[t] = len(order)
                    order.append(t)
                out.append(('s', sst[t]))
            else:
                out.append(('n', _number(v)))
        rows.append(out)

    refs = [_colletter(j) for j in range(ncols)]
    xml_open = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
    z = zipfile.ZipFile(path, 'w', zipfile.ZIP_DEFLATED, 6)
    z.writestr('[Content_Types].xml', xml_open + '<Types xmlns="' + CT_NS + '">'
               '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
               '<Default Extension="xml" ContentType="application/xml"/>'
               '<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>'
               '<Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>'
               '<Override PartName="/xl/sharedStrings.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml"/>'
               '<Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>'
               '</Types>')
    z.writestr('_rels/.rels', xml_open + '<Relationships xmlns="' + PKG_NS + '">'
               '<Relationship Id="rId1" Type="' + R + '/officeDocument" Target="xl/workbook.xml"/></Relationships>')
    z.writestr('xl/workbook.xml', xml_open + '<workbook xmlns="' + M + '" xmlns:r="' + R + '">'
               '<sheets><sheet name="Sheet1" sheetId="1" r:id="rId1"/></sheets></workbook>')
    z.writestr('xl/_rels/workbook.xml.rels', xml_open + '<Relationships xmlns="' + PKG_NS + '">'
               '<Relationship Id="rId1" Type="' + R + '/worksheet" Target="worksheets/sheet1.xml"/>'
               '<Relationship Id="rId2" Type="' + R + '/styles" Target="styles.xml"/>'
               '<Relationship Id="rId3" Type="' + R + '/sharedStrings" Target="sharedStrings.xml"/></Relationships>')
    z.writestr('xl/styles.xml', xml_open + '<styleSheet xmlns="' + M + '">'
               '<fonts count="1"><font><sz val="11"/><name val="Calibri"/></font></fonts>'
               '<fills count="1"><fill><patternFill patternType="none"/></fill></fills>'
               '<borders count="1"><border/></borders>'
               '<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>'
               '<cellXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/></cellXfs>'
               '</styleSheet>')
    z.writestr('xl/sharedStrings.xml', xml_open + '<sst xmlns="' + M + '" count="' + str(total) +
               '" uniqueCount="' + str(len(order)) + '">'
               + ''.join('<si><t>' + s.translate(_ESCAPES) + '</t></si>' for s in order) + '</sst>')

    f = z.open('xl/worksheets/sheet1.xml', 'w')
    buf = [xml_open, '<worksheet xmlns="' + M + '">',
           '<dimension ref="A1:' + _colletter(ncols - 1) + str(len(rows) + 1) + '"/>',
           '<sheetData><row r="1">']
    for j, c in enumerate(cols):
        buf.append('<c r="' + refs[j] + '1" t="s"><v>' + str(sst[c]) + '</v></c>')
    buf.append('</row>')

    for i, row in enumerate(rows, 2):
        buf.append('<row r="' + str(i) + '">')
        for j, v in enumerate(row):
            if v is None:
                continue
            cell = '<c r="' + refs[j] + str(i) + '"'
            if v[0] == 's':
                buf.append(cell + ' t="s"><v>' + str(v[1]) + '</v></c>')
            else:
                buf.append(cell + '><v>' + v[1] + '</v></c>')
        buf.append('</row>')
    buf.append('</sheetData></worksheet>')
    f.write(''.join(buf).encode())
    f.close()
    z.close()
