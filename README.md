# 这份文件夹里有什么

[文章.md](文章.md) 是成稿。剩下的就是跑那组数字时用的输入和脚本。

## 数字（照实抄，文章里的"84 秒 / 7 分 15 秒"就是这些）

| 行数 | 写法 | 写入耗时 | 输出文件 |
|---|---|---|---|
| 96,125 | `pandas.to_excel` 默认引擎（这台机器解析成 xlsxwriter） | 84.09 s | 65,062,142 B |
| 96,125 | `fastxlsx.write_xlsx` | 21.36 s | 57,867,501 B |
| 500,000 | `pandas.to_excel` 默认引擎 | 434.90 s | 339,184,330 B |
| 500,000 | `pandas.to_excel engine='openpyxl'` | 547.40 s | 338,910,466 B |
| 500,000 | `fastxlsx.write_xlsx` | 138.43 s | 305,979,626 B |

1,000,000 行这一档没有跑完计时：默认引擎在关闭文件时抛 `xlsxwriter.exceptions.FileSizeError: Filesize would require ZIP64 extensions`；`fastxlsx` 走标准库 `zipfile`，默认 `allowZip64=True`，同一份输入不会卡在这一步。这一档只确认了报错，没有"跑完"的计时。

跑这些数字的环境：Python 3.12，pandas 3.0.1。

## 流程

1. [profile.py](profile.py) —— **唯一会碰原始业务数据的一步**。从那个 xlsx 里只读出形状（每列类型计数、字符串长度分布、distinct 值数、空值计数、日期分布、正负号分布），写成 [shape.json](shape.json)。这份只有计数，没有任何业务内容，可以公开。
2. [make_sample.py](make_sample.py) —— 只读 shape.json，按固定 seed（`20261113`）生成 [sample_96125x83.xlsx](sample_96125x83.xlsx)。它不需要那个原始文件。字符串按规则：每格独立随机生成，只保长度，所以没有可复用的重复值；数字按（整数位数, 小数位）分组生成 distinct 个值，符号跟值走；日期按真实偏移分布。跑完自己复核：类型计数、长度分布、数字分组逐项一致。
3. [test_export.py](test_export.py) —— 同一个输入，只换写入器；构造 DataFrame 单独计时，不算进写入。默认跑 shape.json 里的行数，给参数就按比例放大。

[fastxlsx.py](fastxlsx.py) 是文章里用到的那个模块，入口只有一个：`write_xlsx(df, path)`。

## 两个已知限制

- 字符串每格随机之后，distinct 值数跟真实样本不一样（真实样本约 19.7:1 的重复率，这份是零重复率）。所以这张表里的倍数是"没有去重红利"时测出来的。
- 放大到 50 万行时，金额、税额这类列的不同取值数会顶到那个位数区间的取值空间上限，只能重复。这是取值空间限制，不是生成方式的问题。

刚才那次移动把之前生成的输出文件（96k 的三份、500k 的三份）弄丢了，输入样本可以重新生成，输出要重新拿只能重跑测试。
