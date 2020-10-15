# XLSXWriter
Python. Формирование xlsx-файла<br/>

Время на формирование файла, мс
| Кол-во ячеек | Python3.6 | PyPy3.5 | -m memory_profiler |
| -- | -- | -- | -- |
| 1000 | 4 | 6 | 0.152 MiB |
| 10000 | 27 | 11 | 0.152 MiB |
| 100000 | 248 | 40 | 0.152 MiB |
| 1000000 | 2375 | 1014 | 0.145 MiB |
<br/>
Простой пример:
<br/>

```python
import XLSXWriter

writer = XLSXWriter.Writer()
writer.sheetAdd('Sheet1')

writer.writeSheetRow( ['text'] )

writer.saveAs('test.xlsx')
```