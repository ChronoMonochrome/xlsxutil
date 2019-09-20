# xlsxutil

A very minimalistic python module to edit Excel workbooks

## Features

- Reading from and writing to the existing workbooks is supported:

```python
from xlsxutil import Workbook

wb = Workbook('1.xlsx')
row = wb.worksheets['Лист1'].rows[0]
row.cells[0].value = "204"
wb.save("test1234.xlsx")
```

However, this won't update any worksheets that has links to the edited cells.

For now creating new workbooks or removing / adding rows to a worksheet is not supported.

## Credits

[Fast xlsx parsing with Python](https://blog.adimian.com/2018/09/04/fast-xlsx-parsing-with-python/)

[Overwriting file in ziparchive](https://stackoverflow.com/questions/4653768/overwriting-file-in-ziparchive)
