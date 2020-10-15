import XLSXWriter
import time
import sys
import os

testFilePath = "test.xlsx"


def getTime():
    return round(time.time() * 1000)


if os.path.isfile(testFilePath):
    os.remove(testFilePath)

rows_count = int(sys.argv[1]) if len(sys.argv) > 1 else 10000
row = [1, 2, 3, 4, 5, 6, 7, 8, 9, 0]

start = getTime()

writer = XLSXWriter.Writer()
writer.sheetAdd('Sheet1')

for i in range(rows_count):
    writer.writeSheetRow(row)

writer.saveAs(testFilePath)

end = getTime() - start

print("cells count: {}x{}. Time: {} ms".format(rows_count, len(row), end))

# rows*10	P3.6	PY3.5
"""
1_000		4		6
10_000 		27		11
100_000 	248		40
1_000_000	2375	1014

python3 ./test.py 10
/opt/pypy3.5/bin/pypy3 ./test.py 10

"""
