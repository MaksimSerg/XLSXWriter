import XLSXWriter
import time
import sys
import os
# from memory_profiler import profile


def getTime():
    return round(time.time() * 1000)


# @profile
def main():
    testFilePath = "test.xlsx"

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


if __name__ == '__main__':
    main()
