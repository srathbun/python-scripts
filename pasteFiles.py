'''
File: pasteFiles.py
Author: Spencer Rathbun
Date: 04/23/2012
Description: Merge multiple files together as one csv file.
'''
import glob, csv
from itertools import cycle, islice, count

def roundrobin(*iterables):
    "roundrobin('ABC', 'D', 'EF') --> A D E B F C"
    # Recipe credited to George Sakkis
    pending = len(iterables)
    nexts = cycle(iter(it).next for it in iterables)
    while pending:
        try:
            for next in nexts:
                yield next()
        except StopIteration:
            pending -= 1
            nexts = cycle(islice(nexts, pending))

Results = [open(f).readlines() for f in glob.glob("*.data")]
fout = csv.writer(open("res.csv", 'wb'), dialect="excel")

row=[]
for line, c in zip(roundrobin(Results), cycle(range(len(Results)))):
    splitline = line.split()
    for item,currItem in zip(splitline, count(1)):
        row[c+currItem] = item
    if count == len(Results):
        fout.writerow(row)
        row = []
del fout
