#!/usr/bin/env python
# Filter ANSI escapes from stdin to stdout
from __future__ import print_function
from colorama import init, AnsiToWin32
import sys
init(wrap=False)

for line in sys.stdin.readlines():
	print(line, file=AnsiToWin32(sys.stdout))
