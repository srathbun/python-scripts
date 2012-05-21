#!/usr/bin/env python
# Filter ANSI escapes from stdin to stdout
from __future__ import print_function
from colorama import init
import sys
init()

for line in sys.stdin.readlines():
	print(line)
