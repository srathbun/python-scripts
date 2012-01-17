#!/usr/bin/env python
'''
Project: Python parser to search php functions for regex
Author: Spencer Rathbun
Date: 1/17/2012

Summary: Search glob files (assumed to be php source) for input regex, and list function and line numbers where found.
'''
import re, argparse
from glob import glob
def main(regex, infiles):
	theRegex = re.compile(regex)

	for globFile in infiles:
		for f in glob(globFile):
			with open(f,'rb') as currFile:
				for lineNum, line in enumerate(currFile.readlines()):
					if line.split(' ')[0] == 'function':
						currFunction = line.split(' ')[1].split('(')[0]
					if theRegex.search(line) != None:
						print("In function {0}\n\tLine {1}: {2}".format(currFunction, lineNum+1, line))


if __name__ == "__main__":
	
	parser = argparse.ArgumentParser(description='Search php source code for regex', version='%(prog)s 1.0')
	parser.add_argument('regex', type=str, default='', help='Regular expression')
	parser.add_argument('infiles', nargs='+', type=str, help='list of input files')
	args = parser.parse_args()
	main(args.regex, args.infiles)
