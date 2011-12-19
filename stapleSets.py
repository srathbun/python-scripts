#!/usr/bin/env python
'''
Project: stapleSets
Author: Spencer Rathbun
Date: 9/29/2011

Summary: Python script to create sets of stapled sheets for print file downloader

'''
import argparse, sys
import itertools
from collections import defaultdict

if __name__ == "__main__":
	
	parser = argparse.ArgumentParser(description='Take a number of pages, and the pages in several sets, to produce an output for copy and paste into the print file downloader', version='%(prog)s 2.2')
	parser.add_argument('pages', type=int, help='Total number of pages to break into sets')
	parser.add_argument('stapleset', nargs='+', type=int, help='number of pages in each set')
	parser.add_argument('-p', action='store_true', default=False, help='Switch the page separator from - to ..')
	parser.add_argument('-g', action='store_true', default=False, help='Switch the group separator from , to ;')
	args = parser.parse_args()

	if args.p:
		pageSep = '..'
	else:
		pageSep = '-'

	if args.g:
		groupSep = ';'
	else:
		groupSep = ','

	partition_lengths = [(pos,s) for pos,s in enumerate(args.stapleset)]
	range_start = 1
	range_end = args.pages
	endpoints = defaultdict(list)
	for pos,p_len in itertools.cycle(partition_lengths):
		end = range_start + p_len - 1
		if end > range_end and range_start < range_end: 
			endpoints[str(pos) + str(p_len)].append((range_start, range_end))
			break
		elif range_start > range_end:
			break
		endpoints[str(pos) + str(p_len)].append((range_start, end))
		range_start += p_len

	for pos,partition in enumerate(args.stapleset):
		for c,t in enumerate(endpoints[str(pos) + str(partition)]):
			if c > 0:
				sys.stdout.write("{0}".format(groupSep))
			if t[0] == t[1]:
				sys.stdout.write("{0}".format(t[0]))
			else:
				sys.stdout.write("{0}{1}{2}".format(t[0],pageSep,t[1]))
		sys.stdout.write("\n\n")
