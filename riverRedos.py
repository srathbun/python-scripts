#!/usr/bin/env python
'''
File: riverRedos.py
Author: Spencer Rathbun
Date: 9/7/11
Description: This file takes a text file of 2d barcodes, one per line, as input. For each line, it grabs characters 8 through 15 and looks it up in the associated inkjet file. Then it uses the statement id from the inkjet file to find the record in the out.csn data file. It grabs all of the redo rows from the out.csn file and creates a new one for psl to run.
'''
import argparse, re, os, datetime, shutil
from glob import glob

if __name__ == "__main__":
	
	parser = argparse.ArgumentParser(description='Read a set of 2d barcodes, and find the associated rows in the original data file for the River Valley project', version='%(prog)s 1.2')
	parser.add_argument('infile', nargs='+', type=str, help='input file')
	parser.add_argument('--folder', type=str, default=r'I:\InkJet\RIVER VALLEY', help='inkjet folder to search')
	parser.add_argument('--inkjet', type=str, default='*.csv', help='inkjet file')
	parser.add_argument('--out', type=str, default=r'redos.csn', help='name of output file')
	parser.add_argument('--psl', type=str, default=r'out.csn', help='name of psl data file')
	args = parser.parse_args()

	inkjetIds = []
	inkjetFiles = []

	for f in args.infile:
		for myfile in glob(f):
			with open(myfile, 'rb') as df:
				for line in df:
					inkjetIds.append(line[8:15])
					inkjetFile = line[0:5] + '_sheets.csv'
					if inkjetFile not in inkjetFiles:
						inkjetFiles.append(inkjetFile)

	if len(inkjetFiles) > 1:
		print "There is more than one inkjet file in your barcode files!\n"
		print "quitting...\n"
		sys.exit(1)

	statementIds = []

	for f in inkjetFiles:
		for myfile in glob(os.path.join(args.folder,f)):
			with open(myfile, 'rb') as df:
				for line in df:
					stuff = line.split(',')
					if stuff[0] in inkjetIds:
						statementIds.append(stuff[1])
						inkjetIds.remove(stuff[0])
			dt = datetime.datetime.now()
			shutil.copy(myfile, myfile[:-3]+'_copy_'+str(datetime.datetime.strftime(dt, "%d-%m-%y_%I:%M"))+'.bak')

	with open(args.psl, 'rb') as df:
		with open(args.out, 'wb') as output:
			output.write('StatementID,PageCount,PageNo,AddressBlock,CheckImages\n')
			for line in df:
				stuff = line.split(',')
				if stuff[0] in statementIds:
					output.write(line)
					statementIds.remove(stuff[0])
