'''
Project: Python skeleton for csv file parsing
Author: Spencer Rathbun
Date: 4/21/2011

Summary: Skeleton file for Rebecca, demonstrating argparse and csv module

'''
import os, re, sys, csv, argparse, win32com.client
from itertools import izip, count, groupby


if __name__ == "__main__":
	
	# sets up command line arguments, this allows us to input our files and switches
	parser = argparse.ArgumentParser(description='Process a pcBuilding report for psl.', version='%(prog)s 1.1')
	parser.add_argument('infile', nargs='+', type=str, help='list of input files')
	parser.add_argument('--std', type=str, default='standard.txt', help='name of output file for standard sets')
	parser.add_argument('--large', type=str, default='nineByTwelve.txt', help='name of output file for large sets')
	# get the current arguments and put them into a variable
	args = parser.parse_args()

	# since we have infile set to nargs='+', we get a list of file names
	# so we create a for loop to get each file in the list
	for f in args.infile:
		# open the file in read binary mode, with the name mycsv
		with open(f, "rb") as mycsv:
			reader = csv.DictReader(mycsv, dialect='excel-tab')
			fieldnames = ['PDF', 'PageInPDF', 'CurrPage', 'TotalPages']
			writer = csv.DictWriter(open(args.std, "wb"), fieldnames, dialect='excel-tab')
			writerLarge = csv.DictWriter(open(args.large, "wb"), fieldnames, dialect='excel-tab')
			# write the header at the top of the file
			writer.writeheader()
			writerLarge.writeheader()
			# now our file is ready for writing
			mylist = []
			i = 0
			for row in reader:
				if row['Field3'].find('CLOSING') != -1:
					i = i + 1
				row['group'] = i
				mylist.append(row.copy())
				
			mylist.sort(key=lambda x: x['group'])
			
			for key, group in groupby(mylist, lambda x: x['group']):
				temp = list(group)
				counted = len(temp)
				for item,current in izip(temp, count(1)):
					item['CurrPage'] = "%02d" % current
					item['TotalPages'] = "%02d" % counted
			
			mylist.sort(key=lambda x: int(x['TotalPages']))
						
			for row in mylist:
				output = {'PDF':row['Field1'], 'PageInPDF':row['Field2'], 'CurrPage':row['CurrPage'], 'TotalPages':row['TotalPages'] }
				if int(output['TotalPages']) > 4:
					writerLarge.writerow(output)
				else:
					writer.writerow(output)
	
			
			
	## clean up and finishing section
	del reader
	del writer
