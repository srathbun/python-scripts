'''
Project: Python barcode and dbf correlator
Author: Spencer Rathbun
Date: 8/3/2011

Summary: Find all of the dbf files from the root directory given, and collect the records listed in the given barcode files. Then print them to a csv file for excel.

'''
import os, sys, csv, argparse 
import dbaseReader
from string import letters, whitespace, punctuation

if __name__ == "__main__":
	
	# variable declarations
	output = []
	filelist = []
	barcodes = []
	# collect command line args
	parser = argparse.ArgumentParser(description='Find all of the dbf files from the root directory given, and collect the records listed in the given barcode files. Then print them to a csv file for excel.', version='%(prog)s 1.0')
	parser.add_argument('barcodeFiles', nargs='+', type=str, help='list of input barcode files, one barcode per line')
	parser.add_argument('-d', type=str, default='.', help='Base directory to find list of dbf files')
	parser.add_argument('--out', type=str, default='temp.csv', help='name of output file for excel')
	args = parser.parse_args()

	# for each barcode file, read it and grab the barcodes, removing any spurious text
	for f in barcodeFiles:
		try:
			with open(f, 'rt') as datafile:
				for line in datafile.readlines():
					barcodes.append(line.translate(None, letters+whitespace+punctuation))
		except Exception, e:
			print >>sys.stderr, e
	# collect a list of all dbf files from the root dir	
	for root,dirs,filenames in os.walk(args.d):
		if root == args.d:
			continue
		# copy all dbf files to the local dir, but only if it is not the local dir
		for filename in filenames:
			if filename.endswith('dbf'):
				shutil.copy(os.path.join(root, filename), '.')

	# read each dbf and check it against our list of barcodes
	for dbf in os.listdir('.'):
		rows = dbaseReader.convertRows(dbf)
		for row in rows:
			if row['IMBRCDDGTS'] in barcodes:
				output.append(row)

	# dump found entries into a csv file for excel
	writer = csv.DictWriter(open(args.out, 'wb'), output[0].keys(), dialect='excel')
	writer.writeheader()
	for row in output:
		writer.writerow(row)
	writer.close()
	del writer
