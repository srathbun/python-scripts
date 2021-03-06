'''
Project: reader
Author: Spencer Rathbun
Date: 3/24/2011

Summary: This project takes a command line input. The input tells it what file to run on.
It converts the excel file into
a csv file, parses it, and outputs a csv file. The output consists of each EID from the original file,
with the number of affidavits, followed by the sheet count. 

'''
import os
import re
import sys
import csv
import argparse
import win32com.client

def check_csv_file(csvfile):
	"""
	Make sure the CSV file exists.
	If it does not exist, try to create it from an XLS or ODS file.
	The passed name can refer to an XLS file or an ODS file.
	The return value will be the name of a CSV file.
	"""
	def find_csv(csvf):
		""" Check to see if the CSV file exists. """
		fname, ext = os.path.splitext(csvf)
		ext        = ext.lower()
		if   ext == '.txt':
			f = '%s.%s' % (fname, 'TXT')
			if os.path.exists(f):
				csvf = f
			else:
				csvf = '%s.%s' % (fname, 'txt')

		elif ext != '.csv':
			f = '%s.%s' % (fname, 'CSV')
			if os.path.exists(f):
				csvf = f
			else:
				csvf = '%s.%s' % (fname, 'csv')

		return csvf

	def find_sheet(ssvf):
		""" Check to see if the spread sheet file exists. """
		sheetf = None
		fname  = os.path.splitext(ssvf)[0]
		for ext in ('xls', 'XLS', 'xlsx', 'XLSX', 'ods', 'ODS'):
			f = '%s.%s' % (fname, ext)
			if os.path.exists(f):
				sheetf = f
				break

		return sheetf

	def convert_sheet(sheetf, csvf):
		""" Convert spreadsheet to a CSV file. """
		try:
			excel = win32com.client.Dispatch('Excel.Application')
			absPath = os.path.abspath(sheetf)
			location = os.path.split(absPath)[0]
			workbook = excel.Workbooks.Open(absPath)
			
			saveName = location + "\\" + os.path.split(csvf)[1]
			workbook.SaveAs(saveName, FileFormat=24) # 24 represents xlCSVMSDOS
			workbook.Close(False)
			excel.Quit()
			del excel
		except Exception, e:
			sys.stderr.write("ERROR: %s\n" % e)
			workbook.Close(False)
			excel.Quit()
			del excel
			sys.exit(1)


	# Find the spreadsheet file that corresponds to the CSV file
	csvfile   = find_csv(csvfile)
	sheetfile = find_sheet(csvfile)

	if sheetfile:
		# If CSV does not exist try to create it.
		if not os.path.exists(csvfile):
			#sys.stderr.write("Creating %s from %s\n" % (csvfile, sheetfile))
			convert_sheet(sheetfile, csvfile)
		else:
			# If spreadsheet is newer the CSV file, re-create it.
			if os.stat(csvfile).st_mtime < os.stat(sheetfile).st_mtime:
				sys.stderr.write("CSV file already exists! Do you wish to overwrite it? ")
				input = raw_input()
				if input == 'y' or input == 'Y':
					#sys.stderr.write("Recreating %s from %s\n" % (csvfile, sheetfile))
					os.unlink(csvfile)
					convert_sheet(sheetfile, csvfile)
				else:
					sys.stderr.write("Exiting... ")
					sys.exit(0)

	if not os.path.exists(csvfile):
		sys.stderr.write('File not found: %s\n' % csvfile)
		sys.exit(1)

	return csvfile

def check_relationship(string):
	if (len(string) != 2) or (re.match(r"[a-zA-Z][a-zA-Z]", string) == None):
		msg = "%r must consist of two letters" % string
		raise argparse.ArgumentTypeError(msg)
	return string
	
if __name__ == "__main__":
	
	parser = argparse.ArgumentParser(description='Process a Chapman Kelly affidavit excel file.', version='%(prog)s 1.1')
	parser.add_argument('infile', type=str, help='excel input file')
	parser.add_argument('pagefile', type=str, help='page count file')
	args = parser.parse_args()

	myCSV = check_csv_file(args.infile)
	
	filepages = {}
	with open(args.pagefile, 'rb') as f:
		reader = csv.reader(f, dialect='excel-tab')
		for row in reader:
			filepages[row[0]] = row[1]
	
	reader = csv.DictReader(open(myCSV, "rb"), dialect='excel')
		
	for row in reader:
		try:
			if int(filepages[row['PacketID']])/2 != int(row['Pages']):
				print row['PacketID'], row['Pages'], int(filepages[row['PacketID']])/2
		except KeyError,e:
			print "%s does not exist in sample directory!" % e

	del reader
	os.unlink(myCSV)