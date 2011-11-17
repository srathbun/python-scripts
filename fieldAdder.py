'''
Project: Python skeleton for csv file parsing
Author: Spencer Rathbun
Date: 4/21/2011

Summary: Skeleton file for Rebecca, demonstrating argparse and csv module

'''
import os, re, sys, csv, argparse, win32com.client
from itertools import izip, count

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


	csvfile   = find_csv(csvfile)
	sheetfile = find_sheet(csvfile)

	if sheetfile:
		if not os.path.exists(csvfile):
			convert_sheet(sheetfile, csvfile)
		else:
			if os.stat(csvfile).st_mtime < os.stat(sheetfile).st_mtime:
				sys.stderr.write("CSV file already exists! Do you wish to overwrite it? ")
				input = raw_input()
				if input == 'y' or input == 'Y':
					os.unlink(csvfile)
					convert_sheet(sheetfile, csvfile)
				else:
					sys.stderr.write("Exiting... ")
					sys.exit(0)

	if not os.path.exists(csvfile):
		sys.stderr.write('File not found: %s\n' % csvfile)
		sys.exit(1)

	return csvfile


if __name__ == "__main__":
	
	parser = argparse.ArgumentParser(description='Process a Chapman Kelly affidavit excel file.', version='%(prog)s 1.0')
	parser.add_argument('infile', nargs='+', type=str, help='list of input files')
	parser.add_argument('--out', type=str, default='temp.txt', help='name of output file')
	args = parser.parse_args()

	for f in args.infile:
		with open(f, "rb") as mycsv:
			reader = csv.DictReader(mycsv, dialect='excel')
			fieldnames = reader.fieldnames
			fieldnames.append('COMBO')
			
			writer = csv.DictWriter(open(args.out, "wb"), fieldnames, dialect='excel')
			writer.writeheader()
			
			for row in reader:
				output = row
				output['COMBO'] = row['src_mbr_id'] + '-' + row['prescriber_id']
					
				writer.writerow(row)
	
			del writer
			
	del reader
	
