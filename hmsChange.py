'''
Project: Chapman Kelly HMS Change
Author: Spencer Rathbun
Date: 2/7/2011

Summary: This project takes a command line input. The input tells it what file to run on,
what affidavits to look for, and the number of pages in the main letter. It converts the excel file into
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
	
	parser = argparse.ArgumentParser(description='Process a Chapman Kelly affidavit excel file.', version='%(prog)s 1.2')
	parser.add_argument('infile', type=str, help='excel input file')
	args = parser.parse_args()
		
	# read file
	myCSV = check_csv_file(args.infile)
	reader = csv.DictReader(open(myCSV, "rb"), dialect='excel')

	fieldnames = reader.fieldnames
	for i in [2, 3, 4, 5]:
		list = ['target drug med i'+str(i),'target drug ndc'+str(i),'target drug acn'+str(i),'target drug name'+str(i),'target drug tier'+str(i),'brand generic  in'+str(i),'maint ind'+str(i),'claim days supply'+str(i),'lis cd'+str(i),'gpi cd'+str(i),'expr 1'+str(i),'ndc'+str(i),'claim days supply'+str(i),'medid'+str(i),'gcn'+str(i),'drug name'+str(i),'label name'+str(i),'alternative 1'+str(i),'alternative 2'+str(i),'alternative 3'+str(i),'alternate 4'+str(i)]
		fieldnames.extend(list)
	#print fieldnames

	writer = csv.DictWriter(open("export.csv", "wb"), fieldnames, dialect='excel')
	writer.writeheader()

	if "member id" in reader.fieldnames:
		idCode = "member id"
	else:
		sys.exit('No Identifiable employee id field!')

#Target_Drug_MedID
#Target_Drug_NDC
#Target_Drug_GCN
#Target_Drug_Name
#Target_Drug_Tier
#Brand_Generic_Ind
#Maint_Ind
#Claim_Days_Supply
#LIS_CD
#GPI_CD
#Expr1

	currEID = ''
	i = 2
	currRow = {}
	longest = 0

	try:		
		for row in reader:
			if currEID != row[idCode]:
				currEID = row[idCode]
				writer.writerow(currRow)
				currRow = row
				#if i > longest:
				#	longest = i
				i = 2
			else:
				#currRow['dep_client_value_'+str(i)] = row['dep_client_value_1']
				currRow['target drug med i'+str(i)] = row['target drug med i']
				currRow['target drug ndc'+str(i)] = row['target drug ndc']
				currRow['target drug acn'+str(i)] = row['target drug acn']
				currRow['target drug name'+str(i)] = row['target drug name']
				currRow['target drug tier'+str(i)] = row['target drug tier']
				currRow['brand generic  in'+str(i)] = row['brand generic  in']
				currRow['maint ind'+str(i)] = row['maint ind']
				currRow['claim days supply'+str(i)] = row['claim days supply']
				currRow['lis cd'+str(i)] = row['lis cd']
				currRow['gpi cd'+str(i)] = row['gpi cd']
				currRow['expr 1'+str(i)] = row['expr 1']
				currRow['ndc'+str(i)] = row['ndc']
				currRow['claim days supply'+str(i)] = row['claim days supply']
				currRow['medid'+str(i)] = row['medid']
				currRow['gcn'+str(i)] = row['gcn']
				currRow['drug name'+str(i)] = row['drug name']
				currRow['label name'+str(i)] = row['label name']
				currRow['alternative 1'+str(i)] = row['alternative 1']
				currRow['alternative 2'+str(i)] = row['alternative 2']
				currRow['alternative 3'+str(i)] = row['alternative 3']
				currRow['alternate 4'+str(i)] = row['alternate 4']
				i = i+1
	except csv.Error, e:
		sys.exit('file %s, line %d: %s' % (filename, reader.line_num, e))
	except Exception, e:
		sys.exit('Python error: %s' % e)
	writer.writerow(currRow)

	# clean up variables and delete temp files
	del reader
	del writer
	#os.unlink(myCSV)
	try:
		os.unlink(os.path.splitext(args.infile)[0]+"-affidavits.csv")
	except:
		pass
	os.rename("export.csv", os.path.splitext(args.infile)[0]+"-EDIT.csv")
