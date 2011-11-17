'''
Project: hms reader
Author: Spencer Rathbun
Date: 3/4/2011

Summary: This project takes a command line input. The input tells it what file to run on.
It creates a new file with the multiple records per mailer converted into single records, based on the EID.

'''
import os
import sys
import csv
import argparse
import excelConverter

if __name__ == "__main__":
	
	parser = argparse.ArgumentParser(description='Process a Chapman Kelly affidavit excel file.', version='%(prog)s 1.1')
	parser.add_argument('infile', type=str, help='excel input file')
	args = parser.parse_args()

	myCSV = excelConverter.check_csv_file(args.infile)
	reader = csv.DictReader(open(myCSV, "rb"), dialect='excel')
	
	fieldnames = reader.fieldnames
	for i in [2, 3, 4, 5, 6, 7, 8, 9, 10]:
		list = ['dep_client_value_'+str(i),'dep_name_'+str(i),'dep_rel_code_'+str(i),'dep_rel_'+str(i),'dep_dob_'+str(i),'products_'+str(i)]
		fieldnames.extend(list)
	#print fieldnames
	
	writer = csv.DictWriter(open("export.csv", "wb"), fieldnames, dialect='excel')
	#output = ['EID','AFFIDAVITS','SHEETS']
	writer.writeheader()
	#writer.writerow(fieldnames)

	if "eid" in reader.fieldnames:
		idCode = "eid"
	elif "EID" in reader.fieldnames:
		idCode = "EID"
	elif "EMPLOYEEID" in reader.fieldnames:
		idCode = "EMPLOYEEID"
	elif "EMPLOYEE_ID" in reader.fieldnames:
		idCode = "EMPLOYEE_ID"
	elif "employee_id" in reader.fieldnames:
		idCode = "employee_id"
	elif "employeeid" in reader.fieldnames:
		idCode = "employeeid"
	else:
		sys.exit('No Identifiable employee id field!')

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
				currRow['dep_client_value_'+str(i)] = row['dep_client_value_1']
				#currRow['dep_client_id_'+str(i)] = row['dep_client_id']
				currRow['dep_name_'+str(i)] = row['dep_name']
				currRow['dep_rel_code_'+str(i)] = row['dep_rel_code']
				currRow['dep_rel_'+str(i)] = row['dep_rel'] 
				currRow['dep_dob_'+str(i)] = row['dep_dob'] 
				#currRow['products_'+str(i)] = row['products']
				i = i+1
	except csv.Error, e:
		sys.exit('line %d: %s' % (reader.line_num, e))
	except Exception, e:
		sys.exit('Python error: %s' % e)
	writer.writerow(currRow)
	#print longest
		
	## clean up and finishing section
	del reader
	del writer
	os.unlink(myCSV)
	try:
		os.unlink(os.path.splitext(args.infile)[0]+"-affidavits.csv")
	except:
		pass
	os.rename("export.csv", os.path.splitext(args.infile)[0]+"-complete.csv")
