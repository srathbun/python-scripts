'''
Project: Map multiple excel rows to one csv row
Author: Spencer Rathbun
Date: 2011-11-09

Summary: This project takes a command line input. The input tells it what file to run on.
It creates a new file with the multiple records per mailer converted into single records.

'''
import os
import csv
import argparse
import excelConverter
from collections import defaultdict
from itertools import izip_longest, count

def getPeople(reader, peopleFields):
	"""Find all unique people in the csv file, based on the people fields given."""
	people = defaultdict(list)
	for row, pos in (reader, count(1)):
		try:
			personlist = filter(lambda x: x != None, [row[f] for f in reader.fieldnames if f in peopleFields])
			person = '|'.join(personlist)
			drug = [row[f] for f in reader.fieldnames if f not in peopleFields]
			people[person].append(drug)
		except Exception, e:
			print "Error on row {0}! Printing stack trace and quitting...".format(pos)
			raise e
	return people

if __name__ == "__main__":

	parser = argparse.ArgumentParser(description='Process an excel file, converting multiple rows to one.', version='%(prog)s 1.0')
	parser.add_argument('infile', type=str, help='excel input file')
	args = parser.parse_args()

	myCSV = excelConverter.check_csv_file(args.infile)
	reader = csv.DictReader(open(myCSV, "rb"), dialect='excel')

	peopleFields = []
	drugFields = []
	for field in reader.fieldnames:
		tmp = raw_input("Is field {0} a person field? (y/n) ".format(field))
		if tmp == 'y' or tmp == 'Y':
			peopleFields.append(field)
		else:
			drugFields.append(field)

	people = getPeople(reader, peopleFields)

	fieldnames = peopleFields
	lengths = [len(drugs) for drugs in people.itervalues()]
	for i in range(1, max(lengths)+1):
		fieldnames.extend(['_'.join((field,str(i))) for field in drugFields])

	writer = csv.DictWriter(open("export.csv", "wb"), fieldnames, dialect='excel')
	writer.writeheader()
	for person, drugs in people.iteritems():
		currData = person.split("|")
		for drug in drugs:
			currData.extend(drug)
		currRow = izip_longest(fieldnames, currData, fillvalue='')
		writer.writerow(dict(currRow))

	## clean up and finishing section
	del reader
	del writer
	os.unlink(myCSV)
	try:
		os.unlink(os.path.splitext(args.infile)[0]+"-complete.csv")
	except OSError:
		pass
	os.rename("export.csv", os.path.splitext(args.infile)[0]+"-complete.csv")
