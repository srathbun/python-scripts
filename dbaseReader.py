#!/usr/bin/env python
'''
File: dbaseReader.py
Author: Spencer Rathbun
Description: Convert dbf files to a python list of dictionaries.
dbfreader function from: http://code.activestate.com/recipes/362715/
'''

import dateAdditionLib
import struct, datetime
from itertools import izip
from decimal import *
import sys, csv, argparse


def dbfreader(f):
    """Returns an iterator over records in a Xbase DBF file.

    The first row returned contains the field names.
    The second row contains field specs: (type, size, decimal places).
    Subsequent rows contain the data records.
    If a record is marked as deleted, it is skipped.

    File should be opened for binary reads.

    """
    # See DBF format spec at:
    #     http://www.pgts.com.au/download/public/xbase.htm#DBF_STRUCT

    numrec, lenheader = struct.unpack('<xxxxLH22x', f.read(32))    
    numfields = (lenheader - 33) // 32

    fields = []
    fields.insert(0, ('DeletionFlag', 'C', 1, 0))
    for fieldno in xrange(numfields):
        name, typ, size, deci = struct.unpack('<11sc4xBB14x', f.read(32))
        name = name.replace('\0', '')       # eliminate NULs from string   
        fields.append((name, typ, size, deci))
    yield [field[0] for field in fields]
    yield [tuple(field[1:]) for field in fields]

    terminator = f.read(1)
    assert terminator == '\r'

    fmt = ''.join(['%ds' % fieldinfo[2] for fieldinfo in fields])
    fmtsiz = struct.calcsize(fmt)
    for i in xrange(numrec):
        record = struct.unpack(fmt, f.read(fmtsiz))
        #if record[0] != ' ':
        #    continue                        # deleted record
        result = []
        for (name, typ, size, deci), value in izip(fields, record):
            if name == 'DeletionFlag':
				pass
				#continue
            if typ == "N":
                value = value.replace('\0', '').lstrip()
                if value == '':
                    value = 0
                elif deci:
                    value = Decimal(value)
                else:
                    value = int(value)
            elif typ == 'D':
                #y = int(str(value[:4]))
                #y, m, d = int(value[:4].encode('ascii')), int(value[4:6].encode('ascii')), int(value[6:8].encode('ascii'))
                #y, m, d = value[:4], value[4:6],value[6:8]
                #value = datetime.date(y, m, d)
                value = value[:4] + '/' + value[4:6] + '/' + value[6:8]
            elif typ == 'L':
                value = (value in 'YyTt' and 'T') or (value in 'NnFf' and 'F') or '?'
            elif typ == 'F':
                value = float(value)
            result.append(value)
        yield result


def convertRows(f, hidden):
	with open(f, 'rb') as fd:
		db = list(dbfreader(fd))
#	for record in db:
#		print record
		fieldnames, fieldspecs, records = db[0], db[1], db[2:]
		rows = []
		for record in records:
			zipped = zip(fieldnames, record)
			dictionary = dict(zipped)
			#rows.append(dictionary)
			if hidden:
				rows.append(dictionary)
			elif dictionary['DeletionFlag'] != '*':
				rows.append(dictionary)
			else:
				print dictionary['DeletionFlag']
		return rows

def column(filename, data, indent=0):
	"""This function takes string input and produces columized output"""
	# get the width of the columns
	width = []
	for mylist in data:
		for count, d in enumerate(mylist):
			if count > (len(width)-1):
				width.append(len(str(d)))
			elif len(str(d)) > width[count]:
				width[count] = len(str(d))
	
	# print the data
	for mylist in data:
		line = '{0:<{indent}}'.format('', indent=indent)	
		for count, d in enumerate(mylist):
			try:
				line = '%s%s' % (line, '{0:{w},} '.format(d, w=width[count]))
			except ValueError, e:
				line = '%s%s' % (line, '{0:{w}} '.format(d, w=width[count]))
		filename.write(line)
		filename.write('\n')



if __name__ == "__main__":
	
	# sets up command line arguments, this allows us to input our files and switches
	parser = argparse.ArgumentParser(description='BCC dbf file reader for producing trover report', version='%(prog)s 1.0')
	parser.add_argument('today', type=str, help='trover file for today')
	parser.add_argument('-s', type=bool, default=True, help='Skip hidden Records')
	parser.add_argument('--out', type=str, default='temp.csv', help='name of output file')
	# get the current arguments and put them into a variable
	args = parser.parse_args()


	rows = convertRows(args.today, args.s)
	writer = csv.DictWriter(open(args.out, 'wb'), rows[0].keys(), dialect='excel')
	writer.writeheader()
	for row in rows:
		writer.writerow(row)
	'''
	todaysRows = convertRows(args.today)
	yesterdaysRows = convertRows(args.yesterday)

	# now that we have loaded the data, we need to collect the various numbers we are interested in
	totalRecs = len(filter(lambda x: x['LTTRSTTS'].strip() == 'ELIGIBLE FOR PRINT', iter(todaysRows)))
	totalNCOA = len(filter(lambda x: x['TEST_FLAG'].strip() == 'N', iter(yesterdaysRows)))
	totalRejects = len(filter(lambda x: x['TEST_FLAG'].strip() == 'N' and x['DSPSTNCD'].strip() in ['B','C','D'], iter(yesterdaysRows)))

	totalSent = len(filter(lambda x: x['TEST_FLAG'].strip() == 'N', iter(todaysRows)))
	testRecs = len(filter(lambda x: x['TEST_FLAG'].strip() == 'Y', iter(todaysRows)))

	# test record categories
	testCert = len(filter(lambda x: x['TEST_FLAG'].strip() == 'Y' and x['PSTGCLMN'].strip().upper() == 'CERT', iter(todaysRows)))
	testDuplex = len(filter(lambda x: x['TEST_FLAG'].strip() == 'Y' and x['PSTGCLMN'].strip().upper() == 'DUPLEX', iter(todaysRows)))
	testLetter1_3 = len(filter(lambda x: x['TEST_FLAG'].strip() == 'Y' and x['PSTGCLMN'].strip().upper()== '1ST AND 3RD NOTICE', iter(todaysRows)))
	testLetter2 = len(filter(lambda x: x['TEST_FLAG'].strip() == 'Y' and x['PSTGCLMN'].strip().upper() == '2ND NOTICE', iter(todaysRows)))
	testSharps = len(filter(lambda x: x['TEST_FLAG'].strip() == 'Y' and x['PSTGCLMN'].strip().upper() == 'SHARPS', iter(todaysRows)))
	testSharpsM = len(filter(lambda x: x['TEST_FLAG'].strip() == 'Y' and x['PSTGCLMN'].strip().upper() == 'SHARPS-MULTI', iter(todaysRows)))
	testIBCP = len(filter(lambda x: x['TEST_FLAG'].strip() == 'Y' and x['PSTGCLMN'].strip().upper() == 'IBCP', iter(todaysRows)))

	# record categories
	Cert = len(filter(lambda x: x['TEST_FLAG'].strip() == 'N' and x['PSTGCLMN'].strip().upper() == 'CERT' and x['DSPSTNCD'].strip() not in ['B','C','D'], iter(yesterdaysRows)))
	Duplex = len(filter(lambda x: x['TEST_FLAG'].strip() == 'N' and x['PSTGCLMN'].strip().upper() == 'DUPLEX' and x['DSPSTNCD'].strip() not in ['B','C','D'], iter(yesterdaysRows)))
	Letter1_3 = len(filter(lambda x: x['TEST_FLAG'].strip() == 'N' and x['PSTGCLMN'].strip().upper()== '1ST AND 3RD NOTICE' and x['DSPSTNCD'].strip() not in ['B','C','D'], iter(yesterdaysRows)))
	Letter2 = len(filter(lambda x: x['TEST_FLAG'].strip() == 'N' and x['PSTGCLMN'].strip().upper() == '2ND NOTICE' and x['DSPSTNCD'].strip() not in ['B','C','D'], iter(yesterdaysRows)))
	Sharps = len(filter(lambda x: x['TEST_FLAG'].strip() == 'N' and x['PSTGCLMN'].strip().upper() == 'SHARPS' and x['DSPSTNCD'].strip() not in ['B','C','D'], iter(yesterdaysRows)))
	SharpsM = len(filter(lambda x: x['TEST_FLAG'].strip() == 'N' and x['PSTGCLMN'].strip().upper() == 'SHARPS-MULTI' and x['DSPSTNCD'].strip() not in ['B','C','D'], iter(yesterdaysRows)))
	IBCP = len(filter(lambda x: x['TEST_FLAG'].strip() == 'N' and x['PSTGCLMN'].strip().upper() == 'IBCP' and x['DSPSTNCD'].strip() not in ['B','C','D'], iter(yesterdaysRows)))

	with open(args.out, 'wt') as output:
		output.write('{:%m/%d/%y} TROVERIS COUNTS REPORT\n'.format(datetime.date.today()))
		output.write('FOR MAIL DATE: {:%m/%d/%y}\n'.format(dateAdditionLib.workdayadd(datetime.date.today(), 1)))

		data = [['# OF RECORDS', 'FROM TROVERIS', totalRecs],
		['', 'FROM NCOA', totalNCOA],
		['','','----------'],
		['','',(totalRecs + totalNCOA), 'TOTAL RECORDS IN']]
		column(output, data, 0)
		output.write('\n\n')

		data = [['TO NCOA', totalSent],
				['REJECTS/UNDELIVERABLES', totalRejects, '{:.2%}'.format(1.00*totalRejects/(totalRecs+totalNCOA)), 'OF TOTAL RECORDS IN' ]]
		column(output, data, 4)
		# write the number of test letters in each category,
		output.write('TEST DATA:\n')

		data = [['CERT. LETTER' , testCert],
		['DUPLEX LETTER' , testDuplex],
		['LETTERS 1 & 3' , testLetter1_3],
		['LETTER # 2' , testLetter2],
		['SHARPS LETTER' , testSharps],
		['SHARPS LETTER M' , testSharpsM],
		['IBCP' , testIBCP],
		['', '----------'],
		['', (totalRecs+totalRejects), 'TOTAL RECORDS NOT MAILED']]
		column(output, data, 4)
		output.write('\n\n')

		# write the number of actual letters received in each category,
		data = [['','','REKEYS'],
		['CERT. LETTER' , Cert, 0, Cert],
		['DUPLEX LETTER' , Duplex, 0, Duplex],
		['LETTERS 1 & 3' , Letter1_3, 0, Letter1_3],
		['LETTER # 2' , Letter2, 0, Letter2],
		['SHARPS LETTER' , Sharps, 0, Sharps],
		['SHARPS LETTER M' , SharpsM, 0, SharpsM],
		['IBCP' , IBCP, 0, IBCP],
		['', '-------', '-------', '-------'],
		['', (Cert+Duplex+Letter1_3+Letter2+Sharps+SharpsM+IBCP), 0, (Cert+Duplex+Letter1_3+Letter2+Sharps+SharpsM+IBCP), 'TOTAL RECORDS MAILED']]
		column(output, data, 4)
		data = [['', '-------'],
		['               ', (totalRecs+totalRejects+Cert+Duplex+Letter1_3+Letter2+Sharps+SharpsM+IBCP), 'TOTAL RECORDS OUT' ]]
		column(output, data, 4)
		'''
