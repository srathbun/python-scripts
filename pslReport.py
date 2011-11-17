#!/usr/bin/env python
'''
Project: psl report for tray switching
Author: Spencer Rathbun
Date: 8/25/2011

Summary: This program reads through a postscript file and determines how many pages there are per statement, and which pages are check images. It outputs this information into a psl report, and also makes a page count report and a zip distribution report.

'''
import argparse, re, csv
from itertools import count, imap
from collections import Counter

def moreThanEight(key, value):
	if key > 8:
		return value
	else:
		return 0

if __name__ == "__main__":
	
	parser = argparse.ArgumentParser(description='Parse over a postscript file and create a report for psl', version='%(prog)s 1.2.1')
	parser.add_argument('infile', type=str, help='input file')
	parser.add_argument('outfile', nargs='?', default='output_for_converting_to_pdf.ps', type=str, help='output file')
	parser.add_argument('--out', type=str, default='out.csn', help='name of output file')
	parser.add_argument('-z', type=str, default='zipDistributionReport.txt', help='name of zip distribution report file')
	parser.add_argument('-p', type=str, default='pageCountReport.txt', help='name of page count report file')
	args = parser.parse_args()

	page = False
	checksOnPage = False
	pages = {}
	pageNo = 0
	startAddressBlock = False
	statementNum = 1
	addressBlock = []
	with open(args.infile, 'rb') as df:
		with open(args.outfile, 'wb') as output:
			for line in df:
				if line.find('%%Page:') != -1:
					pageNo = pageNo + 1
				elif line.find('%%BeginBinary:') != -1:
					checksOnPage = True
				elif line.find('%%PageTrailer') != -1:
					if page:
						try:
							pages[account][0] = pages[account][0] + 1
							pages[account][1].append(str(pageNo))
						except KeyError, e:
							# [totalPages, WhichPages, addressBlock, checkPages]
							pages[account] = [1, [str(pageNo)], '', []]
						if checksOnPage:
							pages[account][3].append(str(pageNo))
							checksOnPage = False
						if not pages[account][2]:
							pages[account][2] = ' '.join(addressBlock)
							addressBlock = []
					page = False
				matches = re.findall(r'^\(.*\)', line)
				if matches:
					page = True
				for match in matches:
					if re.match(r'\(\s*[ADTF]{20,65}\)', match):
						startAddressBlock = False
						line = re.sub(r'\(\s*[ADTF]{20,65}\)', '()', line)
					else:
						words = match[1:-1].upper().split()
						if startAddressBlock:
							addressBlock.append(match[1:-1].strip())
						try:
							if words[0] == 'PRIMARY' and words[1] == 'ACCOUNT':
								account = ''.join((words[2], str(statementNum)))
							elif 'CK' in words and 'NUM:' in words:
								checksOnPage = True
							elif words[2] == 'PAGE' and words[3] == '1':
								statementNum = statementNum + 1
								startAddressBlock = True
						except IndexError, e:
							pass
				output.write(line)


# after parsing the postscript file, write our data to a csv file
	writer = csv.writer(open(args.out, 'wb'), dialect="excel")
	writer.writerow(["StatementID", "PageCount", "PageNo", "AddressBlock", "CheckImages", "RowNumber"])
	bad = csv.writer(open('bad.csn', 'wb'), dialect="excel")
	bad.writerow(["StatementID", "PageCount", "PageNo", "AddressBlock", "CheckImages", "RowNumber"])


	letterheadCount8page = 0
	whiteCount8page = 0
	letterheadCount9page = 0
	whiteCount9page = 0
	for item, rowNumber in zip(sorted(pages.iteritems(), key=lambda x: x[1][0]), count(1)):
		row = []
		row.append(item[0]) # this is the name of the item after sorting into tuple form
		row.append(item[1][0]) # total pages
		if item[1][0] < 9:
			letterheadCount8page = letterheadCount8page + item[1][0] - len(item[1][3])
			whiteCount8page = whiteCount8page + len(item[1][3])
		else:
			letterheadCount9page = letterheadCount9page + item[1][0] - len(item[1][3])
			whiteCount9page = whiteCount9page + len(item[1][3])
		row.append(':'.join(item[1][1])) # pages in original
		row.append(item[1][2]) # address block
		row.append(':'.join(item[1][3])) # check pages
		row.append(rowNumber)
		if item[1][2].find('*************') != -1:
			print "Found bad statement! Added to bad.csn"
			bad.writerow(row)
		else:
			writer.writerow(row)
	del writer
	del bad

	# create zip distribution report
	data = {}
	for name,item in pages.iteritems():
		try:
			data[item[2].split()[-1][:5]] = data[item[2].split()[-1][:5]] + 1
		except Exception, e:
			data[item[2].split()[-1][:5]] = 1

	inCount = 0
	kyCount = 0
	otherCount = 0
	writer = open(args.z, 'wt')
	writer.write("ZIP\tNumber of occurances\n")
	for key, value in data.iteritems():
		writer.write('%s: %s\n' % (str(key), str(value)))
		if 45900 < key < 48000:
			inCount = inCount + value
		elif 40000 < key < 41900:
			kyCount = kyCount + value
		elif 42000 < key < 42800:
			kyCount = kyCount + value
		else:
			otherCount = otherCount + value
	writer.write('\n')
	writer.write('State Report:\n')
	writer.write('IN count: {0}\n'.format(inCount))
	writer.write('KY count: {0}\n'.format(kyCount))
	writer.write('Other count: {0}\n'.format(otherCount))
	writer.close()

	# create page count report
	counts = Counter(value[0] for value in pages.itervalues())
	writer = open(args.p, 'wt')
	writer.write("Page Count\tNumber of occurances\n")
	for key, value in counts.iteritems():
		writer.write('%s: %s\n' % (str(key), str(value)))
	writer.write('\n')
	writer.write('Number of records more than 8 pages: {0}\n'.format(sum(imap(moreThanEight, counts.iterkeys(), counts.itervalues()))))
	writer.write('\n')
	writer.write('Letterhead count for 8 sheets or less: {0}\n'.format(letterheadCount8page))
	writer.write('White paper count for 8 sheets or less: {0}\n'.format(whiteCount8page))
	writer.write('\n')
	writer.write('Letterhead count for 9 sheets or more: {0}\n'.format(letterheadCount9page))
	writer.write('White paper count for 9 sheets or more: {0}\n'.format(whiteCount9page))
	writer.write('\n')
	writer.write('Total letterhead needed: {0}\n'.format(letterheadCount9page+letterheadCount8page))
	writer.write('Total white paper needed: {0}\n'.format(whiteCount9page+whiteCount8page))
	writer.close()
