#!/usr/bin/env python
import re, csv

page = False
checksOnPage = False
pages = {}
pageNo = 0
startAddressBlock = False
addressBlock = []
with open("test4 statement file.ps", 'rb') as df:
	for line in df.readlines():
		if line.find('%%Page:') != -1:
			page = True
			pageNo = pageNo + 1
		elif line.find('%%PageTrailer') != -1:
			page = False
			try:
				pages[account][0] = pages[account][0] + 1
				pages[account][1].append(str(pageNo))
			except KeyError, e:
				# [totalPages, WhichPages, addressBlock, checkPages]
				pages[account] = [1, [str(pageNo)], '', 0]
			if checksOnPage:
				pages[account][3] = pages[account][3] + 1
				checksOnPage = False
			if not pages[account][2]:
				pages[account][2] = ' '.join(addressBlock)
				addressBlock = []
		matches = re.findall(r'^\(.*\)', line)
		for match in matches:
			if re.match(r'\(\s*[ADTF]{20,65}\)', match):
				startAddressBlock = False
			else:
				words = match[1:-1].upper().split()
				if startAddressBlock:
					addressBlock.append(match[1:-1].strip())
				try:
					if words[0] == 'PRIMARY' and words[1] == 'ACCOUNT':
						account = words[2]
					elif 'CK' in words and 'NUM:' in words:
						checksOnPage = True
					elif words[2] == 'PAGE' and words[3] == '1':
						startAddressBlock = True
				except Exception, e:
					pass

# after parsing the postscript file, write our data to a csv file
writer = csv.writer(open("out.csv", 'wb'), dialect="excel")
writer.writerow(["Statement ID", "Page Count", "Page No", "Address Block"])

for name,item in pages.iteritems():
	row = []
	row.append(name)
	row.append(item[0]) # total pages
	row.append(':'.join(item[1])) # pages in original
	row.append(item[2]) # address block
	row.append(item[3]) # check pages
	writer.writerow(row)
del writer
