#!/usr/bin/env python
'''
Project: pageSwitch.py
Author: Spencer Rathbun
Date: 8/25/2011

Summary: Add commands to switch trays on the heidelberg printers. Note that using the Kodak print file downloader will append a command to the header that causes the printer to ignore tray switching commands!

'''
import argparse, re, os
from glob import glob

if __name__ == "__main__":
	
	parser = argparse.ArgumentParser(description='Parse over a postscript file and set the MediaType for each page.', version='%(prog)s 1.1')
	parser.add_argument('infile', nargs='+', type=str, help='input file')
	parser.add_argument('--out', type=str, default='_tray_switch.ps', help='name of output file')
	args = parser.parse_args()

	page = False
	checksOnPage = False
	linesOnPage = []


	for f in args.infile:
		for myfile in glob(f):
			output = open(os.path.splitext(myfile)[0]+args.out, 'wb')
			with open(myfile, 'rb') as df:
				for line in df:
					if line.find('%%Page:') != -1:
						page = True
					elif line.find('%%PageTrailer') != -1:
						page = False

					if page:
						if line.find('%%BeginBinary:') != -1:
							checksOnPage = True
						linesOnPage.append(line)
#					matches = re.findall(r'^\(.*\)', line)
#					for match in matches:
#						words = match[1:-1].upper().split()
#						try:
#							if 'CK' in words and 'NUM:' in words:
#								checksOnPage = True
#						except Exception, e:
#							pass
					else:
						if linesOnPage:
							'''letterhead is used for standard pages, but check pages are printed on plain'''
							if checksOnPage:
								linesOnPage.insert(linesOnPage.index("%%BeginPageSetup\r\n")+1, "<<  /MediaColor (white) /MediaType (plain)>> setpagedevice\r\n")
							else:
								linesOnPage.insert(linesOnPage.index("%%BeginPageSetup\r\n")+1, "<<  /MediaColor (white) /MediaType (letterhead)>> setpagedevice\r\n")
							for entry in linesOnPage:
								output.write(entry)
							linesOnPage = []
							checksOnPage = False
						output.write(line)

			output.close()
