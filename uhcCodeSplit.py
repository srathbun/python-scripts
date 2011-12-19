'''
Project: uhc code splitter
Author: Spencer Rathbun
Date: 12/12/2011

Summary: Split uhc audit file based upon list of input codes

'''
import os, csv, argparse

if __name__ == "__main__":
	parser = argparse.ArgumentParser(description='Process an UHC audit file.', version='%(prog)s 1.1')
	parser.add_argument('infile', type=str, help='input filename')
	parser.add_argument('codes', nargs='+', type=str, help='list of letter codes, such as Z1220 C1330 B1490')
	parser.add_argument('--out', type=str, default='-good.csv', help='appended name of good entries')
	parser.add_argument('--bad', type=str, default='-bad.csv', help='appended name of stripped entries')
	args = parser.parse_args()
	
	codes = [x.upper() for x in args.codes]

	with open(args.infile, "rb") as mycsv:
		reader = csv.reader(mycsv)
		goodWriter = csv.writer(open(os.path.splitext(args.infile)[0]+args.out, "wb"), quoting=csv.QUOTE_ALL)
		badWriter = csv.writer(open(os.path.splitext(args.infile)[0]+args.bad, "wb"), quoting=csv.QUOTE_ALL)
		
		for row in reader:
			if row[1]+row[2] in codes:
				badWriter.writerow(row)
			else:
				goodWriter.writerow(row)
		del goodWriter
		del badWriter
	del reader
