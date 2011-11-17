'''
Project: Python skeleton for csv file parsing
Author: Spencer Rathbun
Date: 4/21/2011

Summary: Skeleton file for Rebecca, demonstrating argparse and csv module

'''
import os, csv, argparse, win32com.client
from HTMLParser import HTMLParser
import htmlentitydefs

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

class d(csv.Dialect):
	escapechar = chr(92)
	doublequote = False
	delimiter = ','
	quotechar = '"'
	skipinitialspace = False
	lineterminator = '\r\n'
	quoting = csv.QUOTE_MINIMAL

class BaseHTMLProcessor(HTMLParser):
    def reset(self):                       
        # extend (called by HTMLParser.__init__)
        self.pieces = []
        HTMLParser.reset(self)

    def handle_starttag(self, tag, attrs):
        # called for each start tag
        # attrs is a list of (attr, value) tuples
        # e.g. for <pre class="screen">, tag="pre", attrs=[("class", "screen")]
        # Ideally we would like to reconstruct original tag and attributes, but
        # we may end up quoting attribute values that weren't quoted in the source
        # document, or we may change the type of quotes around the attribute value
        # (single to double quotes).
        # Note that improperly embedded non-HTML code (like client-side Javascript)
        # may be parsed incorrectly by the ancestor, causing runtime script errors.
        # All non-HTML code must be enclosed in HTML comment tags (<!-- code -->)
        # to ensure that it will pass through this parser unaltered (in handle_comment).
		if tag == 'b': 
			v = r'%b[1]'
		elif tag == 'li': 
			v = r'%f[1]'
		elif tag == 'strong': 
			v = r'%b[1]%i[1]'
		elif tag == 'u': 
			v = r'%u[1]'
		elif tag == 'ul': 
			v = r'%n%'
		else:
			v = ''
		self.pieces.append("{0}".format(v))

    def handle_endtag(self, tag):         
        # called for each end tag, e.g. for </pre>, tag will be "pre"
        # Reconstruct the original end tag.
		if tag == 'li': 
			v = r'%f[0]' 
		elif tag == '/b': 
			v = r'%b[0]'
		elif tag == 'strong': 
			v = r'%b[0]%i[0]'
		elif tag == 'u': 
			v = r'%u[0]'
		elif tag == 'ul': 
			v = ''
		elif tag == 'br': 
			v = r'%n%' 
		else: 
			v = '' # it matched but we don't know what it is! assume it's invalid html and strip it
		self.pieces.append("{0}".format(v))

    def handle_charref(self, ref):         
        # called for each character reference, e.g. for "&#160;", ref will be "160"
        # Reconstruct the original character reference.
        self.pieces.append("&#%(ref)s;" % locals())

    def handle_entityref(self, ref):       
        # called for each entity reference, e.g. for "&copy;", ref will be "copy"
        # Reconstruct the original entity reference.
        self.pieces.append("&%(ref)s" % locals())
        # standard HTML entities are closed with a semicolon; other entities are not
        if htmlentitydefs.entitydefs.has_key(ref):
            self.pieces.append(";")

    def handle_data(self, text):           
        # called for each block of plain text, i.e. outside of any tag and
        # not containing any character or entity references
        # Store the original text verbatim.
		output = text.replace("\xe2\x80\x99","'").split('\r\n')
		for count,item in enumerate(output):
			output[count] = item.strip()
		self.pieces.append(''.join(output))

    def handle_comment(self, text):        
        # called for each HTML comment, e.g. <!-- insert Javascript code here -->
        # Reconstruct the original comment.
        # It is especially important that the source document enclose client-side
        # code (like Javascript) within comments so it can pass through this
        # processor undisturbed; see comments in unknown_starttag for details.
        self.pieces.append("<!--%(text)s-->" % locals())

    def handle_pi(self, text):             
        # called for each processing instruction, e.g. <?instruction>
        # Reconstruct original processing instruction.
        self.pieces.append("<?%(text)s>" % locals())

    def handle_decl(self, text):
        # called for the DOCTYPE, if present, e.g.
        # <!DOCTYPE html PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN"
        #     "http://www.w3.org/TR/html4/loose.dtd">
        # Reconstruct original DOCTYPE
        self.pieces.append("<!%(text)s>" % locals())

    def output(self):              
        """Return processed HTML as a single string"""
        return "".join(self.pieces)

if __name__ == "__main__":
	
	parser = argparse.ArgumentParser(description='Convert html tags to markup used in chapman kelly psl project', version='%(prog)s 2.0')
	parser.add_argument('infile', type=str, help='input file to use')
	parser.add_argument('--out', type=str, default='-fixed.csv', help='what to append to the output file')
	args = parser.parse_args()

	myCSV = check_csv_file(args.infile)
	reader = csv.DictReader(open(myCSV, "rb"), dialect='excel')
	
	writer = csv.DictWriter(open(os.path.splitext(args.infile)[0]+args.out, "wb"), reader.fieldnames, dialect=d)
	writer.writeheader()
	parser = BaseHTMLProcessor()
	
	for row in reader:
		for k,v in row.iteritems():
			try:
				parser.feed(v)
				parser.close()
				row[k] = parser.output()
				parser.reset()
			except TypeError, e:
				row[k] = v
		
		writer.writerow(row)
