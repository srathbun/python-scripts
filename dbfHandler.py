#!/usr/bin/env python
# -*- coding: ascii -*-
'''This class reads dbf files, and contains methods for returning them as useful data structures.'''
import logging, struct
from itertools import izip
from decimal import *
class dbfHandler:
	def __init__(self, dbf):
		self.logger = logging.getLogger('dirMon.dbfHandler')
		self.dbfFilename = dbf
		self.logger.debug('filename: {0}'.format(dbf))

	def __del__(self):
		'''Closes file handle during garbage collection. No guarantee this will run in a timely manner!'''
		self.dbfFile.close()

	def __enter__(self):
		'''Open the file when entering a with block.'''
		self.dbfFile = open(self.dbfFilename, 'rb')
		return self

	def __exit__(self, *args):
		'''Close the file when leaving a with block.'''
		self.dbfFile.close()

	def reset(self):
		'''Reset file handle, so that we can rerun the data file.'''
		self.dbfFile.seek(0)
	
	def close(self):
		'''Close the file handle manually.'''
		self.dbfFile.close()
	
	def open(self):
		'''Open the file handle manually.'''
		self.dbfFile = open(self.dbfFilename, 'rb')

	def convertRows(self):
		'''Convert the records returned from a dbf iterator into a list of dictionaries.'''
		dbfiter = self.dbfreader(self.dbfFile)
		fieldnames = dbfiter.next()
		fieldspecs = dbfiter.next()
		rows = []
		for record in dbfiter:
			zipped = zip(fieldnames, record)
			dictionary = dict(zipped)
			rows.append(dictionary)
		return rows

	def dbfreader(self, f):
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
		for fieldno in xrange(numfields):
			name, typ, size, deci = struct.unpack('<11sc4xBB14x', f.read(32))
			name = name.replace('\0', '')       # eliminate NULs from string   
			fields.append((name, typ, size, deci))

		terminator = f.read(1)
		assert terminator == '\r'

		fields.insert(0, ('DeletionFlag', 'C', 1, 0))
		fmt = ''.join(['%ds' % fieldinfo[2] for fieldinfo in fields])
		fmtsiz = struct.calcsize(fmt)

		yield [field[0] for field in fields]
		yield [tuple(field[1:]) for field in fields]
		for i in xrange(numrec):
			record = struct.unpack(fmt, f.read(fmtsiz))
			# bcc uses dbf deletion for "hiding" records
			# therefore, we have to read ALL the recs in and deal with the deleted accordingly
			#if record[0] != ' ':
			#	continue                        # deleted record
			result = []
			for (name, typ, size, deci), value in izip(fields, record):
				try:
					if name == 'DeletionFlag':
						if value == ' ':
							value = False
						else:
							value = True
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
				except Exception, e:
					self.logger.debug(str(e))
					self.logger.debug(' '.join((name, value)))
					result.append(None)
			yield result

	def dbfwriter(f, fieldnames, fieldspecs, records):
		""" Return a string suitable for writing directly to a binary dbf file.

		File f should be open for writing in a binary mode.

		Fieldnames should be no longer than ten characters and not include \x00.
		Fieldspecs are in the form (type, size, deci) where
			type is one of:
				C for ascii character data
				M for ascii character memo data (real memo fields not supported)
				D for datetime objects
				N for ints or decimal objects
				L for logical values 'T', 'F', or '?'
			size is the field width
			deci is the number of decimal places in the provided decimal object
		Records can be an iterable over the records (sequences of field values).
		
		"""
		# header info
		ver = 3
		now = datetime.datetime.now()
		yr, mon, day = now.year-1900, now.month, now.day
		numrec = len(records)
		numfields = len(fieldspecs)
		lenheader = numfields * 32 + 33
		lenrecord = sum(field[1] for field in fieldspecs) + 1
		hdr = struct.pack('<BBBBLHH20x', ver, yr, mon, day, numrec, lenheader, lenrecord)
		f.write(hdr)
						  
		# field specs
		for name, (typ, size, deci) in itertools.izip(fieldnames, fieldspecs):
			name = name.ljust(11, '\x00')
			fld = struct.pack('<11sc4xBB14x', name, typ, size, deci)
			f.write(fld)

		# terminator
		f.write('\r')

		# records
		for record in records:
			f.write(' ')                        # deletion flag
			for (typ, size, deci), value in itertools.izip(fieldspecs, record):
				if typ == "N":
					value = str(value).rjust(size, ' ')
				elif typ == 'D':
					value = value.strftime('%Y%m%d')
				elif typ == 'L':
					value = str(value)[0].upper()
				else:
					value = str(value)[:size].ljust(size, ' ')
				assert len(value) == size
				f.write(value)

		# End of file
		f.write('\x1A')
