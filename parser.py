#!/usr/bin/env python
'''
Author: Spencer Rathbun
Date: 1/24/2012

Summary: Parse a jetLetter file for all %m[\d+] items, and print out their corresponding name, collected from the fieldnames file.
'''
import argparse
from glob import glob
tokens = ('VAR', 'NUMBER', 'CLOSE', 'JUNK')

# Tokens

t_VAR     = r'%[mM]\['
t_CLOSE   = r'\]'
t_JUNK    = r'.'

# Ignored characters
t_ignore = " \t\r"

def t_NUMBER(t):
	r'\d+'
	try:
		t.value = int(t.value)
	except ValueError:
		print("Integer value too large %d", t.value)
		t.value = 0
	return t

def t_newline(t):
	r'\n+'
	t.lexer.lineno += t.value.count("\n")

def t_error(t):
	print("Illegal character '%s'" % t.value[0])
	t.lexer.skip(1)

# Build the lexer
import ply.lex as lex
lex.lex()

# Parsing rules

# dictionary of names
entries = []
fieldnames = []

def p_statements(p):
	'''statements : statement statements
                  | empty'''
	pass

def p_empty(p):
	'''empty :'''
	pass

def p_statement(p):
	'''statement : field'''
	try:
		entries.append(fieldnames[p[1]])
	except IndexError:
		pass

def p_trash(p):
	'''statement : JUNK
				 | NUMBER
				 | CLOSE'''
	pass

def p_field(p):
	'''field : VAR NUMBER CLOSE'''
	#print p[1], p[2], p[3]
	p[0] = p[2]

def p_error(p):
	print("Syntax error at '%s'" % repr(p)) #p.value)

import ply.yacc as yacc
yacc.yacc()

def main(**kwargs):
	with open(kwargs['f'], 'rb') as f:
		for line in f.readlines():
			fieldnames.append(line.strip('\r\n'))

	for currFile in kwargs['infile']:
		for globfile in glob(currFile):
			with open(globfile, 'rb') as f:
				for line in f.readlines():
					yacc.parse(line)
	from pprint import pprint

	pprint(entries)
	print ""
	print "Sorted and Duplicates removed:"
	pprint(sorted(list(set(entries))))

if __name__ == '__main__':
	parser = argparse.ArgumentParser(description='Read jetletter file and find all "M" variables', version='%(prog)s 1.0')
	parser.add_argument('infile', nargs='+', type=str, help='input files')
	parser.add_argument('-f', type=str, default='fieldnames.txt', help='File with fieldnames, one per line')
	parser.add_argument('--out', type=str, default='temp.txt', help='name of output file')
	args = parser.parse_args()
	main(**vars(args))
