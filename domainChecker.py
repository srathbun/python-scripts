'''
Project: domain validity checker
Author: Spencer Rathbun
Date: 7/6/2011

Summary: Skeleton file for Rebecca, demonstrating argparse and csv module

'''
import argparse
from socket import getaddrinfo

# a python script does not need this line defining a main function
# but we need it for creating a program with command line arguments, rather than a script
if __name__ == "__main__":
	
	# sets up command line arguments, this allows us to input our files and switches
	parser = argparse.ArgumentParser(description='Check for validity of domains in list exported from exchange', version='%(prog)s 1.0')
	parser.add_argument('infile', nargs='+', type=str, help='list of input files')
	# get the current arguments and put them into a variable
	args = parser.parse_args()

	domains = []
	for f in args.infile:
		with open(f, 'rt') as data:
			for line in data.readlines():
				split = line.replace('\x00',"").split(':')
				if split[0].strip() == 'Domain':
					domains.append(split[1].strip())

	for domain in domains:
		try:
			getaddrinfo(domain, None)
		except Exception, e:
			print "Unable to resolve:", domain
