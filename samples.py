# python design patterns book
# http://dpip.testingperspective.com/

# Python Lex-Yacc
# http://www.dabeaz.com/ply/

# python 25 scripts: http://pythoncodingforum.com/showthread.php?tid=3&pid=3#pid3

# console hacks: http://matt.might.net/articles/console-hacks-exploiting-frequency/
# roll your own unix: http://www.jamesmolloy.co.uk/tutorial_html/
def combineDicts(dictionary1, dictionary2):
	output = {}
	for item, value in dictionary1.iteritems():
		if dictionary2.has_key(item):
			if isinstance(dictionary2[item], dict):
				output[item] = combineDicts(value, dictionary2.pop(item))
		else:
			output[item] = value
	for item, value in dictionary2.iteritems():
		output[item] = value
	return output
