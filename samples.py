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
