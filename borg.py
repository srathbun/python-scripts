#!/usr/bin/env python
# original: http://code.activestate.com/recipes/66531-singleton-we-dont-need-no-stinkin-singleton-the-bo/?in=lang-python
class Borg(object):
	_state = {}
	def __new__(cls, *p, **k):
		self = object.__new__(cls, *p, **k)
		self.__dict__ = cls._state
		return self

