from distutils.core import setup
import py2exe, sys, os

#setup(
#	options={'py2exe': {'bundle_files': 1}},
#	console=['excelConverter.py'],
#	zipfile=None,
#)
setup(console = [{"script":"uniquify.py", "icon_resources":[(1, "c:\python27\DLLs\py.ico")]}],
	  zipfile = None,
	  options = {"py2exe":{"compressed": 1, "optimize": 1, "ascii": 1, "bundle_files": 1}})
