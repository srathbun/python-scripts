from distutils.core import setup
import py2exe, sys, os

#setup(
#	options={'py2exe': {'bundle_files': 1}},
#	console=['excelConverter.py'],
#	zipfile=None,
#)
# NOTE: on w7  'dll_excludes': [ "mswsock.dll", "powrprof.dll" ] is needed for py2exe, or else you get MemoryLoadFailure exceptions
setup(console = [{"script":"AffidavitReader.py", "icon_resources":[(1, "c:\python27\DLLs\py.ico")]}],
	  zipfile = None,
	  options = {"py2exe":{"compressed": 1, "optimize": 1, "ascii": 1, "bundle_files": 1,  'dll_excludes': [ "mswsock.dll", "powrprof.dll" ]}})
