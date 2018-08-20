import sys
import os

#Uses pip to install packages
def install(package):
	pip.main(['install', package])
#end of function

#tries to import pip; if pip is not installed, installs pip using get-pip.py
try:
	import pip
	print "Pip is already installed."
except ImportError:
	print "Pip is not installed, installing it now!"
	os.system("python get-pip.py")
	print "Just installed pip."
#end of try-except

#tries to import docx; if docx is not installed, installs docx by downloading package
try:
	import docx
	print "Python-docx is already installed."
except ImportError:
	print "Python-docx is not installed, installing it now!"
	install('python-docx')
	print "Just installed python-docx. Please re-run this script at your convenience."
#end of try-except

#exit program
sys.exit(1)