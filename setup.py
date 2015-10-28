import glob
from distutils.core import setup

# get all of the scripts
scripts = glob.glob("bin/*")

setup(
  name = 'docx2txt',
  packages = ['docx2txt'], # this must be the same as the name above
  version = '0.2',
  description = 'A pure python-based utility to extract text from docx files.',
  author = 'Ankush Shah',
  author_email = 'ankush.shah.nitk@gmail.com',
  url = 'https://github.com/ankushshah89/python-docx2txt', # use the URL to the github repo
  download_url = 'https://github.com/ankushshah89/python-docx2txt/tarball/0.2', # I'll explain this in a second
  keywords = ['python', 'docx', 'text'], # arbitrary keywords
  scripts = scripts,
  classifiers = [],
)
