import glob
from distutils.core import setup
from docx2txt import VERSION

# get all of the scripts
scripts = glob.glob('bin/*')

setup(
  name='pydocx',
  packages=['pydocx'],
  version=VERSION,
  description='A pure python-based utility to extract text, hyperlinks and images'
              'from docx files.',
  author='Ankush Shah, Yalei Du',
  author_email='yaleidu@163.com',
  url='https://github.com/badbye/python-docx2txt',
  keywords=['python', 'docx', 'text', 'links', 'images', 'extract'],
  scripts=scripts,
  classifiers=[],
)
