import glob
from distutils.core import setup
from docx2txt import VERSION

# get all of the scripts
scripts = glob.glob('bin/*')

setup(
  name='docx2txt',
  packages=['docx2txt'],
  version=VERSION,
  description='A pure python-based utility to extract text, links and images'
              'from docx files.',
  author='Ankush Shah',
  author_email='ankush.shah.nitk@gmail.com',
  url='https://github.com/ankushshah89/python-docx2txt',
  download_url='https://github.com/ankushshah89/python-docx2txt/tarball/0.6',
  keywords=['python', 'docx', 'text', 'links', 'images', 'extract'],
  scripts=scripts,
  classifiers=[],
)
