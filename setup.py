import os
import glob
from distutils.core import setup
from docxpy import VERSION

# get all of the scripts
scripts = glob.glob('bin/*')


def read(fname):
    return open(os.path.join(os.path.dirname(__file__), fname)).read()


setup(
    name='docxpy',
    packages=['docxpy'],
    version=VERSION,
    description='A pure python-based utility to extract text, hyperlinks and images'
              'from docx files.',
    long_description=open("README.rst").read(),
    author='Ankush Shah, Yalei Du',
    author_email='yaleidu@163.com',
    url='https://github.com/badbye/docxpy',
    keywords=['python', 'docx', 'text', 'links', 'images', 'extract'],
    scripts=scripts,
    test_suite='nose.collector',
    tests_require=['nose'],
    classifiers=[
        "Programming Language :: Python :: 2.7",
        "Programming Language :: Python :: 3.3",
        "Programming Language :: Python :: 3.4",
        "Programming Language :: Python :: 3.5"
  ]
)
