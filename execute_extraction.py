"""
- This file extracst all text from docx within a folder into a excel format
- There are 2 files, where data.csv provides the text, while title.csv provides
the title of the docx file extracted.
- An excellent code from the Internet was extracted to use. The weblink is
attached below.
- In total, I think I took 2-3 hours to write this.
- There were issues with unicode, so I had to include .encode('ascii', 'ignore')
- Had to use glob to identify all files that have docx.
- All docx files have to be in the same folder as this file, although this can
be modified.
- All text was stripped of the \n, which were something like the "Enter Key". 
- Python 2.7.6 was used, and no additional packages are required, except the 
script from the Internet was utilised.
"""

import csv, os, glob

# Change working directory
directory = os.path.dirname(os.path.realpath(__file__))
os.chdir(directory)
from extract import get_docx_text

file_names = glob.glob('*.docx')
files = []
for i in range(len(file_names)):
    file_i = get_docx_text(file_names[i])
    file_i = str((file_i.encode('ascii','ignore')).replace("\n\n", " "))
    files.append(file_i)

f = open('data.csv', 'wb')
w = csv.writer(f, delimiter = ',')
for flies in files:
    w.writerows([(flies,)])
f.close()

f2 = open('title.csv', 'wb')
w2 = csv.writer(f2, delimiter = ',')
for flies in file_names:
    w2.writerows([(flies,)])
f2.close()

print len(files)
print "All the files have been copied into data.cav."
