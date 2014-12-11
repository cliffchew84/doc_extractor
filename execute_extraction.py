"""
- This file extract all text from docx and text within a folder into a excel format
- There are 2 files, where data.csv provides the text, while title.csv provides the title of the docx file extracted.
- There were issues with unicode, so I had to include .encode('ascii', 'ignore')
- Had to use glob to identify all files that have docx and text.
- All docx and text files have to be in the same folder as this file.
- Python 2.7.6 was used, and no additional packages are required, except the script from the Internet was utilised.
"""

import csv, os, glob

# Change working directory
directory = os.path.dirname(os.path.realpath(__file__))
os.chdir(directory)
from extract import get_docx_text

# Locate all docx files in the folder
file_names_docx = glob.glob('*.docx')
files = []
for i in range(len(file_names_docx)):
    file_i = get_docx_text(file_names_docx[i])
    file_i = str((file_i.encode('ascii','ignore')).replace("\n\n", " "))
    files.append(file_i)

# Inclusion of extractor for text files
file_names_text = glob.glob("*.txt")

for i in file_names_text:
	text_i = open(i,"r")
	text = text_i.read()
	files.append(text)

""" Write the content into the csv file. """
class write_content:
	"""Write contents of docx into file"""
	f = open('data.csv', 'wb')
	w = csv.writer(f, delimiter = ',')
	for i in files:
	    w.writerows([(i,)])
	f.close()

class write_title:
	"""Write title of docx and text into file"""
	f2 = open('title.csv', 'wb')
	w2 = csv.writer(f2, delimiter = ',')

	# Write docx title into title.csv 
	for i in file_names_docx:
	    w2.writerows([(i,)])

	# Write text title into title.csv
	for i in file_names_text:
	    w2.writerows([(i,)])
	f2.close()

def main():
	write_title()
	write_content()
	print "All files copied into data.cav."

if __name__ == '__main__':
	main()