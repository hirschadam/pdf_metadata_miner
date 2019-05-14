# Author: Adam Hirsch
# Platform compatability: Windows
# Program: This script iterates through a directory of PDF & TXT files, extracts metadata from each PDF, 
#		   and the URL where the PDF was download from a corresponding TXT file.
#          then outputs the metadata to a CSV file.
#          e.g. File input: file.pdf 
#               File output: file_pdf.csv
#                            file_pdfmeta.csv
import os
import csv
from pdfminer.pdfinterp import PDFResourceManager, PDFPageInterpreter
from pdfminer.converter import TextConverter
from pdfminer.layout import LAParams
from pdfminer.pdfpage import PDFPage
from pdfminer.pdfparser import PDFParser
from pdfminer.pdfdocument import PDFDocument
from io import StringIO

# Modify this path to where your PDF/TXT files are
ROOTDIR = ''

def convert(data):
	"""Convert a dictionary keys & values to strings
	"""
	if isinstance(data, bytes):      return data.decode()
	if isinstance(data, (str, int)): return str(data)
	if isinstance(data, dict):       return dict(map(convert, data.items()))
	if isinstance(data, tuple):      return tuple(map(convert, data))
	if isinstance(data, list):       return list(map(convert, data))
	if isinstance(data, set):        return set(map(convert, data))
  
def convert_pdf_to_txt(path):
	"""Parse PDF document data into text
	return text: PDF text
	"""
	rsrcmgr = PDFResourceManager()
	retstr = StringIO()
	codec = 'utf-8'
	laparams = LAParams()
	device = TextConverter(rsrcmgr, retstr, codec=codec, laparams=laparams)
	fp = open(ROOTDIR + '\\' + path, 'rb') # Open PDF 
	interpreter = PDFPageInterpreter(rsrcmgr, device)
	password = ""
	maxpages = 0
	caching = True
	pagenos=set()

	# Iterate through and process each page of the PDF
	for page in PDFPage.get_pages(fp, pagenos, maxpages=maxpages, password=password,caching=caching, check_extractable=True):
		interpreter.process_page(page)
	# Convert PDF data to text
	text = retstr.getvalue() 
	
	fp.close()
	device.close()
	retstr.close()

	return text
		
def convert_pdf_to_metadata(file):
	"""Parse PDF document metadata into text and returns the 
	metadata key-value pairs as a dictionary
	"""
	fp = open(ROOTDIR + '\\' + file, 'rb')
	parser = PDFParser(fp) 
	doc = PDFDocument(parser)
	# List containing a dictionary of metadata key-value pairs
	text = doc.info 
	text_dict = text[0]
	text_dict['PDF File'] = file
	#fp.close()
	return text_dict 

def process_pdf_files():
	"""Extracts metadata from all PDFS in ROOTDIR then builds a list of all possible
	metadata keys
	
	return  header_keys: list containing keys for the CSV header
			metadata_list: list containing PDF metadata dictionaries
	"""
	# Initialize the list to hold each PDF's metadata dictionary
	metadata_list = list() 
	# Initialize dictionary to hold a PDF key-value pairs
	header_dict = dict() 
	# Build the metadata list of dictionaries
	for filename in os.listdir(ROOTDIR):
		if filename.endswith('.pdf'):
		
			try:	
				keyval = convert_pdf_to_metadata(filename)
			except:
				continue
			#metadata_list.append(keyval) # Add PDF metadata dictionary to the list
			for key, value in keyval.items(): 
				if(key in header_dict):
					# Append new value to the existing array at this slot
					header_dict[key].append(value)
				else:
					# Create a new array in this slot
					header_dict[key] = [value]
			# Extract URL from corresponding TXT file
			try:
				file_name = ROOTDIR + "\\" + filename.replace(".pdf", ".txt")
				#print(file_name)
				with open(file_name, encoding="UTF-16") as f:
					URL = f.readline().strip("\n")
					keyval['URL'] = URL
			except:
				pass
			metadata_list.append(keyval) # Add PDF metadata dictionary to the list
			
	# Build the CSV header list				
	header_keys = list()
	header_keys.append('URL')
	header_keys.append('PDF File')
	for key, value in header_dict.items():
		header_keys.append(key)
	return header_keys, metadata_list
	
def build_pdf_report(fieldnames, dict_list):
	""" Creates a CSV file report containing metadata findings
	"""
	with open('pdf_report.csv', mode='w') as report_file:
		report_writer = csv.DictWriter(report_file, fieldnames=fieldnames)
		report_writer.writeheader()		
		# Iterate through list of dictionary objects and convert any byte key/values to strings
		for item in dict_list:
			try:
				new_item = convert(item)
				report_writer.writerow(new_item)
			except:
				continue
				
		
		report_file.close()
			
if __name__ == "__main__":
	if not ROOTDIR:
		print("You must specify a directory to scan the PDF files.")
	else:
		print("...Building report...")
		metadata_key_list, metadata_dict_list = process_pdf_files()
		build_pdf_report(metadata_key_list, metadata_dict_list)
	
