""" open a zip file containing pdf files, then parse the pdf file and
put them in a xlsx file"""
import os
import shutil
import zipfile
import re
from collections import namedtuple
from PyPDF2 import PdfFileReader

TMP_DIR = 'tmp'
SRC_NAME = 'src.zip'
FILE_EXTENSION = '.pdf'

Entry = namedtuple('Entry', ['kod', 'nev', 'ME', 'mennyiseg', 'BEgysegar',
                             'Kedv', 'NEgysegar', 'osszesen', 'AFA'])

KOD_PATTERN = '[ ]*([0-9]{6}-[0-9]{3})'
KODPROG = re.compile(KOD_PATTERN)

ENTRY_PATTERN = "".join([KOD_PATTERN,  #termek kod
                 ("(.*)" #termek megnevezes
                 "(Pár|Darab)" # ME
                 "[ ]+([0-9]+)" # mennyiseg
                 "[ ]+([0-9]+\.?[0-9]*)" # Brutto Egysegar
                 "[ ]+([0-9]+%)" # Kedvezmeny
                 "[ ]+([0-9]+\.?[0-9]*)" # Netto Egysegar
                 "[ ]+([0-9]+\.?[0-9]*\.?[0-9]*)" # Osszesen
                 "[ ]+([0-9]+%)") # Afa
    ])

EPROG = re.compile(ENTRY_PATTERN)

def line2entry(pdfline):
    try:
        tg = EPROG.match(pdfline).groups()
        return Entry(tg[0], tg[1], tg[2], tg[3], tg[4], tg[5], tg[6], tg[7], tg[8])
    except AttributeError as e:
        print("Entry pattern regex didn't match for line: {}".format(pdfline))
        raise e

def pdf2rawtxt(pdfile, entries):
    with open(pdfile, 'rb') as fd:
        tmp_input = PdfFileReader(fd)
        for i in range(tmp_input.getNumPages()):
            tmp_page = tmp_input.getPage(i)
            txt = tmp_page.extractText()
            txt = txt.split('\n')
            txt2 = []
            entry_found = False
            tmp_str = ""
            for line in txt:
                if entry_found:
                    txt2.append(" ".join([tmp_str, line]))
                    entry_found = False
                elif KODPROG.match(line):
                    tmp_str = line
                    entry_found = True
            #print(len(txt2))
            j = 0
            for line in txt2:
                print("{0}: {1}".format(j, line))
                #tmp_entry = line2entry(line)
                #print(tmp_entry)
                entries.append(line2entry(line))
                j += 1
   
            
    

# create tmp directory, clean it up first if it already exists, if possible
try:
    shutil.rmtree(TMP_DIR)
except FileNotFoundError as e:
    print("The directory is not there, this was the exception\n {}".format(e))
finally:
    os.mkdir(TMP_DIR)

# get the pdf files from the zip
with zipfile.ZipFile(SRC_NAME) as myzip:
    myzip.extractall(TMP_DIR)

# collect the every pdf file to process
pdf_list = []
for dir_path, _dummy, file_list in os.walk(os.path.join(os.getcwd(),TMP_DIR)):
    for filename in file_list:
        if filename.endswith(FILE_EXTENSION):
            pdf_list.append(os.path.join(dir_path, filename))

#get the raw text from pdf files
invoice_entries = []
for pdfile in pdf_list:
    pdf2rawtxt(pdfile, invoice_entries)


#clean_up, tmp should exist
shutil.rmtree(TMP_DIR)
    
print("script has been finished")
