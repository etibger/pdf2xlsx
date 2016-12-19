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
                 "[ ]+([0-9]+)%" # Kedvezmeny
                 "[ ]+([0-9]+\.?[0-9]*)" # Netto Egysegar
                 "[ ]+([0-9]+\.?[0-9]*\.?[0-9]*)" # Osszesen
                 "[ ]+([0-9]+)%") # Afa
    ])

EPROG = re.compile(ENTRY_PATTERN)

def line2entry(pdfline):
    """
    Extracts entry information from the given line. Tries to search for nine different
    group in the line. See implementation of ENTRY_PATTERN. This should match the
    following pattern:
    DDDDDD-DDD STR+WSPACE PREDEFSTR INTEGER INTEGER-. INTEGER% INTEGER-. INTEGER-. INTEGER%
    Where:
    D: a number 0-9
    STR+WSPACE: string containing white spaces, possibly numbers and special characters
    PREDEFSTR: string without white space ( predefined )
    INTEGER: decimal number, unknown length
    INTEGER-.: a decimal number, grouped with . by thousends e.g 1.589.674
    INTEGER%: an integer with percentage at the end

    :param str pdfline: Line to parse, this line should be begin with NNNNNNN-NNN
    """
    try:
        tg = EPROG.match(pdfline).groups()
        return Entry(tg[0], tg[1], tg[2], tg[3], tg[4], tg[5], tg[6], tg[7], tg[8])
    except AttributeError as e:
        print("Entry pattern regex didn't match for line: {}".format(pdfline))
        raise e

def pdf2rawtxt(pdfile, entries):
    """
    Extracts text from the given pdf file, searches the invoice entries (the lines
    starting with NNNNNN-NNN pattern. This line and the next represents the full
    invoce entry line. This line is processed with line2entry.
    TODO: possibly it would be nice to refactor this as a generator to decouple it
    from line2entry function
    
    :param str pdfile: file path of the pdf to process
    :param list entries: The found invoice entries will be appended to this list
    """
    with open(pdfile, 'rb') as fd:
        tmp_input = PdfFileReader(fd)
        for i in range(tmp_input.getNumPages()):
            entry_found = False
            tmp_str = ""
            for line in tmp_input.getPage(i).extractText().split('\n'):
                if entry_found:
                    entries.append(line2entry(" ".join([tmp_str, line])))
                    entry_found = False
                elif KODPROG.match(line):
                    tmp_str = line
                    entry_found = True

def _init_clean_up(tmp_dir='tmp'):
    """
    Create tmp directory, clean it up first if it already exists, if possible
    :param str tmp_dir: this is the directory to work during the execution
    """
    try:
        shutil.rmtree(tmp_dir)
    except FileNotFoundError as e:
        print("The directory is not there, this was the exception\n {}".format(e))
    finally:
        os.mkdir(tmp_dir)

def extract_zip(src_name, dir):
    """
    Extract the zip file files from the zip, there is no sanity check for the arguments
    :param str src_name: Path to a zip file to extract
    :param str dir: Path to the target directory to extract the zip file
    """
    with zipfile.ZipFile(src_name) as myzip:
        myzip.extractall(dir)

def get_pdf_files(dir, extension='.pdf'):
    """
    Walks through the given **dir** and collects every files with **extension**
    :param str dir: the root directory to start the walk
    :param str extension: '.pdf' by default, if the file has this extension it is selected
    """
    pdf_list = []
    for dir_path, _dummy, file_list in os.walk(dir):
        for filename in file_list:
            if filename.endswith(extension):
                pdf_list.append(os.path.join(dir_path, filename))
    return pdf_list

def extract_invoce_entries(pdf_list):
    """
    Get the invoice entries from the pdf files in th pdf_list
    Calls pdf2rawtxt.
    TODO: Refactor this to a generator to decouple it from pdf2rawtct
    :param list pdf_list: List of pdf files path to process.
    """
    invoice_entries = []
    for pdfile in pdf_list:
        pdf2rawtxt(pdfile, invoice_entries)
    return invoice_entries

def _post_clean_up(tmp_dir='tmp'):   
    """
    Cleanup after execution, remove the extracted zip file and tmp directory
    :param str tmp_dir: Temporary directory to clean_up
    """
    shutil.rmtree(tmp_dir)

def main():
    _init_clean_up(TMP_DIR)
    
    extract_zip(SRC_NAME, TMP_DIR)

    pdf_list = get_pdf_files(os.path.join(os.getcwd(),TMP_DIR), FILE_EXTENSION)

    invoice_entries = extract_invoce_entries(pdf_list)

    for ie in invoice_entries:
        print(ie)

    _post_clean_up(TMP_DIR)
    
    print("script has been finished")

if __name__ == '__main__': main()
