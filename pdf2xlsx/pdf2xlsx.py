""" open a zip file containing pdf files, then parse the pdf file and
put them in a xlsx file"""
import os
import shutil
import zipfile
import re
from copy import deepcopy
from collections import namedtuple
from PyPDF2 import PdfFileReader
import xlsxwriter


TMP_DIR = 'tmp'
SRC_NAME = 'src.zip'
FILE_EXTENSION = '.pdf'

class Invoice:
    """
    """
    NO_PATTERN = '[ ]*Számla sorszáma:([0-9]+)'
    NO_CMP = re.compile(NO_PATTERN)
    ORIG_DATE_PATTERN = ('[ ]*Számla kelte:'
                         '([0-9]{4}\.[0-9]{2}\.[0-9]{2}|[0-9]{2}\.[0-9]{2}\.[0-9]{4})')
    ORIG_DATE_CMP = re.compile(ORIG_DATE_PATTERN)
    PAY_DUE_PATTERN = ('[ ]*FIZETÉSI HATÁRIDÕ:([0-9]{4}\.[0-9]{2}\.[0-9]{2}|'
                       '[0-9]{2}\.[0-9]{2}\.[0-9]{4})[ ]+([0-9]+\.?[0-9]*\.?[0-9]*)')
    PAY_DUE_CMP = re.compile(PAY_DUE_PATTERN)
    AWKWARD_DATE_PATTERN = '([0-9]{2})\.([0-9]{2})\.([0-9]{4})'
    AWKWARD_DATE_CMP = re.compile(AWKWARD_DATE_PATTERN)
    def __init__(self, no=0, orig_date="", pay_due="", total_sum=0, entries=None):
        self.no = no
        self.orig_date = orig_date
        self.pay_due = pay_due
        self.total_sum = total_sum
        self.entries = entries

        self.no_parsed = False
        self.orig_date_parsed = False
        self.pay_due_parsed = False

    def __str__(self):
        tmp_str = '\n    '.join([str(entr) for entr in self.entries])
        return ('Sorszam: {no}, Kelt: {orig_date}, Fizetesihatarido: {pay_due},'
                'Vegosszeg: {total_sum}\nBejegyzesek:\n    {entry_list}'
                '').format(**self.__dict__,entry_list=tmp_str)

    def __repr__(self):
        return ('{__class__.__name__}(no={no!r},orig_date={orig_date!r},'
                'pay_due={pay_due!r},total_sum={total_sum!r},'
                'entries={entries!r})').format(__class__=self.__class__, **self.__dict__)

    def _normalize_str_date(self,strdate):
        mo = self.AWKWARD_DATE_CMP.match(strdate)
        if mo:
            strdate = ''.join([mo.group(3),'.',mo.group(2),'.',mo.group(1)])
        return strdate

    def parse_line(self, line):
        if not self.no_parsed:
            mo = self.NO_CMP.match(line)
            if mo:
                self.no = int(mo.group(1))
                self.no_parsed = True
                #print("Sorszam found: {}".format(sorszam))
        elif not self.orig_date_parsed:
            mo = self.ORIG_DATE_CMP.match(line)
            if mo:
                self.orig_date = self._normalize_str_date(mo.group(1))
                self.orig_date_parsed = True
                #print("kelt found: {}".format(kelt))
        elif not self.pay_due_parsed:
            mo = self.PAY_DUE_CMP.match(line)
            if mo:
                self.pay_due = self._normalize_str_date(mo.group(1))
                self.total_sum = int(mo.group(2).replace('.',''))
                self.pay_due_parsed = True
                #print("fizetesi es vegosszeg found: {} {}".format(fizetesihatar, vegosszeg))

    def xlsx_write(self, worksheet, row, col):
        worksheet.write(row, col+1, self.no)
        worksheet.write(row, col+2, self.orig_date)
        worksheet.write(row, col+3, self.pay_due)
        worksheet.write(row, col+4, self.total_sum)
        return row+1, col



EntryTuple = namedtuple('EntryTuple', ['kod', 'nev', 'ME', 'mennyiseg', 'BEgysegar',
                             'Kedv', 'NEgysegar', 'osszesen', 'AFA'])

class Entry:
    """
    """
    CODE_PATTERN = '[ ]*([0-9]{6}-[0-9]{3})'
    CODE_CMP = re.compile(CODE_PATTERN)
    ENTRY_PATTERN = "".join([CODE_PATTERN,  #termek kod
                 ("(.*)" #termek megnevezes
                 "(Pár|Darab)" # ME
                 "[ ]+([0-9]+)" # mennyiseg
                 "[ ]+([0-9]+\.?[0-9]*)" # Brutto Egysegar
                 "[ ]+([0-9]+)%" # Kedvezmeny
                 "[ ]+([0-9]+\.?[0-9]*)" # Netto Egysegar
                 "[ ]+([0-9]+\.?[0-9]*\.?[0-9]*)" # Osszesen
                 "[ ]+([0-9]+)%") # Afa
    ])
    ENTRY_CMP = re.compile(ENTRY_PATTERN)

    def __init__(self, entry_tuple=None):
        self.entry_tuple = entry_tuple

        self.entry_found = False
        self.new_entry = False
        self.tmp_str = ""

    def __str__(self):
        return '{entry_tuple}'.format(**self.__dict__)

    def __repr__(self):
        return ('{__class__.__name__}(entry_tuple={entry_tuple!r}'
                ')').format(__class__=self.__class__, **self.__dict__)

    def line2entry(self,line):
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
            tg = self.ENTRY_CMP.match(line).groups()
            return EntryTuple(tg[0], tg[1], tg[2],
                         int(tg[3]), int(tg[4].replace('.','')),
                         int(tg[5]), int(tg[6].replace('.','')),
                         int(tg[7].replace('.','')), int(tg[8]))
        except AttributeError as e:
            print("Entry pattern regex didn't match for line: {}".format(pdfline))
            raise e

    def parse_line(self,line):
        self.new_entry = False
        if self.entry_found:
            self.entry_tuple = self.line2entry(" ".join([self.tmp_str, line]))
            self.entry_found = False
            self.new_entry = True
        elif self.CODE_CMP.match(line):
            self.tmp_str = line
            self.entry_found = True

def pdf2rawtxt(pdfile):
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
        invo = Invoice(entries=[])
        entry = Entry()
        for i in range(tmp_input.getNumPages()):
            for line in tmp_input.getPage(i).extractText().split('\n'):
                invo.parse_line(line)
                entry.parse_line(line)
                if(entry.new_entry):
                    invo.entries.append(deepcopy(entry))
        return invo

def _init_clean_up(tmp_dir='tmp'):
    """
    Create tmp directory, clean it up first if it already exists, if possible
    :param str tmp_dir: this is the directory to work during the execution
    """
    try:
        shutil.rmtree(tmp_dir)
    except FileNotFoundError as e:
        #Do not print any silly message as this is the expected behaviour
        pass
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

def extract_invoces(pdf_list):
    """
    Get the invoice entries from the pdf files in th pdf_list
    Calls pdf2rawtxt.
    TODO: Refactor this to a generator to decouple it from pdf2rawtct
    :param list pdf_list: List of pdf files path to process.
    """
    invoice_list = []
    for pdfile in pdf_list:
        invoice_list.append(pdf2rawtxt(pdfile))
    return invoice_list

def _post_clean_up(tmp_dir='tmp'):
    """
    Cleanup after execution, remove the extracted zip file and tmp directory
    :param str tmp_dir: Temporary directory to clean_up
    """
    shutil.rmtree(tmp_dir)

def _gen_header(worksheet, row, col):
    worksheet.write(row, col+1, "Invoice Number")
    worksheet.write(row, col+2, "Date of Invoice")
    worksheet.write(row, col+3, "Payment Date")
    worksheet.write(row, col+4, "Amount")
    return row+1, col

def invoices2xlsx(invoices):
    workbook = xlsxwriter.Workbook('Invoices01.xlsx')
    worksheet = workbook.add_worksheet()
    row = 0
    col = 0
    row, col = _gen_header(worksheet, row, col)
    for invo in invoices:
        row, col = invo.xlsx_write(worksheet, row, col)

    workbook.close()

def main():
    _init_clean_up(TMP_DIR)

    extract_zip(SRC_NAME, TMP_DIR)

    pdf_list = get_pdf_files(os.path.join(os.getcwd(),TMP_DIR), FILE_EXTENSION)

    invoice_list = extract_invoces(pdf_list)

    _post_clean_up(TMP_DIR)

    invoices2xlsx(invoice_list)

    print("script has been finished")


if __name__ == '__main__': main()

