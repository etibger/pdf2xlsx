"""
Contains framework to upen Zip a zip file which contains multiple pdf files
representing invoices. The invoices are parsed into Invoice and Entry (invoice
entries) classes. These are converted to XLSX format.
"""

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
DST_DIR = ''
FILE_EXTENSION = '.pdf'
XLSX_NAME= 'Invoices01.xlsx'

EntryTuple = namedtuple('EntryTuple', ['kod', 'nev', 'ME', 'mennyiseg', 'BEgysegar',
                             'Kedv', 'NEgysegar', 'osszesen', 'AFA'])
NO_INDENT = 1
ORIG_DATE_INDENT = 2
PAY_DUE_INDENT = 3
TOTAL_SUM_INDENT = 4

class StatLogger():
    """
    Collect statistic about the zip to xlsx process. Assembles a list containin invoice
    number of items. Every item is the number of entries found during the invoice parsing.
    """
    def __init__(self):
        self.invo_list = []

    def __str__(self):
        return '{invo_list}'.format(**self.__dict__)

    def new_invo(self):
        """
        When a new invoice was found create a new invoice log instance
        The current implementation is a simple list of numbers
        """
        self.invo_list.append(0)

    def new_entr(self):
        """
        When a new entry was found increase the entry counter for the current
        invoice.
        """
        self.invo_list[-1] += 1


def list2row(worksheet, row, col, values=[], positions=[]):
    """
    Create header of the template xlsx file

    :param Worksheet worksheet: Worksheet class to write info
    :param int row: Row number to start writing
    :param int col: Column number to start writing

    :return: the next position of cursor row,col
    :rtype: tuple of (int,int)
    """
    if not positions or len(positions) != len(values):
        positions = range(len(values))
    for v,p in zip(values,positions):
        worksheet.write(row, col+p, v)
    return row+1, col

class Invoice():
    """
    Parse, store and write to xlsx invoce informations. Such as Invoice Number,
    Invoice Date, Payment Date, Total Sum Price. It also contains a list of Entry,
    which is also extracted form raw string.
    The parsing of the raw string is controlled by three state variables: no_parsed,
    orig_date_parsed and pay_due_parsed. These represent the structure of the pdf.

    :param int no: Invoice number, default:0
    :param str orig_date: Invoice date stored as a string YYYY.MM.DD
    :param str pay_due: Payment Date stored as string YYYY.MM.DD
    :param int total_sum: Total price of invoice
    :param list entries: List of :class:`Entry` containing each entries in invoice

    
    [TODO] use proper datetime instead of string representation
    [TODO] implement state pattern for parsing ???
    """
    
    NO_PATTERN = '[ ]*Számla sorszáma:([0-9]{10})'
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
        """
        The date is represented in two different format in the pdf: YYYY.MM.DD and
        DD.MM.YYYY. The second needs to be converted to the first one.

        :param str strdate: string representation of the parsed date to normalize

        :return: a string representing a date with the format YYYY.MM.DD
        
        [TODO] remove this function when proper datetime is uzed to handle datest.
        """
        mo = self.AWKWARD_DATE_CMP.match(strdate)
        if mo:
            strdate = ''.join([mo.group(3),'.',mo.group(2),'.',mo.group(1)])
        return strdate

    def parse_line(self, line):
        """
        Parse through a raw text which is supplied line-by-line. This is the structure
        of the pdf (the brackets() indicate what should be collected):
        <disinterested rubish>
        Számla sorszáma: (NNNNNNNN) ...
        <disinterested rubish>
        Számla kelte: (YYYY.MM.DD|DD.MM.YYYY) ...
        <disinterested rubish>
        FIZETÉSI HATÁRIDŐ:(YYYY.MM.DD|DD.MM.YYYY) (NNN[.NNN.NNN])
        <disinterested rubish>
        This is structure is paresed using the three state variable, and stored inside
        the class attributes
        
        :param str line: The actual line to parse

        :return: True when the parsing of the Invoice was started
        :rtype: bool
        """
        if not self.no_parsed:
            mo = self.NO_CMP.match(line)
            if mo:
                self.no = int(mo.group(1))
                self.no_parsed = True
                return True
        elif not self.orig_date_parsed:
            mo = self.ORIG_DATE_CMP.match(line)
            if mo:
                self.orig_date = self._normalize_str_date(mo.group(1))
                self.orig_date_parsed = True
                return False
        elif not self.pay_due_parsed:
            mo = self.PAY_DUE_CMP.match(line)
            if mo:
                self.pay_due = self._normalize_str_date(mo.group(1))
                self.total_sum = int(mo.group(2).replace('.',''))
                self.pay_due_parsed = True
                return False
        return False

    def xlsx_write(self, worksheet, row, col):
        """
        Write the invoice information to a template xlsx file.

        :param Worksheet worksheet: Worksheet class to write info
        :param int row: Row number to start writing
        :param int col: Column number to start writing

        :return: the next position of cursor row,col
        :rtype: tuple of (int,int)
        """    
        values = [self.no, self.orig_date, self.pay_due, self.total_sum]
        positions = [NO_INDENT, ORIG_DATE_INDENT, PAY_DUE_INDENT,
                 TOTAL_SUM_INDENT]
        row, col = list2row(worksheet, row, col, values, positions)
        return row, col


class Entry():
    """
    Parse, store and write to xlsx invoice entries. The invoice informations are
    stored in the EntryTuple namedtuple. The parsing is contolled by a state
    variable (:entry_found:) Because the invoice entries are split into two line,
    the tmp_str attribute is used to store the first part of the entire

    :param EntryTuple entry_tuple: The invoice entry
    :param Invoice invo: The parent invoice containing this entry
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

    def __init__(self, entry_tuple=None, invo=None):
        self.entry_tuple = entry_tuple
        self.invo = invo

        self.entry_found = False
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
        NNNNNN-NNN STR+WSPACE PREDEFSTR INTEGER INTEGER-. INTEGER% INTEGER-. INTEGER-.
        INTEGER%
        Where:
        N: a single digit: 0-9
        STR+WSPACE: string containing white spaces, numbers and special characters
        PREDEFSTR: string without white space ( predefined )
        INTEGER: decimal number, unknown length
        INTEGER-.: a decimal number, grouped with . by thousends e.g 1.589.674
        INTEGER%: an integer with percentage at the end

        :param str pdfline: Line to parse, this line should be begin with NNNNNNN-NNN

        :return: The actual invoice entry
        :rtype: EntryTuple
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
        """
        Parse through raw text which is supplied line-by-line. This is the structure
        of the pdf (the brackets() indicate what should be collected):
        n times:
        <disinterested rubish>
        (NNNNNN-NNN ... \n
        ...)
        <disinterested rubish>
        When the Invoice code is found, an additional line is waited, and then it is
        sent to the line2entry converter.
        
        :param str line: The actual line to parse

        :return: True when an entry was found
        :rtype: bool
        """
        if self.entry_found:
            self.entry_tuple = self.line2entry(" ".join([self.tmp_str, line]))
            self.entry_found = False
            return True
        elif self.CODE_CMP.match(line):
            self.tmp_str = line
            self.entry_found = True
        return False

    def xlsx_write(self, worksheet, row, col):
        """
        Write the entry information to a template xlsx file.

        :param Worksheet worksheet: Worksheet class to write info
        :param int row: Row number to start writing
        :param int col: Column number to start writing

        :return: the next position of cursor row,col
        :rtype: tuple of (int,int)
        """
        values = [self.invo.no] + list(self.entry_tuple._asdict().values())
        row, col = list2row(worksheet, row, col, values)
        return row, col


#[TODO] Put this to a manager class???
def pdf2rawtxt(pdfile, logger):
    """
    Read out the given pdf file to Invoice and Entry classes to parse it. Utilize
    PyPFD2 PdfFileReader. Go through every page of the pdf. When a new invoice
    entry was found by the Entry.parse_line it is appended to the Invoice.entries

    :param str pdfile: file path of the pdf to process
    :param logger: :class:`StatLogger`, collect statistical data about parsing

    :return: The invoice entry filled up with the information from pdf file
    :rtype: :class:`Invoice`
    """
    with open(pdfile, 'rb') as fd:
        tmp_input = PdfFileReader(fd)
        invo = Invoice(entries=[])
        entry = Entry(invo=invo)
        for i in range(tmp_input.getNumPages()):
            for line in tmp_input.getPage(i).extractText().split('\n'):
                if invo.parse_line(line):
                    logger.new_invo()
                if(entry.parse_line(line)):
                    invo.entries.append(deepcopy(entry))
                    logger.new_entr()
        return invo

def _init_clean_up(tmp_dir='tmp'):
    """
    Create tmp directory, delete it first if it already exists, if possible
    
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
    Extract the pdf files from the zip, there is no sanity check for the arguments
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

    :return: list of pdf file path
    :rtype: list of str
    """
    pdf_list = []
    for dir_path, _dummy, file_list in os.walk(dir):
        for filename in file_list:
            if filename.endswith(extension):
                pdf_list.append(os.path.join(dir_path, filename))
    return pdf_list

def extract_invoces(pdf_list, logger):
    """
    Get the invoices from the pdf files in th pdf_list
    Wrapper around the pdf2rawtxt call
    
    :param list pdf_list: List of pdf files path to process.
    :param logger: :class:`StatLogger`, collect statistical data about parsing

    :return: list of invoices
    :rtype: list of :class:`Invoice`
    """
    invoice_list = []
    for pdfile in pdf_list:
        invoice_list.append(pdf2rawtxt(pdfile, logger))
    return invoice_list

def _post_clean_up(tmp_dir='tmp'):
    """
    Cleanup after execution, remove the extracted zip file and tmp directory
    
    :param str tmp_dir: Temporary directory to clean_up
    """
    shutil.rmtree(tmp_dir)
    

def invoices2xlsx(invoices, dir='', name='Invoices01.xlsx'):
    """
    Write invoice information to xlsx template file. Go through every invoce and
    write them out. Simple. Utilizes the xlsxwriter module

    :param invoices list of Invocie: Representation of invoices from the pdf files
    """
    workbook = xlsxwriter.Workbook(os.path.join(dir,name))
    worksheet_invo = workbook.add_worksheet()
    worksheet_entr = workbook.add_worksheet()
    row_invo = col_invo = row_entr = col_entr = 0
    
    labels = ["Invoice Number", "Date of Invoice", "Payment Date", "Amount"]
    positions = [NO_INDENT, ORIG_DATE_INDENT, PAY_DUE_INDENT,
                 TOTAL_SUM_INDENT]
    row_invo, col_invo = list2row(worksheet_invo, row_invo,
                                             col_invo, labels, positions)
    
    labels = ["Invoice Number"] + list(EntryTuple._fields)
    row_entr, col_entr = list2row(worksheet_entr, row_entr,
                                             col_entr, labels)
    #[TODO] there is no specification how to write out invocie entries yet
    for invo in invoices:
        row_invo, col_invo = invo.xlsx_write(worksheet_invo, row_invo, col_invo)
        for entr in invo.entries:
            row_entr, col_entr = entr.xlsx_write(worksheet_entr, row_entr, col_entr)

    workbook.close()

def do_it( src_name, dst_dir='', xlsx_name='Invoices01.xlsx',
           tmp_dir='tmp', file_extension='.pdf'):
    """
    Main script to manage the zip to xls process. It is responsible to create/cleanup
    temporary directories and files. After zip extraction, seraches every file which
    ends with `file_extension` Then it builds up a list of invoices and writes them
    to xlsx format.

    :param str src_name: path to the zip file to extract
    :param str dst_dir: path to the directory to put the generated xlsx file by default
        the cwd
    :param str tmp_dir: temporary directory to work in. **This directory is erased
        at the beginning of the script** By default it is `tmp`
    :param str file_extension: the file extension to use during file selection. By
        default it is `.pdf`
    :param str xlsx_name: Name of the oputput file
    """
    _init_clean_up(tmp_dir)

    extract_zip(src_name, tmp_dir)

    pdf_list = get_pdf_files(os.path.join(os.getcwd(),tmp_dir), file_extension)

    logger = StatLogger()
    
    invoice_list = extract_invoces(pdf_list, logger)

    _post_clean_up(tmp_dir)

    invoices2xlsx(invoice_list, dst_dir)
    print(logger)
    print("script has been finished")
    return logger
        

def main():
    do_it(SRC_NAME, dst_dir=DST_DIR, xlsx_name=XLSX_NAME,
          tmp_dir=TMP_DIR, file_extension=FILE_EXTENSION)


if __name__ == '__main__': main()

