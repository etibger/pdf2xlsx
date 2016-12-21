#!/usr/bin/env python
import sys
from cx_Freeze import setup, Executable

#from distutils.core import setup, Command

if sys.version_info < (3, 5, 0):
    warn("The minimum Python version supported by pdf2xlsx is 3.5.")
    exit()

build_exe_options = {
    "packages": [
        "os",
        "shutil",
        "zipfile",
        "re",
        "collections",
        "PyPDF2",
        "xlsxwriter"]}

base = None
if sys.platform == "win32":
    base = "Win32GUI"

    
long_description = """
    Extract zip, search for pdf, get invoices from pdf, write them to xlsx file
"""

setup(
        name="pdf2xlsx",
        version="0.1",
        description="Invoice extraction from zip",
        options = {"build_exe": build_exe_options},
        executables = [Executable("pdf2xlsx.py", base=base)],
        long_description=long_description,
        author="Tibor Gerlai",
        author_email="tibor.gerlai@gmail.com",
        maintainer="Tibor Gerlai",
        maintainer_email="tibor.gerlai@gmail.com",
        url="https://github.com/etibger/pdf2xlsx",
        classifiers = [
            "Development Status :: 5 - Production/Stable",
            "Intended Audience :: Developers",
            "License :: OSI Approved :: BSD License",
            "Programming Language :: Python :: 3",
            "Operating System :: Windows",
            "Topic :: Software Development :: Libraries :: Python Modules",
            ],
        packages=["pdf2xlsx"],
    )
