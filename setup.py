#!/usr/bin/env python
import sys

try:
    from setuptools import setup, Command
except ImportError:
    from distutils.core import setup, Command

if sys.version_info < (3, 5, 0):
    warn("The minimum Python version supported by pdf2xlsx is 3.5.")
    exit()
    
long_description = """
    Extract zip, search for pdf, get invoices from pdf, write them to xlsx file
"""

setup(
        name="pdf2xlsx",
        version="0.1",
        author="Tibor Gerlai",
        author_email="tibor.gerlai@gmail.com",
        url="https://github.com/etibger/pdf2xlsx",
        packages=['pdf2xlsx']
        license='MIT'
        description="Invoice extraction from zip",
        long_description=long_description,
        maintainer="Tibor Gerlai",
        maintainer_email="tibor.gerlai@gmail.com",
        classifiers = [
            "Development Status :: 2 - Pre-Alpha",
            "Environment :: Console",
            "Environment :: Win32 (MS Windows)",
            "Intended Audience :: End Users/Desktop",
            "License :: OSI Approved :: MIT License",
            "Natural Language :: English",
            "Operating System :: Microsoft :: Windows",
            "Programming Language :: Python :: 3.5",
            "Programming Language :: Python :: 3 :: Only",
            "Topic :: Office/Business :: Financial :: Spreadsheet",
            ],
        packages=["pdf2xlsx"],
    )
