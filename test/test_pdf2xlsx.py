# -*- coding: utf-8 -*-
import pytest
import os
import pdf2xlsx.pdf2xlsx


def test_simple_run():
    pdf2xlsx.do_it(os.path.join('test','src.zip'), 'test')
    with open(os.path.join('test','Invoices.xlsx'), 'rb') as fd:
        content1 = fd.read()
    with open(os.path.join('test','Invoices01.xlsx'), 'rb') as fd:
        content2 = fd.read()
    
    assert content1 == content2
