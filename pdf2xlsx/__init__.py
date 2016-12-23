# -*- coding: utf-8 -*-
""" setup """
__version__ = '0.1.1'
__all__ = ["pdf2xlsx"]
from .pdf2xlsx import do_it
from .gui import main as gui_main

def main():
    gui_main()
    

