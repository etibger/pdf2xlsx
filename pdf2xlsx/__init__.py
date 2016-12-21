""" setup """
__all__ = ["pdf2xlsx"]
from .pdf2xlsx import do_it
from .gui import main as gui_main

def main():
    gui_main()
    

