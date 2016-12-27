# -*- coding: utf-8 -*-
""" setup """
__version__ = '1.0.0'
__all__ = ["__main__", "pdf2xlsx", "gui", "logger", "config"]

"""
By default load the managment function and the gui when the module is loaded.
"""
from .pdf2xlsx import do_it
from .gui import main as gui_main
