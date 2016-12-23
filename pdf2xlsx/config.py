# -*- coding: utf-8 -*-
"""
Configuration structure, loading and storing
"""
from json import dumps, loads
from collections import OrderedDict
import os


"""
Calculate path to store config at
"""
try:
    HOME = os.environ['HOME']
except KeyError as e:
    print("Home envirnment variable not set, use current working directory for config")
    HOME = os.getcwd()

CONF_DEFAULT_PATH = os.path.join(HOME,'.pdf2xlsx','config.txt')

class JsonDict(OrderedDict):
    """
    Simple dict class extended with serialization functions, store and load
    """
    def store(self, path=CONF_DEFAULT_PATH):
        """
        Store the actual configuration to config file (path)

        :param str path: Path and filename of the config file
        """
        with open(path, 'w', encoding="utf-8") as conf_out:
            conf_out.write(dumps(self, indent=4, ensure_ascii=False))

    def load(self, path=CONF_DEFAULT_PATH):
        """
        Update the config from the config file (path)

        :param str path: Path and filename of the config file
        """
        with open(path, 'r', encoding="utf-8") as conf_in:
            self.update(loads(conf_in.read()))

_keys = ['value', 'text']
config = JsonDict([
    ('tmp_dir', dict(zip(_keys, ['tmp', 'tmp dir']))),
    ('file_extension', dict(zip(_keys, ['pdf', 'file ext']))),
    ('xlsx_name', dict(zip(_keys, ['invoices.xlsx', 'xlsx name:']))),
    ('invo_header_ident', dict(zip(_keys, [[1,2,3,4], 'invo header pos']))),
    ('ME', dict(zip(_keys, [['Pár', 'Darab'], 'Me category']))),
])

def init_conf(conf=config, cfg_path=CONF_DEFAULT_PATH):
    """
    Load the config file from $HOME/pdf2xlsx/cfg_name. If it doesn't exist
    try to create it. First create the pdf2xlsx directory, and then write out
    the default config
    """
    try:
        conf.load(cfg_path)
    except FileNotFoundError as e:
        cfg_dir, cfg_name = os.path.split(cfg_path)
        try:
            os.mkdir(cfg_dir)
        except FileExistsError:
            pass
        finally:
            conf.store(cfg_path)
