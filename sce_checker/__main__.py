#!/usr/bin/env python
# -*- coding: utf-8 -*-

import argparse
from datetime import datetime
from pathlib import Path
import sys

try:
    import openpyxl as _
except ImportError:
    print("openpyxl isn't installed!", file=sys.stderr)
    print("Install using the command 'pip install openpyxl'", file=sys.stderr)
    sys.exit(1)

from . import zip_preparer
from .outputers import outputer_csv
from .inputters import inputter_excel
from .ui import main_window

def checker_main():
    parser = argparse.ArgumentParser(prog='sce-checker')
    parser.add_argument('-E', '--extract', dest='extract', metavar='PATH', type=Path,
                        help='Extract files.zip')
    parser.add_argument('-D', '--dest-dir', dest='dest_dir', metavar='PATH', type=Path,
                        help='Destination dir for extraction')
    parser.add_argument('-c', '--csv', dest='csv', metavar='PATH', type=Path,
                        help='csv file')
    parser.add_argument('-x', '--excel', dest='excel', metavar='PATH', type=Path,
                        help='Excel file')
    args = parser.parse_args()

    if args.extract and args.dest_dir:
        zip_preparer.prepare_check_xlsm(args.extract, args.dest_dir)
    elif args.csv and args.excel:
        csv = outputer_csv.OutputCSV(args.csv)
        for s in inputter_excel.InputterExcel(args.excel):
            if not csv.save_student(s):
                print(s)
            else:
                print(datetime.now(), s.name)
        csv.save(args.csv.parent / 'new.csv')
    else:
        main_window.main()

if __name__ == '__main__':
    checker_main()
