# -*- coding: utf-8 -*-

from re import sub as re_sub
from typing import NamedTuple

XLSM_ROW_NAMES = 4
XLSM_COL_NAMES_START = 5
XLSM_ROW_TITLE_TOTAL_MARKS = 'Total Mark'
XLSM_ROW_TITLE_CLEAR_CHAR = 'Clear Char'

CSV_COL_TITLE_NAME = 'שם מלא'
CSV_COL_TITLE_GROUP = 'קבוצה'
CSV_COL_TITLE_GRADE = 'ציונים'
CSV_COL_TITLE_COMMENTS = 'הערות למשוב'
CSV_COL_TITLE_LAST_UPDATE = 'עדכון אחרון (ציון)'

HTML_INDENT = '&emsp;'
HTML_NEWLINE = '<br/>\n'

def str_cleanup(s: str) -> str:
    return re_sub(r'\s\s+', ' ', s.strip().replace('\'', '').replace('"', ''))

class StudentGrade(NamedTuple):
    name: str
    grade: str
    comment: str
