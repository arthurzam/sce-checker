

import csv
import functools
from collections import defaultdict
from datetime import datetime
from pathlib import Path
from typing import Iterable

from .. import consts


class OutputCSV:
    def __init__(self, old_csv: Path):
        with old_csv.open('r', encoding='utf-8') as old_csv_file:
            self.data = tuple(csv.reader(old_csv_file, delimiter=',', quotechar='"'))
        self.names_dict = defaultdict(list)
        for row in self.data[1:]:
            self.names_dict[consts.str_cleanup(row[self.column_name])].append(row)

    def save(self, new_csv: Path):
        with new_csv.open('w', encoding='utf-8') as new_csv_file:
            newcsv = csv.writer(new_csv_file, delimiter=',', quotechar='"', lineterminator='\n')
            newcsv.writerows(self.data)

    @functools.cached_property
    def column_name(self) -> int:
        try:
            return self.data[0].index(consts.CSV_COL_TITLE_GROUP)
        except ValueError:
            ...
        return self.data[0].index(consts.CSV_COL_TITLE_NAME)

    @functools.cached_property
    def column_grade(self) -> int:
        return self.data[0].index(consts.CSV_COL_TITLE_GRADE)

    @functools.cached_property
    def column_comments(self) -> int:
        return self.data[0].index(consts.CSV_COL_TITLE_COMMENTS)

    @functools.cached_property
    def column_last_update(self) -> int:
        return self.data[0].index(consts.CSV_COL_TITLE_LAST_UPDATE)

    @property
    def names(self) -> Iterable[str]:
        return self.names_dict.keys()

    @property
    def ungraded_names(self) -> Iterable[str]:
        return (name for name, row in self.names_dict.items() if not row[self.column_grade])

    def save_student(self, student: consts.StudentGrade) -> bool:
        rows = self.names_dict[consts.str_cleanup(student.name)]
        for row in rows:
            row[self.column_grade] = student.grade
            row[self.column_comments] = f'<p dir="rtl" style="text-align: right;">\n{student.comment}\n</p>'
            row[self.column_last_update] = f'{datetime.now():%d/%m/%Y, %H:%M}'
        return bool(rows)
