import logging
from pathlib import Path
import sys
from tkinter import *
from tkinter.ttk import *
from tkinter.filedialog import askdirectory, askopenfilename, asksaveasfilename
import tkinter.messagebox as msgbox
from typing import Tuple, List

from .. import consts
from ..zip_preparer import prepare_check_xlsm
from ..outputers import outputer_csv
from ..inputters import inputter_excel

def text_show(text: str) -> str:
    if sys.platform == 'linux':
        return text[::-1]
    else:
        return text

class SelectName(Toplevel):
    def __init__(self, master, student_name: str, possible_names: Tuple[str]) -> None:
        super().__init__(master)
        self.title("Select correct name")

        Label(self, text=f"Couldn't find student [{text_show(student_name)}]").pack(fill=X, padx=5, pady=5)
        Label(self, text="Select the corrent name:").pack(fill=X, padx=5, pady=5)

        self._selected = StringVar()
        self.value = ''
        OptionMenu(self, self._selected, text_show(possible_names[0]), *map(text_show, possible_names)).pack(fill=X, padx=5, pady=5)

        Button(self, text="Select", command=self.select).pack(padx=5, pady=5)
        Button(self, text="Cancel", command=self.destroy).pack(padx=5, pady=5)

    @property
    def selected(self) -> str:
        return text_show(self.value)

    def select(self):
        self.value = self._selected.get()
        self.destroy()


class MainPanel(Frame):
    def __init__(self) -> None:
        super().__init__()
        self.master.title("Voter")
        self.pack(fill=BOTH, expand=True)

        Button(self, text="Prepare Check", command=self.prepare_excel).pack(fill=X, padx=5, pady=5)
        Button(self, text="Fill Grading", command=self.fill_grading).pack(fill=X, padx=5, pady=5)

    def prepare_excel(self):
        files_zip = askopenfilename(parent=self, title='Select zip',
            filetypes=(('Zip file', '*.zip'), ('All files', '*.*'))
        )
        if not files_zip:
            return
        dest_dir = askdirectory(title='Select destination directory', parent=self)
        if not dest_dir:
            return
        try:
            students_count = prepare_check_xlsm(Path(files_zip), Path(dest_dir))
            msgbox.showinfo("check.xlsm ready",
                f'Check file ready at {dest_dir}/check.xlsm\nwith {students_count} students')
        except Exception as exc:
            msgbox.showerror("preparation failed", f'Sadly we failed with:\n{exc}')

    def fill_grading(self):
        check_file = askopenfilename(parent=self, title='Select check file',
            filetypes=(('Excel file', ('*.xlsx', '*.xlsm', '*.xls')), ('All files', '*.*'))
        )
        if not check_file:
            return
        csv_file = askopenfilename(parent=self, title='Select grading csv file',
            filetypes=(('csv file', '*.csv'), ('All files', '*.*'))
        )
        if not csv_file:
            return
        try:
            csv = outputer_csv.OutputCSV(Path(csv_file))
            mismatch_students: List[consts.StudentGrade] = []
            for student in inputter_excel.InputterExcel(Path(check_file)):
                if not csv.save_student(student):
                    mismatch_students.append(student)

            for student in mismatch_students:
                dialog = SelectName(self.master, student.name, tuple(csv.ungraded_names))
                self.master.wait_window(dialog)
                if new_name := dialog.selected:
                    csv.save_student(consts.StudentGrade(
                        name=new_name,
                        grade=student.grade,
                        comment=student.comment,
                    ))
                else:
                    msgbox.showwarning("preparation canceled", 'preparation canceled')
                    return

            # TODO: helper for fixing mismatching students
            dest_csv_file = asksaveasfilename(parent=self, title='Select output csv file',
                defaultextension='csv',
                filetypes=(('csv file', '*.csv'), ('All files', '*.*')),
            )
            if not dest_csv_file:
                return
            csv.save(Path(dest_csv_file))
        except Exception as exc:
            logging.error('fail', exc_info=exc)
            msgbox.showerror("preparation failed", f'Sadly we failed with:\n{exc}')


def main():
    root = Tk()
    MainPanel()
    root.mainloop()
