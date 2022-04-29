# Installation

## Windows

1. Check that you have Python installed with Tkinter support (one of the
   checkboxes during installation). NOTE: minimal version of Python is `3.9`.
2. Run the following command to install this utility:
  `python -m pip install https://github.com/arthurzam/sce-checker/archive/master.zip`
3. Hopefully it installed the dependency, otherwise install it manually using
  `python -m pip install openpyxl`
4. Check it was installed by running `python -m sce_checker -h` , and seeing a
   simple help message.
5. Run the command `WHERE pythonw` to find the location of PythonW on your PC.
   Possible value can be `C:\Program Files\Python310\pythonw.exe`
6. Now go to the Desktop and create a new shortcut, and fill the following
   target: `"[PATH to pythonw]" -m sce_checker`, while replacing with the PATH
   you found in step 5.
7. Double click the shortcut to view the basic GUI.

# Graphical User Interface Usage

The GUI is very simple, with only two features.

## Prepare Check

Opens an open file dialog, for which you should pass the ZIP file downloaded
from Moodle of all the assignment of the students. After selecting the ZIP
file, the user is asked for the target directory where to extract and prepare
the checker file.

After successful extraction and preparation, you can open the generated
`check.xlsm` file in that directory, and start grading the students.

## Fill Grading

Opens an open file dialogs, for accepting the filled `check.xlsm` file and
current grade file (CSV file). The utility will build the new grading, format
the comments, and finally ask the user where to save the resulting CSV file,
which then you can upload to Moodle as graded assignment.

In case there was a mismatch finding a student from `check.xlsm` in the CSV
file, a dialog will show which enables the user to select the correct name of
student (only students without a grade will be shown).
