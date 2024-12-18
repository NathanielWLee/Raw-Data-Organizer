clone repo into local directory

copy import_lib.bat anywher on your machine and execute that file

this will ask to install python, if you already have it, hit cancel and then yes

then, it will ask to give a command: just hit space and then enter

locate the \Users\%username%\RDO\Main folder that was just created and copy the main.py program into that folder

all that should be in the Main folder is an RDO executable and a main.py file

execute RDO and it should prompt for a CSV file

make sure that when choosing your raw data file that it is formated correctly: CSV, one sheet only, and contains the standardized column headers

if done correctly, an outputExcel Excel sheet and a log text file should appear in the Main folder

the log text file is for storing any errors that may occur during the last executable which inlcude:
  raw data file is still open, file is not a CSV format, or spreadsheet contains more than one sheet

the outputExcel sheet should contain a table and a bar chart representing the totals for each service total for each month including the grand totals

the second sheet includes more in depth numbers that represent the request types for each of those service types

these tables and charts can be copied into any Microsoft supported service such as PowerPoint, Word, OneNote
