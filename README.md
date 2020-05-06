The Script 'DataManipulation.py' does the following:
1. Takes the two excel files as an input from config.ini file.
2. First processed the file 1 (Rulesfile) and then write it.
3. this processed Rule file is given a input to perform lookup on common column 'Display Application Name' in both the input files. 
   Normalization of data is done then including filtering of the data.
4. File writing is done taking the consideration of limited number of rows in excel (10,48000 rows) in excel.


The Script 'FileWrite.py' does the following:

1. Write the data of three tables (table1,2,3) into one Macro-Enabled worksheet.
2. This Macro-enabled worksheet has the pivot sheet with all the underlying data.
3. An empty .xslm file is encoded to vbaextract.py which contains macros named 'newm'.
4. This vbaextract.py makes vba.Project.bin. You can find more here 
      'https://xlsxwriter.readthedocs.io/working_with_macros.html#macros"
