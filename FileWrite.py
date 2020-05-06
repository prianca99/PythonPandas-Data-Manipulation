import pandas as pd
import xlsxwriter
import pyodbc
import datetime
import pandas as pd
from configparser import ConfigParser
from pathlib import Path
import os
import logging
import numpy as np


# print = logging.info
config = ConfigParser()
parser= ConfigParser()
#============================creating log file===============================================================

logging.basicConfig(filename='data_normalization_log',
                    format='%(asctime)s:%(levelname)s:%(lineno)d:%(message)s',
                    filemode='w+')

#Creating an object
logger = logging.getLogger()

#Setting the threshold of logger to DEBUG
logger.setLevel(logging.DEBUG)




server = ''
database = ''
username = ''
password = ''
conn = pyodbc.connect('DRIVER={ODBC Driver 17 for SQL Server};SERVER='+server+';DATABASE='+database+';UID='+username+';PWD='+ password)
cursor = conn.cursor()
script1 = """select * from [dbo].[table1]"""
script2= """select * from [dbo].[table2]"""
script3="""select * from [dbo].[table3]"""
script4="""select * from [dbo].[table4]"""

logging.info("reading the files from DB at : " +str(datetime.datetime.now()))

finalfile_df = pd.read_sql_query(script1, conn)
summary_df = pd.read_sql_query(script2, conn)
norules_df = pd.read_sql_query(script3, conn)
proccessed_rules=pd.read_sql_query(script4,conn)

#delete the blank DB table columns
finalfile_df.drop(['newcol1','newcol2','newcol3','newcol4','newcol5','newcol6','newcol7','newcol8','newcol9','newcol10','misc2','misc3'] ,axis=1, inplace=True)
proccessed_rules.drop(['misc2','misc3'] ,axis=1, inplace=True)

print("all db tables reading done at: " +str(datetime.datetime.now()))
logging.info("all db tables reading done at: " +str(datetime.datetime.now()))
#logging.info("db tables:" , finalfile_df,summary_df,norules_df,proccessed_rules)
logging.info(finalfile_df)
logging.info(summary_df)
logging.info(norules_df)
logging.info(proccessed_rules)


#=========writing new rules file==========================================
proccessed_rules['Standard Application Name']=proccessed_rules['Standard Application Name'].replace('DisplayAppFound but standardApp Blank',np.nan)
rulefile_name='Processed_RuleFile' + '.xlsx' 
sh1='Processed_Rules'
writer1 = pd.ExcelWriter(rulefile_name,engine='xlsxwriter')
proccessed_rules.to_excel(writer1, sheet_name=sh1, startrow=1, header=False)
workbook  = writer1.book
worksheet = writer1.sheets[sh1]
header_format = workbook.add_format({
        'bold': True,
        'text_wrap': True,
        'valign': 'top',
        'fg_color': '#D7E4BC',
        'border': 1})
for col_num, value in enumerate(proccessed_rules.columns.values):
    worksheet.write(0, col_num + 1, value, header_format)
    
writer1.save()


#===============writing the processed file================

writer = pd.ExcelWriter('11march.xls', engine='xlsxwriter')
k=1
sh='Normalized_Data'
r,c=finalfile_df.shape
if r >= 1048500:
    row_limit = 1048500
    for i in range(0, r, row_limit):
        print("sheet " + sh+str(k) + " writing in progress at " +str(datetime.datetime.now()))
        logging.info("sheet " + sh+str(k) + " writing in progress at " +str(datetime.datetime.now()))
        finalfile_df.iloc[i:i+row_limit].to_excel(writer,sh+str(k),startrow=1, header=False)
        workbook  = writer.book
        worksheet = writer.sheets[sh+str(k)]
        k=k+1
        header_format = workbook.add_format({
                'bold': True,
                'text_wrap': True,
                'valign': 'top',
                'fg_color': '#D7E4BC',
                'border': 1})
        for col_num, value in enumerate(finalfile_df.columns.values):
            worksheet.write(0, col_num + 1, value, header_format)
        
else:
    row_limit = 1048500
    for i in range(0, r, row_limit):
        print("sheet " + sh+str(k) + " writing in progress at " +str(datetime.datetime.now()))
        logging.info("sheet " + sh+str(k) + " writing in progress at " +str(datetime.datetime.now()))
        finalfile_df.iloc[i:i+row_limit,].to_excel(writer,sh+str(k),startrow=1, header=False)
        workbook  = writer.book
        worksheet = writer.sheets[sh+str(k)]
        k=k+1
        header_format = workbook.add_format({           
                'bold': True,
                'text_wrap': True,
                'valign': 'top',
                'fg_color': '#D7E4BC',
                'border': 1})
        for col_num, value in enumerate(finalfile_df.columns.values):
                worksheet.write(0, col_num + 1, value, header_format)
    
print("final file writing done")
print("final raw data writing done at: " +str(datetime.datetime.now()))
logging.info("processing done at: " +str(datetime.datetime.now()))
del finalfile_df

#finalfile_df.to_excel(writer, sheet_name='Normalized_Data')
summary_df.to_excel(writer, sheet_name='Summary')
norules_df.to_excel(writer, sheet_name='Rules_Not_found')
print("summary and rules not found writing done at: " +str(datetime.datetime.now()))
logging.info("summary and rules not found " +str(datetime.datetime.now()))

workbook  = writer.book
workbook.filename = 'Finalfile.xlsm'
worksheet = workbook.add_worksheet('pivot')
worksheet.set_column('A:A', 30)
worksheet.set_column('A:A', 30)
workbook.add_vba_project('./vbaProject.bin')
worksheet.write('A3', 'Press the button to get Pivot.')
worksheet.insert_button('B3', {'macro': 'newm','caption': 'Press Me','width': 80,'height': 30})
workbook.close()
print("pivot macros writing done at: " +str(datetime.datetime.now()))
logging.info("pivot macros writing done at: " +str(datetime.datetime.now()))