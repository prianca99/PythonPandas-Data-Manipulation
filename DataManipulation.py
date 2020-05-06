import pandas as pd
import numpy as np
import datetime
from pandas import ExcelWriter
from pandas import ExcelFile
import xlsxwriter
from configparser import ConfigParser
from pathlib import Path
import os
import logging


# print = logging.info
config = ConfigParser()
parser= ConfigParser()
configpath=os.getcwd() + '\\config.ini'
pa='r' + configpath
# parser.read('C:\\Users\\703255218\\Desktop\\new\\config.ini')

parser.read(configpath)
#xyz=parser.sections()

xyz=parser.options('PATHS')


#============================creating log file===============================================================

logging.basicConfig(filename='data_normalization_log',
                    format='%(asctime)s:%(levelname)s:%(lineno)d:%(message)s',
                    filemode='w+')

#Creating an object
logger = logging.getLogger()

#Setting the threshold of logger to DEBUG
logger.setLevel(logging.DEBUG)


print("Files read from config file: " +str(xyz))
logging.info("Files read from config file: " +str(xyz))

read_input1=parser.get('PATHS','rawdata')
read_input2=parser.get('PATHS','ruledata')
read_outputpath=parser.get('PATHS','Outputpath')
read_ruleoutputpath=parser.get('PATHS','RuleOutputpath')
fil=parser.get('PATHS','filters')



#==============reading of files=======================

print("reading the files at: " +str(datetime.datetime.now()))
logging.info("reading the files at: " +str(datetime.datetime.now()))

mainfile=pd.read_excel (read_input1, sheet_name=None)
rulefile=pd.read_excel (read_input2)

print("file reading done at: " +str(datetime.datetime.now()))
logging.info("file reading done at: " +str(datetime.datetime.now()))

z=len(mainfile)
emptydict={}
extra={}
df = pd.DataFrame()
summary={}
newabc={}
rulenotfound={}
ruledict={}
newdict={}


#=========writing new rules file==========================================
ruleslength=str(len(rulefile))
print("Number of rules in rulefile " + ruleslength)
logging.info("Number of rules in rulefile " + ruleslength)
rulefile['Display Application Name']=rulefile['Display Application Name'].str.strip()
rulefile = rulefile.drop_duplicates(keep='first')
newruleslength=str(len(rulefile))
print("Number of rules file after removing duplicates " +newruleslength)
logging.info("Number of rules file after removing duplicates " +newruleslength)


rulefile['Exception'] =np.where(rulefile['Display Application Name'].duplicated(keep=False),'1','0')
rulefile['Exception']= rulefile['Exception'].replace('1', 'yes')
rulefile['Exception']= rulefile['Exception'].replace('0', np.nan)
rulefile['Standard Application Name']= rulefile['Standard Application Name'].str.strip()
rulefile['Standard Application Name']= rulefile['Standard Application Name'].replace('', np.nan)
# rulefile['Standard Application Name']= rulefile['Standard Application Name'].fillna('No')
print(rulefile['Standard Application Name'].tail(10))
ruledict['Rules']=rulefile
rulefile_name='Processed_RuleFile' + '.xlsx' 
excel_file_rule = os.path.join(os.sep,read_outputpath,rulefile_name)
writer1 = pd.ExcelWriter(excel_file_rule, engine='xlsxwriter')

for i in ruledict:
    dfrule= ruledict[i]
    dfrule.to_excel(writer1, sheet_name=i, startrow=1, header=False)
    workbook  = writer1.book
    worksheet = writer1.sheets[i]
    header_format = workbook.add_format({
            'bold': True,
            'text_wrap': True,
            'valign': 'top',
            'fg_color': '#D7E4BC',
            'border': 1})
    for col_num, value in enumerate(dfrule.columns.values):
        worksheet.write(0, col_num + 1, value, header_format)
	
	
writer1.save()
rulefile['FOUND']='yes'
rulefile['Standard Application Name']= rulefile['Standard Application Name'].replace(np.nan,'Standard Application Name Missing in rule file(Display Application Name Found in Rule File)')
#===================processing the raw data=======================


file_name='FinalFile' + '.xlsx' 
excel_file_full = os.path.join(os.sep,read_outputpath,file_name)

filters={'Display Application Name': fil}
mac=[]
for i in filters:
	mac.append(filters[i])
	
keys=[]
for i in filters:
	keys.append(i)
	
key1=keys[0]

writer = pd.ExcelWriter(excel_file_full, engine='xlsxwriter')
for i in mainfile:
	j=pd.read_excel(read_input1, sheet_name=i, index=False)
	print(j[key1])
	print(fil)
	j[key1]=j[key1].replace(np.nan,'blank')
	print(j)
	j=j[j[key1].str.contains(fil)]
	j=j.rename(columns={'Display Name': 'Display Application Name'})
	j['Sounce_Domain']=i
	print("sheet " + i + " reading in progress at " +str(datetime.datetime.now()))
	logging.info("sheet " + i + " reading in progress at " +str(datetime.datetime.now()))
	df3=j.merge(rulefile, on='Display Application Name' , how='left')
	df3['FOUND']= df3['FOUND'].replace(np.nan, 'NO')
	rows=str(len(df3))
	print("no of rows in sheet: " + i + ", " + rows)
	logging.info("no of rows in sheet: " + i + ", " + rows)
	extra[i]=df3
	df = df.append(df3 ,sort=True)
	print("sheet " + i + " processing done, moving ahead")
	logging.info("sheet " + i + " processing done, moving ahead")
	logging.info(datetime.datetime.now())
	del df3

print("processing done..now will start writing new file at " + str(datetime.datetime.now()))
 
df1=df.pop('Sounce_Domain')
df['Sounce_Domain']=df1
print("printing dfsize")
print(df.size)
newdict['Normalized']=df
print("printing dict")
print(newdict)

#=========processing the summary tab=================

df['Standard Application Name']= df['Standard Application Name'].replace(np.nan, 'N/A(Display Application Name not found in rules)')
z=df.groupby([df['Standard Application Name'],df['Classification']]).size().to_frame('count').reset_index()
z.loc['Total'] = pd.Series(z['count'].sum(), index = ['count'])
df['Standard Application Name']= df['Standard Application Name'].replace('Standard Application Name Missing in rule file(Display Application Name Found in Rule File)',np.nan)
summary['snap']=z
logging.info(summary)
newdict.update(summary)

#===============processing the rules not found tab=============

filterrows=df['FOUND']=='NO'
newdfforrules=df[filterrows]
norules=newdfforrules['Display Application Name'].str.strip()
norules=norules.drop_duplicates(keep='first')

print("printing norules")

print("total number of rules not found " +str(len(norules)))

logging.info("printing norules")
logging.info(norules)
logging.info("total number of rules not found " +str(len(norules)))


norules=norules.to_frame().reset_index()
del norules['index']
newabc['Not Found Rules']=norules
newdict.update(newabc)

for i in newdict:
	print(i)

#===============writing the processed file================

for sh in newdict:
	df= newdict[sh]
	print("sheet: " + sh)
	print(df)
	logging.info("sheet: " + sh)
	logging.info(df)
	k=1
	r,c=df.shape
	if r >= 1048500:
		row_limit = 1048500
		for i in range(0, r, row_limit):
			df.iloc[i:i+row_limit].to_excel(writer,sh+str(k),startrow=1, header=False)
			workbook  = writer.book
			worksheet = writer.sheets[sh+str(k)]
			k=k+1
			header_format = workbook.add_format({
					'bold': True,
					'text_wrap': True,
					'valign': 'top',
					'fg_color': '#D7E4BC',
					'border': 1})
			for col_num, value in enumerate(df.columns.values):
				worksheet.write(0, col_num + 1, value, header_format)
			
			
			
	else:
		row_limit = 1048500
		for i in range(0, r, row_limit):
			df.iloc[i:i+row_limit,].to_excel(writer,sh,startrow=1, header=False)
			workbook  = writer.book
			worksheet = writer.sheets[sh]
			k=k+1
			header_format = workbook.add_format({	
					'bold': True,
					'text_wrap': True,
					'valign': 'top',
					'fg_color': '#D7E4BC',
					'border': 1})
			for col_num, value in enumerate(df.columns.values):
				worksheet.write(0, col_num + 1, value, header_format)
				

writer.save()
print("processing done at: " +str(datetime.datetime.now()))
logging.info("processing done at: " +str(datetime.datetime.now()))
