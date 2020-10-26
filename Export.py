import pyodbc 
import pandas as pd
import time
import xlrd
import openpyxl

## Get time of execution 
start_time = time.time()

## Define path for result file
excel_file = r'C:\Users\Desktop\GitHub\2.Pyodbc\Export.xlsx'

## Connection Properties
user = 'user_001'
password = 'password_001'
server = 'server_001'
database = 'database_001'
conn = pyodbc.connect('Driver={SQL Server};'
                      'Server='+server+';'
                      'Database='+database+';'
                      'Trusted_Connection=no;'
                      'uid='+user+';'
                      'pwd='+password+';'
                      )
pyodbc.pooling = False
cursor = conn.cursor()
cursor.fast_executemany = True

query = ''' SELECT 
              [UF], [NUM_ENVIO], CONVERT(INT,[QTDE_INFO]), 
              [INVOICE_NUMBER], [VALOR], [MES], [REGISTRO_ID]
              
            FROM [DataBase].[dbo].[Tabela] 
            WHERE [MES] in ('setembro_20')
        '''
cursor.execute(query)

columns = [column[0] for column in cursor.description]

df = pd.DataFrame(columns = columns)

## Appending the cursor rows to a DataFrame
for row in cursor:
       line = {}
       for index,value in enumerate(row):
              line[columns[index]] = value
       df = df.append(line, ignore_index=True)

## Convert to float Money Value      
df['VALOR'] = df['VALOR'].astype(float)

## Exporting the DataFrame to an Excel Sheet
df.to_excel (excel_file, index = None, header=True, engine="xlsxwriter", float_format="%.3f") 
conn.close()
print("--- %s seconds ---" % (time.time() - start_time))
