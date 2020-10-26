import pyodbc 
import pandas as pd
import time
import xlrd
import openpyxl

## python -m pip install library --user --no-warn-script-location --proxy http://proxy.company.enterprise:0000

## Get time of execution 
start_time = time.time()

## Define path for importable file
excel_file = r'C:\Users\Desktop\GitHub\2.Pyodbc\Export.xlsx'
df = pd.read_excel(excel_file, header=0, na_filter=False, sheet_name='Sheet1')
df2 = pd.DataFrame(df, columns = ['UF', 'NUM_ENVIO', 'QTDE_INFO', 'INVOICE_NUMBER', 'VALOR', 'MES', 'REGISTRO_ID'])

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

query1 = ''' UPDATE [DataBase].[dbo].[Tabela] with (updlock, serializable)
             SET  [UF] = ?,
                  [NUM_ENVIO] = ?,
                  [QTDE_INFO] = ?, 
                  [INVOICE_NUMBER] = ?, 
                  [VALOR] = ?, 
                  [MES] = ?
             WHERE [REGISTRO_ID] = ?  
         '''     

query2 = ''' INSERT INTO [DataBase].[dbo].[Tabela] with (updlock, serializable)
            (  [UF],
               [NUM_ENVIO],
               [QTDE_INFO], 
               [INVOICE_NUMBER], 
               [VALOR], 
               [MES]
            )
              values 
              (?, ?, ?, ?, ?, ?) 
         '''

## Filter DataFrame, one with pre-existing values, one with new values
df_existing_values = df2[(df2.REGISTRO_ID != '') | (df2.REGISTRO_ID != 'None') | (df2.REGISTRO_ID != 'NULL') ]
df_new_values = df2[(df2.REGISTRO_ID == '') | (df2.REGISTRO_ID == 'None') | (df2.REGISTRO_ID == 'NULL') ]

## Update pre-existing values
for row in df_existing_values.itertuples():
   cursor.execute(query1, row.UF, row.NUM_ENVIO, row.QTDE_INFO, row.INVOICE_NUMBER, row.VALOR, row.MES, row.REGISTRO_ID)
   conn.commit()

## Insert new values
for row in df_new_values.itertuples():
   cursor.execute(query2, row.UF, row.NUM_ENVIO, row.QTDE_INFO, row.INVOICE_NUMBER, row.VALOR, row.MES, row.REGISTRO_ID)
   conn.commit()   

conn.close()
print("--- %s seconds ---" % (time.time() - start_time))