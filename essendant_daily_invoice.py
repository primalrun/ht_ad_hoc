import sys
import pypyodbc
import csv
import os
from tkinter.filedialog import askopenfilenames
import pandas as pd


def str_to_float(str_temp):    
    try:        
        str_temp = str_temp.replace("$", "")
        str_temp = str_temp.replace(",", "")
        float_return = float(str_temp)        
    except:
        float_return = 0
    return float_return    

def str_to_int(str_temp):    
    try:
        str_temp = str_temp.replace(",", "")
        int_return = int(str_temp)        
    except:
        int_return = 0
    return int_return

file_names = askopenfilenames()

if len(file_names) == 0:
    print('No File Selected, Process Cancelled')
    sys.exit()
    
path = os.path.dirname(file_names[0]) + '/'

#SQL Server Connection
connection_string = """Driver={SQL Server Native Client 11.0};
                        Server=TNDCSQL02;
                        Database=Playground;
                        Trusted_Connection=yes;
                    """
                    
connection = pypyodbc.connect(connection_string)
cur = connection.cursor()

for f in range(0, len(file_names)):
    xl = pd.ExcelFile(file_names[f])
    sheet_names = xl.sheet_names
    
    try:
        invoice_sheet_name = [s for s in sheet_names if 'Invoice' in s]
        df = xl.parse(sheetname=invoice_sheet_name[0], index_col=None, 
                      skiprows=[1])
        df.to_csv(path_or_buf=path + 't1.csv', index=False, header=False)
        csv_file = path + 't1.csv'
        
        if len(df.columns) != 28:
            print("Column Header Count is not 28 in Invoice Sheet")                   
            sys.exit()
    except:
        print(os.path.basename(file_names[f]) + 
              " Error importing Invoice sheet, Process Cancelled")
        cur.close()
        connection.close()
        sys.exit()    

        
    print(os.path.basename(file_names[f]) + ' attempted')

    with open(csv_file, 'r') as f:
        reader = csv.reader(f)
        purchase_data = list(reader)
    
    
    #remove 25th column, null, extra blank columns    
    x1 = [x for x in purchase_data[0] if len(x) > 0]
    col_count = len(x1)
        
    t1 = [r[:25] + r[26:col_count] for r in purchase_data if len(r[4]) > 0
          if r[0][0:2] != 'SH' 
          if r[4][:2] in ('OE', 'WO')] 
    col_count = len(t1[0])
    
    
    #delete temp csv file
    os.remove(path + 't1.csv')
    
    date_file = t1[0][3]
    
    #format lists 0-27 columns
    f1 = [r[:8] for r in t1] 
    f2 = [[str_to_int(e) for e in r[8:10]] for r in t1]
    f3 = [r[10] for r in t1]
    f4 = [[str_to_float(e) for e in r[11:13]] for r in t1]
    f5 = [r[13] for r in t1]
    f6 = [[str_to_float(e) for e in r[14:19]] for r in t1]
    f7 = [r[19:] for r in t1]
    
    #combine formatted lists into list of tuples
    t2 = [] 
    
    for r in range(0, len(t1)):
        x1 = tuple(f1[r] + f2[r] + [f3[r]] + f4[r] + [f5[r]]  + f6[r] + f7[r])
        t2.append(x1)
    
    #delete import date from sql table
    sql = ("delete from Playground.[myop\jason.walker].essendant_daily_invoice " 
           "where invoice_date = '{d1}'").format(d1 = date_file)
    
    cur.execdirect(sql)
    cur.commit()
    
    #import from list
    var_temp = '?, ' * (col_count)    
    var_temp = var_temp[:-2]
    
    
    sql = ("insert into Playground.[myop\jason.walker].essendant_daily_invoice "
        "values ({v1})"
        ).format(v1= var_temp)    

    cur.executemany(sql, t2)    
    cur.commit()
    
cur.close()
connection.close()
    
print('Success: Essendant File(s) are loaded')
    
    

