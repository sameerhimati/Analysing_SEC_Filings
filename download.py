import excel
import requests
import os
import time
excel_path="File path here"  # Here is the path to where the excel file is
norexcel=excel.get_rows(excel_path,"Sheet1")  # Gets the number of rows
nocexcel=excel.get_columns(excel_path,"Sheet1")  # Gets the number of columns

for ros in range(236,237):
    fileurl='https://'+excel.read_data(excel_path,"Sheet1",ros,5)
    filename=str(excel.read_data(excel_path,"Sheet1",ros,1))+'.txt'
    print(filename)
    print(fileurl)
    import requests
    r = requests.get(fileurl)
    time.sleep(1)
    f = open(filename, 'wb')
    for chunk in r.iter_content(chunk_size=8192): 
        if chunk: # filter out keep-alive new chunks
            f.write(chunk)
    f.close()
    time.sleep(1)
   

