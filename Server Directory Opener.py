import os
import xlrd

book = xlrd.open_workbook(filename='Sites.xlsx')
sheet = book.sheet_by_index(0)
col1 = []
col2 = []
col1 = sheet.col_values(0)
col2 = sheet.col_values(1)

for i in range(len(col1)):
    col1[i] = col1[i].lower()
for i in range(len(col2)):
    col2[i] = col2[i].lower()    

print("Open server Directory app by Jerry :D")
print('Enter "quit" to quit')
while True:
    x = input('Enter your site location you want to open: ')
    if x == "quit":
        break
    elif x.lower() in col1 or x.lower() in col2:
        for i in range(0, sheet.nrows):
            a = sheet.cell_value(i,0)  #Site name
            b = sheet.cell_value(i,1)  #Abbrv names
            if x.lower() == a.lower() or x.lower() == b.lower():
                c = sheet.cell_value(i,2)
                print(c)
                os.startfile(c,"open")
    else:
        print("Invalid Entry.")
        continue
        

