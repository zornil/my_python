import xlrd
import xlwt
import numpy
from datetime import datetime

book = xlrd.open_workbook("myfile.xls")
print("The number of worksheets is {0}".format(book.nsheets))
print("Worksheet name(s): {0}".format(book.sheet_names()))
sh = book.sheet_by_index(0)
print("{0} {1} {2}".format(sh.name, sh.nrows, sh.ncols))
'print("Cell D30 is {0}".format(sh.cell_value(rowx=29, colx=3)))'
'print("Cell D30 is {0}".format(sh.cell_value(29,3)))'
for rx in range(sh.nrows):
    print(sh.row(rx))

files = r'my_test.xlsx'
book = xlrd.open_workbook(files)

sheet_name = book.sheet_names()[0]
print(sheet_name)
sheet = book.sheet_by_name(sheet_name)
sheet0 = book.sheet_by_index(0)

n_rows = sheet.nrows
n_clos = sheet.ncols
print(n_rows)
print(n_clos)
row_data = sheet.row_values(0)
col_data = sheet.col_values(0)

cell_value = sheet.cell_value(0,1)

print(cell_value)
cell_value2 = sheet.cell(0,1)

x = numpy.linspace(0,2*numpy.pi,30)
x1 = [1,2,3,4,5,6]
y = numpy.mean(x1)
y1 = numpy.var(x1)
print(y)
print(y1)
wave = numpy.cos(x)
print(wave)
transformed = numpy.fft.fft(wave)
print(transformed)

print(numpy.all(numpy.abs(numpy.fft.ifft(transformed)-wave) < 10**-9))

wb = xlwt.Workbook()
ws = wb.add_sheet('sheet1')
ws.write(0,0,123.45)
ws.write(1,0,datetime.now())
ws.write(2,0,1.1)
for i in range(0,30):
    ws.write(i,1,wave[i])

for j in range(0,6):
    ws.write(j,2,x1[j])
value = numpy.sqrt(numpy.real(transformed)**2 + numpy.imag(transformed)**2)
print(value)
for i in range(0,30):
    ws.write(i,3,value[i])

v = numpy.complex(2,4)
print(numpy.real(v)**2)
print(numpy.imag(v)**2)
v1 = numpy.sqrt(numpy.real(v)**2 + numpy.imag(v)**2)
print v1
print(numpy.size(wave))

wb.save('example.xls')
