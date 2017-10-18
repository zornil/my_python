#coding:utf-8
import xlrd
import xlwt
import numpy as np
#拷贝文件，并将文件名改为my_data.xls
book = xlrd.open_workbook('my_data.xlsx')
new_book = xlwt.Workbook()
new_sheet = new_book.add_sheet('sheet1')

sheet = book.sheet_by_index(0)
'n_rows = sheet.nrows'
'n_cols = sheet.ncols'
#获取第一列的行数n
cols_data0 = sheet.col_values(0)
n = 0
for i in range(0, np.size(cols_data0)):
    if cols_data0[i] != '':
        n += 1
print'n = ', n
new_cols_data = [[]*6]*n
new_cols_data[0] = cols_data0[0:n]
print('**计算每一列幅值**')
#计算每一列的幅值
pos = 0
for i in range(1, 7, 2):
    'print(i)'
    cols_data = sheet.col_values(i)
    new_sheet.write(0, pos + 7, np.mean(cols_data[0:n]))
    new_sheet.write(1, pos + 7, np.std(cols_data[0:n]))
    value = np.fft.fft(cols_data[0:n])
    pos += 1
    new_cols_data[pos] = np.sqrt(np.real(value)**2 + np.imag(value)**2)

print('**幅值计算完成**\n\n**开始处理无规则数据，求解mean，std**')

cols_data6 = sheet.col_values(6)
cols_data7 = sheet.col_values(7)
cols_data_mean1 = []
'cols_data_std1 = []'
count = 0
pos = 0
num = 0
for i in range(0, np.size(cols_data6)):
    if cols_data6[i] != '':
        num += 1
print'num = ', num
for i in range(0, num):
    if cols_data6[pos] == cols_data6[i]:
        count += 1
        if i == num - 1:
            '''print(i+1,'   ',count)
            print(cols_data6[num-count:num])'''
            cols_data_mean1.append(np.mean(cols_data7[num - count:num]))
            'cols_data_std1.append(np.std(cols_data7[num-count:num]))'
    elif cols_data6[pos] != cols_data6[i]:
        pos = i
        '''
        print('i = ',i)
        print('pos-count = ',pos-count)
        print('pos = ',pos)
        print(cols_data6[pos-count:pos])
        '''
        cols_data_mean1.append(np.mean(cols_data7[pos-count:pos]))
        'cols_data_std1.append(np.std(cols_data7[pos-count:pos]))'
        count = 1
new_sheet.write(0, 10, np.mean(cols_data_mean1[0:n]))
new_sheet.write(1, 10, np.std(cols_data_mean1[0:n]))
value1 = np.fft.fft(cols_data_mean1)
new_cols_data[4] = np.sqrt(np.real(value1)**2 + np.imag(value1)**2)

print('**第七列数据处理完成**\n\n**开始处理第九列数据**')
cols_data8 = sheet.col_values(8)
cols_data9 = sheet.col_values(9)
cols_data_mean2 = []
cols_data_std2 = []
count = 0
pos = 0
num = 0
for i in range(np.size(cols_data8)):
    if cols_data8[i] != '':
        num += 1
print'num = ', num
for i in range(0, num):
    if cols_data8[pos] == cols_data8[i]:
        count += 1
        if i == num - 1:
            '''print(i + 1,'   ',count)
            print(cols_data8[num - count:num])'''
            cols_data_mean2.append(np.mean(cols_data9[num - count:num]))
            cols_data_std2.append(np.std(cols_data9[num - count:num]))
    elif cols_data8[pos] != cols_data8[i]:
        pos = i
        '''
        print('i = ',i)
        print('pos-count = ',pos-count)
        print('pos = ',pos)
        print(cols_data8[pos-count:pos])
        '''
        cols_data_mean2.append(np.mean(cols_data9[pos - count:pos]))
        cols_data_std2.append(np.std(cols_data9[pos - count:pos]))
        count = 1
new_sheet.write(0, 11, np.mean(cols_data_mean2[0:n]))
new_sheet.write(1, 11, np.std(cols_data_mean2[0:n]))
value2 = np.fft.fft(cols_data_mean2)
new_cols_data[5] = np.sqrt(np.real(value2)**2 + np.imag(value2)**2)
print('**第九列数据处理完成**\n\n**开始保存数据到excel文件中**')

#将处理后的数据写入新的文件中
'''
title = ['等差数列','幅值','幅值','幅值','mean','mean-std','mean','mean-std']
for i in range(0,8):
    new_sheet.write(0,i,)
'''
#2~901
new_sheet.write(0, 6, 'mean')
new_sheet.write(1, 6, 'std')
for i in range(0, 6):
    for j in range(0, n-900-1):
        new_sheet.write(j, i, new_cols_data[i][j+1])
#保存处理后数据到excel文件中
new_book.save('my_data_process.xls')

'print(new_cols_data)'
print('数据处理完成，保存在文件my_data_process.xls中')
