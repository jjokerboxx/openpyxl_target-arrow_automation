import openpyxl
import time


#---------------------------LOADING-----------------------------------------------


print('Start!')
start = time.time()  # 시작 시간 저장

filename_target = '강진소방서 주택관리대장 2.xlsx'
filename_arrow = '기초수급자 명단(장흥군0857)2.xlsx'


wb_t = openpyxl.load_workbook(filename_target, data_only=True)
sheet_t = wb_t.worksheets[1]

wb_a = openpyxl.load_workbook(filename_arrow, data_only=True)
sheet_a = wb_a.worksheets[0]

print('load complete!')
print("Loading time :", time.time() - start)  # 현재시각 - 시작시간 = 실행 시간


#---------------------------WORKING-----------------------------------------------


#print(sheet2.cell(row=2283, column=21).value)
#print(sheet.cell(row=17744, column=56).value)

'''
if sheet2.cell(row=2283, column=21).value == sheet.cell(row=17744, column=56).value:
    sheet.cell(row=6, column=15).value = sheet2.cell(row=10, column=14).value
    print(sheet.cell(row=6, column=15).value)
    print('Success!!')
'''

counter = 0 
for i in range(3, sheet_a.max_row, 1):
    
    if sheet_a.cell(row= i, column=13).value and sheet_a.cell(row= i, column=14).value  is None:
        print('PASS')
        continue
    print('NEW CELL DETECTED')
    for t in range(6, sheet_t.max_row, 1):
        if sheet_a.cell(row=i, column=21).value == sheet_t.cell(row= t, column= 56).value:
            sheet_t.cell(row= t, column= 15).value = sheet_a.cell(row= i, column= 13).value
            sheet_t.cell(row= t, column= 18).value = sheet_a.cell(row= i, column= 14).value
            sheet_t.cell(row= t, column= 24).value = sheet_a.cell(row= i, column= 17).value
            sheet_t.cell(row= t, column= 23).value = sheet_a.cell(row= i, column= 16).value
            sheet_t.cell(row= t, column= 25).value = "안전문화조성"
            counter += 1
            print(counter)
        

print('Success!!')
print("Work time :", time.time() - start)  # 현재시각 - 시작시간 = 실행 시간



#-----------------------------SAVING---------------------------------------------


wb_t.save('33강진소방서 주택관리대장 3.xlsx')
print("Save Done!  time :", time.time() - start)  # 현재시각 - 시작시간 = 실행 시간

print('Job Done!')
print("time :", time.time() - start)  # 현재시각 - 시작시간 = 실행 시간

