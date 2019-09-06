import openpyxl

with open('E:/Dreamweaver/workspace/input.txt','r') as f1:
    keyword = f1.readline().strip('\n').split()

wb = openpyxl.load_workbook("E:/Dreamweaver/workspace/FGSG gene database.xlsx")
ws = wb["Sheet1"]
number = []
with open('E:/Dreamweaver/workspace/output.txt','w') as f2:
    for i in keyword:
        num = 0
        for row in ws.rows:
            for cell in row:
                if(cell.value == i):
                    f2.write(str(cell.value) + ' module:' + str(cell.row) + '\n')
                    number.append(int(cell.row))
                    num = 1
                    break
            if(num == 1):
                break
    number.sort()
    max, sum, temp, final = 0, 0, 0, 0
    for i in range(len(number)):
        if(i == 0):
            temp = number[i]
            sum += 1
        if(i > 0):
            if(temp == number[i]):
                sum += 1
            else:
                temp = number[i]
                if(max < sum):
                    final = number[i - 1]
                    max = sum
                    sum = 1
    f2.write('The module is :'+str(final))

