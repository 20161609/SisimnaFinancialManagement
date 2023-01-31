from group import*
from os import listdir
from openpyxl import load_workbook
from datetime import datetime

Groups = {}
Receipts = []

Now = datetime.now().now()
Year = Now.year
Month = Now.month
Day = Now.day

if __name__ == '__main__':
    # 엑셀파일 불러오기
    load_wb = load_workbook("test.xlsx", read_only=True)
    load_ws = load_wb['총계']

    #그룹 정보 생성
    code = ''
    row = 'A';    col = '1'
    curCode = code

    while(True):
        code = str(load_ws['A'+col].value) + str(load_ws['B'+col].value)
        if(code == 'AD'): break
        if(code[0] == 'A' and load_ws['C'+col].value != '00'):
            newGroup = group(
                name = load_ws['D'+ col].value,
                code = code + str(load_ws['C'+col].value),
                times = load_ws['G'+col].value,
                allowance = load_ws['F'+col].value
            )
            Groups[newGroup.name] = newGroup

        col = str(int(col)+1)

    #영수증 내역 저장 및 각 그룹에 추가
    folder_list = listdir("Receipts")
    TotalCost = 0
    a = []
    for fileName in folder_list:
        print(fileName)
        groupName = fileName.split()[1]
        cost = int(fileName.split('.')[0].split()[2])
        TotalCost += cost
        if(groupName in Groups == False):
            print("Error", fileName)
            continue

        Groups[groupName].receipts.append(fileName)

    for g in Groups:
        Groups[g].Makefile_AccountBook()

    print(TotalCost)