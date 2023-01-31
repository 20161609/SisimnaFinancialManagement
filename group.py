import openpyxl
from excelMaking import*
from main import Groups
from main import Receipts
from main import Now

class group():
    def __init__(self, name, code, times, allowance):
        self.name = name
        self.code = code
        self.times = times
        self.allowance = allowance
        self.receipts = list()#number로 출력
        self.partition = [self]

    def Makefile_AccountBook(self):
        #0. Make File
        wb_accountBook = openpyxl.Workbook()
        ws_accountBook = wb_accountBook.active

        #1. Make Frame
        rows = len(self.receipts) + 2 + (Now.month + 1)
        cols = 7
        line_bordors(ws_accountBook, "B{}".format(2), "{}{}".format(Alphabets[cols], 1 + rows))

        ws_accountBook['B1'] = "회계장부 : {}".format(self.name)
        ws_accountBook['B2'] = "지출코드"
        ws_accountBook['C2'] = "날짜"
        ws_accountBook['D2'] = "내역"
        ws_accountBook['E2'] = "Input"
        ws_accountBook['F2'] = "Output"
        ws_accountBook['G2'] = "잔액"
        ws_accountBook['H2'] = "기타"
        ws_accountBook['B{}'.format(1 + rows)] = "계"

        #2. Scan Recipe
        balance = 0
        curRow = 2
        period = 12//self.times
        R = 0 # index of recipe
        for month in range(0, Now.month + 1)[::period]:
            curRow += 1
            ws_accountBook['B{}'.format(curRow)] = self.code

            if month == 0:
                balance += self.allowance
                ws_accountBook['C{}'.format(curRow)] = self.dateString(20221201)
                ws_accountBook['D{}'.format(curRow)] = "예산 입금"
                ws_accountBook['E{}'.format(curRow)] = self.allowance
                ws_accountBook['G{}'.format(curRow)] = balance
            else:
                balance += self.allowance
                ws_accountBook['C{}'.format(curRow)] = self.dateString(20230001 + month * 100)
                ws_accountBook['D{}'.format(curRow)] = "예산 입금"
                ws_accountBook['E{}'.format(curRow)] = self.allowance
                ws_accountBook['G{}'.format(curRow)] = balance

            curRow += 1
            while R < len(self.receipts):
                curReceipt = self.receipts[R]
                date = curReceipt.split()[0]
                if int(date) % 10000 > (month + period) * 100:
                    break
                pay = int(curReceipt.split()[2])
                content = curReceipt.split()[3]
                balance -= pay
                ws_accountBook['B{}'.format(curRow)] = self.code
                ws_accountBook['C{}'.format(curRow)] = self.dateString(date)
                ws_accountBook['D{}'.format(curRow)] = content
                ws_accountBook['F{}'.format(curRow)] = pay
                ws_accountBook['G{}'.format(curRow)] = balance
                R += 1
                curRow += 1

        wb_accountBook.save("accountBooks/회계장부({}).xlsx".format(self.name))

    def dateString(self, date):#20230130
        Year = str(date)[:4:]
        Month = str(date)[4:6:]
        Day = str(date)[6:8:]
        return Year + '/' + Month + '/' + Day