import openpyxl
import re

def excel_get():
    opn = openpyxl.open("21718972__2022-01-01__2022-01-31.xlsx", read_only=True)
    sheet = opn.active
    print(sheet[2][15].value)
    print(sheet.max_row)
    for i in range (3, sheet.max_row):
        s = re.findall(r'https.+,', sheet[i][15].value)
        print(s)

def main():
    excel_get()

if __name__ == '__main__':
    main()

