import openpyxl
import requests
import os
from PyPDF2 import PdfFileMerger
from pathlib import Path

def excel_get():
    opn = openpyxl.open("21718972__2022-01-01__2022-01-31.xlsx", read_only=True)
    sheet = opn.active
    #print(sheet[2][15].value)
    #print(sheet.max_row)

    #создаем пустой список для хранения ссылок
    links = []

    for i in range (0, sheet.max_row):
        try:
            s = str(sheet[i][15].value)
            s = s.split("\"")
            links.append(s[1])
        except Exception:
            continue

    return links

def download_pdf(links):
    #print(links)
    #https://lk.platformaofd.ru/web/noauth/cheque/pdf/z-report-66990954347-1642921504000-4131373248000.pdf
    os.mkdir('pdfs')
    os.chdir("pdfs")
    for link in links:
        link = str(link)
        link_id = link.split('=')[1].split('&')[0]
        link_date = link.split('=')[2].split('&')[0]
        link_fp = link.split('=')[3].split('&')[0]

        url = f"https://lk.platformaofd.ru/web/noauth/cheque/pdf/z-report-{link_id}-{link_date}-{link_fp}.pdf"
        responce = requests.get(url)
        with open(f'{link_id}.pdf', 'wb') as file:
            file.write(responce.content)

def make_big_pdf():
    #pdf_merger = PdfFileMerger()
    reports_dir = (Path.home()/"pdfs")
    for path in reports_dir.glob("*.pdf"):
        print(path.name)    

def main():
    #download_pdf(links=excel_get())
    make_big_pdf()

if __name__ == '__main__':
    main()

