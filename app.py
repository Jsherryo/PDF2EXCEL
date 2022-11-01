import camelot.io as camelot
import pandas as pd
import openpyxl as xl
import PySimpleGUI as sg

sg.theme('Dark Blue 3')
layout = [
    [sg.Text('PDF表格轉EXCEL', font=("Helvetica", 25), justification='center', size=(100, 2), text_color='white')],
    [sg.Text("請選擇PDF檔: ", font=("Helvetica", 15), text_color='white'), sg.FileBrowse("選擇")],
    [sg.Text("", size=(45, 2))],
    [sg.Button('轉換', font=("Helvetica", 15), size=(10, 1)), sg.Button('關閉', font=("Helvetica", 15), size=(10, 1))]
]


def wdop():
    window = sg.Window('PDF TO EXCEL', layout, size=(300, 300), element_justification='c')
    while True:
        event, values = window.read()
        if event in (None, '關閉'):
            break
        else:
            try:
                if "pdf" in values['選擇']:
                    sg.PopupAnimated(sg.DEFAULT_BASE64_LOADING_GIF, background_color='white', time_between_frames=100)
                    pdf2ecl(values['選擇'])
                    sg.PopupAnimated(None)
                    sg.Popup(f"EXCEL檔已輸出至\n{values['選擇'].replace('pdf','xlsx')}")
            except Exception as e:
                sg.Popup(e)
                continue
    window.close()


def pdf2ecl(url):
    tables = camelot.read_pdf(url, pages='all', flavor='stream')
    wb = xl.Workbook()
    sheets = wb.worksheets
    for i in sheets:
        wb.remove(i)
    wb.create_sheet("Sheet")
    dt = []
    for i in range(len(tables)):
        tbb = list(tables[i].df.values)
        for tb_temp in tbb:
            tb_temp = list(tb_temp)
            if tb_temp not in dt and tb_temp[0] != "":
                dt.append(tb_temp)
    for tb_t in dt:
        try:
            if 0 < int(tb_t[0].replace("*", "")) < 10000:
                wb["Sheet"].append(tb_t)
        except:
            pass
    wb.save(f'{url.replace("pdf","xlsx")}')


wdop()

