# 엑셀 확장자 변환작업(xls to xlsx)
import win32com.client as win32
import os

def xls_to_xlsx(file_path):
    excel: win32 = win32.Dispatch('Excel.Application')
    wb: excel = excel.Workbooks.Open(file_path) # xls 파일 열기

    # 그냥 문자열을 가져와서 저장하려고 하면 당연히 에러가 난다. 슬래시(또는 역슬래시)로 표현되는 문자는 os.sep로 치환하도록 한다.
    wb.SaveAs(file_path.replace('/', os.sep)+"x", FileFormat = 51) #FileFormat = 51 is for .xlsx extension
    wb.Close() #FileFormat = 56 is for .xls extension
    excel.Application.Quit()