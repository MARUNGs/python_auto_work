# 엑셀 확장자 변환작업(xls to xlsx)
import win32com.client as win32

def xlsToXlsx(self):
    fname = self.file_path.toPlainText()
    excel = win32.Dispatch('Excel.Application')
    wb = excel.Workbooks.Open(fname)

    wb.SaveAs(fname+"x", FileFormat = 51) #FileFormat = 51 is for .xlsx extension
    wb.Close() #FileFormat = 56 is for .xls extension
    excel.Application.Quit()