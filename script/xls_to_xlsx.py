# xls to xlsx
import win32com.client as win32
fname = "D:\ysk\Python\python_auto_work\download_file\마음손거래내역_20230712045504테스트용.xls"
excel = win32.gencache.EnsureDispatch('Excel.Application')
wb = excel.Workbooks.Open(fname)

wb.SaveAs(fname+"x", FileFormat = 51) #FileFormat = 51 is for .xlsx extension
wb.Close() #FileFormat = 56 is for .xls extension
excel.Application.Quit()