import win32com.client

excel = win32com.client.Dispatch("Excel.Application")
excel.Visible = True
workBook = excel.Workbooks.Open('C:\\Users\\Han\\Desktop\\test.xlsx')
workSheet = workBook.ActiveSheet
print(workSheet.Cells(1, 1).Value)
excel.Quit()