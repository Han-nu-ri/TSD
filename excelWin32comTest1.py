import win32com.client

excel = win32com.client.Dispatch("Excel.Application")
excel.Visible = True
workBook = excel.Workbooks.Add()
workSheet = workBook.Worksheets("Sheet1")
workSheet.Cells(1, 1).Value = "Hello TSD"
workBook.SaveAs('C:\\Users\\Han\\Desktop\\test.xlsx')
excel.Quit()
