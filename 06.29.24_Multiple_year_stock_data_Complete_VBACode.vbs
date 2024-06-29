Dim excelApp, workbook

' Create an instance of Excel
Set excelApp = CreateObject("Excel.Application")

' Make Excel visible (optional)
excelApp.Visible = True

' Open the workbook
Set workbook = excelApp.Workbooks.Open("c:\users\isfun\desktop\06.29.24_Multiple_year_stock_data_complete")

' Run the macro
excelApp.Run "QuarterlyStockAnalysisAllSheets"

' Save and close the workbook
workbook.Save
workbook.Close

' Quit Excel
excelApp.Quit

' Clean up
Set workbook = Nothing
Set excelApp = Nothing
