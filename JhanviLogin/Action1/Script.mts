'systemutil.Run "chrome.exe", "www.amazon.in"
''Datatable.import "‪D:\JhanvisLogin\Input\Login.xls"
'wait(5)
'DataTable.SetCurrentRow(2)
'
'Browser("Online Shopping site in").Page("Online Shopping site in").Link("Hello, sign in Account").Click
'Browser("Online Shopping site in").Page("Amazon Sign In").WebEdit("email").Set datatable.Value("User", 1)
'Browser("Online Shopping site in").Page("Amazon Sign In").WebButton("Continue").Click
'Browser("Online Shopping site in").Page("Amazon Sign In").WebEdit("password").Set datatable.Value("Pass", 1)
''Browser("Online Shopping site in").Page("Amazon Sign In").WebButton("Sign-In").Click

Set objExcel = CreateObject("Excel.Application")
    Set objWorkbook = objExcel.Workbooks.Open("D:\JhanvisLogin\Input\Login.xlsx")
    Set objSheet = objWorkbook.Worksheets("Sheet1")
    ColCount = objSheet.UsedRange.Columns.Count
    RowCount = objSheet.UsedRange.Rows.Count
     For i = 1 To RowCount 
        For j = 1 To ColCount
            fieldvalue = objSheet.Cells(i,j)
           
            MsgBox fieldvalue
        Next
   Next
    systemutil.Run "chrome.exe", "www.amazon.in"
    Set objSheet  = Nothing
   Set objWorkbook = Nothing
   Set objExcel = Nothing
