if WScript.Arguments.Count < 2 Then 
    WScript.Echo "Error! Please specify the source path and the destination. Usage: XlsToCsv SourcePath.xls Destination.csv" 
    Wscript.Quit 
End If 

Dim oExcel 
Set oExcel = CreateObject("Excel.Application") 
oExcel.DisplayAlerts = FALSE 'to avoid prompts
Dim oBook, local
Set oBook = oExcel.Workbooks.Open(Wscript.Arguments.Item(0))
local = true 
'call oBook.SaveAs(WScript.Arguments.Item(1), 6, 0, 0, 0, 0, 0, 0, 0, 0, 0, local) 'this changed
call oBook.SaveAs(WScript.Arguments.Item(1), 6) 'this changed
oBook.Close False 
oExcel.Quit 
WScript.Echo "Done" 