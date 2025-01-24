On Error Resume Next

dimension objExcel, objWorkbook

Set objExcel = GetObject(, "Excel.Application")

if Err.Number <> 0 then
  ' Handle the error
  MsgBox "Could not find Excel application. Please ensure Excel is installed and running.", vbCritical
  Err.Clear
  WScript.Quit
end if

Set objWorkbook = objExcel.Workbooks.Open("C:\path\to\your\excel.xlsx")

' ... rest of your code ...

'Clean up
on error goto 0
Set objWorkbook = Nothing
Set objExcel = Nothing