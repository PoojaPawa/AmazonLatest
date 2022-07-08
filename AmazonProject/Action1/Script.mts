SystemUtil.CloseProcessByName"Chrome.exe"
'SystemUtil.Run"Chrome.exe","www.amazon.in"

'TC_06()
FilePath= "C:\Users\user240\Documents\Amazon\Test_Data\Test Data.xlsx"
ExcelSheet="Test Data"
SheetName="Amazon Data"
DataTable.AddSheet ExcelSheet
DataTable.ImportSheet FilePath,SheetName,ExcelSheet
rowCount = DataTable.GetSheet(ExcelSheet).GetRowCount

For i= 1 To rowCount
DataTable.SetCurrentRow (i)
If DataTable.Value("Expected_Flag",ExcelSheet)="Y" Then
  SystemUtil.Run"Chrome.exe",DataTable.Value("URL",ExcelSheet)
'  SystemUtil.Run"Chrome.exe","www.amazon.in"
  ExecuteTestCase(DataTable.Value("testCaseID",ExcelSheet))
  SystemUtil.CloseProcessByName"Chrome.exe"
  DataTable.Value("Result",ExcelSheet)=Environment.Value("Result")
End If
Next
DataTable.ExportSheet FilePath,ExcelSheet,SheetName


'SignIn()

'SignOut() @@ script infofile_;_ZIP::ssf98.xml_;_

'TC_01()

'TC_02()

'TC_03()

'TC_04()

'TC_05()

'TC_06()
 @@ script infofile_;_ZIP::ssf171.xml_;_
' TC_07()

' TC_08()
 
' TC_09()
 
  'TC_10()



