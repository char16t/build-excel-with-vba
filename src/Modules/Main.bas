Sub CreateWorksheet(Name As String)
    Dim wb As Workbook: Set wb = ActiveWorkbook
    Dim shts As Sheets: Set shts = wb.Sheets
    Dim obj As Object
    Set obj = shts.Add(After:=ActiveSheet, Count:=1, Type:=XlSheetType.xlWorksheet)
    obj.Name = Name
End Sub

Sub DeleteWorksheetByIndex(Index As Integer)
    Dim wb As Workbook: Set wb = ActiveWorkbook
    Dim shts As Sheets: Set shts = wb.Sheets
    shts(Index).Delete
End Sub

Public Sub Main()
    CreateWorksheet ("Overview")
    CreateWorksheet ("Data")
    DeleteWorksheetByIndex (1)
End Sub
