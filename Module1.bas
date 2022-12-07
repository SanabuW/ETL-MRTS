Attribute VB_Name = "Module1"
Sub csv_separator()
'Before running macro, ensure 'csv' folder is
'   created in the same directory as mrtssales92


Dim wbk As Workbook
Dim wks As Worksheet
Dim wks_name As String
Dim const_name As String
Dim filepath As String
'Dim ref_book As Workbook


Application.ScreenUpdating = False
Application.DisplayAlerts = False

'Find workbook
'Change later to 'find the only Excel sheet in the folder' for
'   transferrable code
Set wbk = ThisWorkbook
filepath = ThisWorkbook.path


'Find/replace all "May " with "May. "
Cells.Replace What:="May ", Replacement:="May. ", LookAt:=xlPart, _
    SearchOrder:=xlByRows, MatchCase:=True, SearchFormat:=False, _
    ReplaceFormat:=False

'Create construction sheet
Worksheets.Add
ActiveSheet.Name = "const"

For Each wks In wbk.Worksheets
    'Check that this sheet isn't the construction sheet
    If wks.Name <> "const" Then
        'Main loop
        'Get sheet name
        wks_name = wks.Name
        
        'Assemble data on construction_sheet
        Sheets(wks_name).Activate
            
        'Copy top left col names
        Range("A4:B4").Select
        Selection.Copy
        Sheets("const").Activate
        Range("A1").Select
        ActiveSheet.Paste
        Application.CutCopyMode = False
        
        'copy month headers
        'C5:N5
        Sheets(wks_name).Activate
        Range("C5:N5").Select
        Selection.Copy
        Sheets("const").Activate
        Range("C1").Select
        ActiveSheet.Paste
        Application.CutCopyMode = False
        
        
        'copy data
        'A7:O71
        Sheets(wks_name).Activate
        Range("A7:N71").Select
        Selection.Copy
        Sheets("const").Activate
        Range("A2").Select
        ActiveSheet.Paste
        Application.CutCopyMode = False
        
        'Export construction sheet as csv
        ActiveSheet.Copy
        ActiveWorkbook.SaveAs Filename:=filepath & "\csv\mrtssales92_" & wks_name & ".csv", FileFormat:=xlCSV
        ActiveWorkbook.Close
    End If
Next wks

'Clean up construction sheet
Worksheets("const").Delete

Application.ScreenUpdating = True
Application.DisplayAlerts = True

End Sub


