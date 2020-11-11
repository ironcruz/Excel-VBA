Attribute VB_Name = "SaveAsPDF"
Sub SaveAsPDF()
    
    Dim wb As Workbook
    Set wb = ThisWorkbook
    Dim saveLocation As String
    'BELOW COMMENTS FOR MULTIPLE SHEETS IN WORKBOOK
    'Dim WS_Count As Integer
    'Dim I As Integer
    'BELOW FOR RANGE INSTEAD OF SHEET
    'Dim rng As Range
    
    
    saveLocation = "C:\DEV\Excel\" + ActiveWorkbook.ActiveSheet.Name + ".pdf"
    ActiveSheet.ExportAsFixedFormat Type:=xlTypePDF, _
    Filename:=saveLocation
    
    'LOOP
    'WS_Count = ActiveWorkbook.Worksheets.Count
    'For I = 1 To WS_Count
        'Worksheets(ActiveWorkbook.Worksheets(I).Name).Activate
        'saveLocation = "C:\DEV\Excel\" + ActiveWorkbook.Worksheets(I).Name + ".pdf"
        'ActiveSheet.ExportAsFixedFormat Type:=xlTypePDF, _
        'Filename:=saveLocation
    
    'Next I
    
    'Save a range as PDF
    'Set Rng = Sheets(ActiveWorkbook.ActiveSheet.Name).Range("RANGE HERE")
    'saveLocation = "DESIRED PATH" + rng NAME(IF APPLICABLE) + ".pdf"
    'rng.ExportAsFixedFormat Type:=xlTypePDF, _
    'Filename:=saveLocation
End Sub
