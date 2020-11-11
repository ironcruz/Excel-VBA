Attribute VB_Name = "Extraction"
Sub WorksheetLoop()
    
    Dim wb As Workbook
    Set wb = ThisWorkbook
    Dim WS_Count As Integer
    Dim I As Integer
    Dim strPattern As String: strPattern = "[^a-Z0-9-]"
    
    
    WS_Count = ActiveWorkbook.Worksheets.Count
    lastRow = Range("F" & Rows.Count).End(xlUp).Row
    
    For I = 1 To WS_Count
        orgName = ActiveWorkbook.Worksheets(I).Name
        'MsgBox ActiveWorkbook.Worksheets(I).Name
        If orgName Like "*.*" Then ActiveSheet.Name = Replace(orgName, ".", "_")
        orgName = Replace(orgName, ".", "_")
        
        
        Workbooks.Add.SaveAs "C:\PLC_3.0\Final_Reqs_15Sep20\" & orgName
        ActiveSheet.Name = orgName
        ThisWorkbook.Worksheets(orgName).Range("F:V").Copy
        ActiveWorkbook.Worksheets(orgName).Range("A1").PasteSpecial (xlPasteValues)
        Rows(1).EntireRow.Delete
        Rows(1).EntireRow.Delete
        Rows(1).EntireRow.Delete
        Rows(1).EntireRow.Delete
        Range("C:C").Copy Range("ZZ:ZZ")
        Range("D:H").Delete Shift:=xlToLeft
        
        ActiveWorkbook.Close SaveChanges:=True
      
                  
    Next I

End Sub

