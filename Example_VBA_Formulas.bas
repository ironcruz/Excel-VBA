Attribute VB_Name = "Example_VBA_Formulas"

Sub Macro_ExampleConfig()
'
' Layer fields configuation for site.xml, HTML feature description, and mobile JSON field configs
'

'
    On Error Resume Next
    Range("A:A").Select
    Selection.SpecialCells(xlCellTypeBlanks).EntireRow.Delete
    Dim lastRow As Long
    lastRow = Range("A" & Rows.Count).End(xlUp).Row
    
    Columns("A:A").Select
    Selection.Replace What:="Shape", Replacement:="SHAPE", LookAt:=xlPart, _
    SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
    ReplaceFormat:=False
    Columns("A:A").Select
    Selection.Replace What:="Shape_Length", Replacement:="SHAPE_LENGTH", LookAt:=xlPart, _
    SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
    ReplaceFormat:=False
    Columns("A:A").Select
    Selection.Replace What:="Shape_Area", Replacement:="SHAPE_AREA", LookAt:=xlPart, _
    SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
    ReplaceFormat:=False

    Range("M1").Select
    ActiveCell.FormulaR1C1 = "Searchable"
    Range("M2").Select
    ActiveCell.FormulaR1C1 = "=IF(RC[-7]<>"""",""true"",""false"")"
    lastRow = Range("A" & Rows.Count).End(xlUp).Row
    Range("M2").AutoFill Destination:=Range("M2:M" & lastRow)
    Range("N1").Select
    ActiveCell.FormulaR1C1 = "Visible"
    Range("N2").Select
    ActiveCell.FormulaR1C1 = _
        "=IF(ISNUMBER(SEARCH(""*hid*"",RC[-10])),""false"",""true"")"
        lastRow = Range("A" & Rows.Count).End(xlUp).Row
        Range("N2").AutoFill Destination:=Range("N2:N" & lastRow)
    
    Range("O1").Select
    ActiveCell.FormulaR1C1 = "<Field CanSymbolizeClassBreaks=""true"" CanSymbolizeUniqueValue=""true"""
    Range("P1").Select
    ActiveCell.FormulaR1C1 = "DisplayName="""
    Range("Q1").Select
    ActiveCell.FormulaR1C1 = "FocusField=""false"" Name="""
    
    Range("O2").Select
    ActiveCell.FormulaR1C1 = _
        "=IF(RC[-3]<>"""",R1C12&"" ""&R1C13&RC9&""""""""&"" ""&R1C14&RC[-11]&""""""""&"" ""&""Searchable=""""""&RC[-2]&""""""""&"" ""&""Visible=""""""&RC[-1]&""""""""&"" ""&""/>"",IF((AND(RC13=""false"",RC14=""true"")),"""",R1C15&"" ""&R1C17&RC[-14]&""""""""&"" ""&""Searchable=""""""&RC[-2]&""""""""&"" ""&""Visible=""""""&RC[-1]&""""""""&"" ""&""/>""))"
        Range("O2").AutoFill Destination:=Range("O2:O" & lastRow)
    
    Range("P2").Select
    ActiveCell.FormulaR1C1 = "=IF(RC[-1]<>"""",RC[-1])"
    Range("P2").AutoFill Destination:=Range("P2:P" & lastRow)
    Columns("P:P").Select
    Application.CutCopyMode = False
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
        
    For Each C In [P1:P200]
    
    If UCase(C.Value) = "FALSE" Or C.Value = "False" Then
        C.Delete
    End If
    Next

    Columns("O:O").Select
    Selection.Delete Shift:=xlToLeft
    Range("O1").Select

    ActiveCell.FormulaR1C1 = _
        "<Fields>"
        
    Range("O:O").Select
    
    lastRow = Range("O" & Rows.Count).End(xlUp).Row + 1
        
    ActiveSheet.Cells(lastRow, "O").Value = "</Fields>"
    lastRow = Range("O" & Rows.Count).End(xlUp).Row + 1
    ActiveSheet.Cells(lastRow, "O").Value = "</Layer>"
    
    
    Range("Q1").Select
    ActiveCell.FormulaR1C1 = "FeatureDescription"
    Range("Q2").Select
    ActiveCell.FormulaR1C1 = "=INDEX(C[-15],MATCH(RC[-11],C[-11],0))&"" ""&""-"""
    lastRow = Range("A" & Rows.Count).End(xlUp).Row
    Range("Q2").AutoFill Destination:=Range("Q2:Q" & lastRow)
    Columns("Q:Q").Select
    Selection.Replace What:="#N?A", Replacement:="", LookAt:=xlPart, _
    SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
    ReplaceFormat:=False
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
    :=False, Transpose:=False
    Columns("Q:Q").Select
    Selection.Replace What:="#N/A", Replacement:="", LookAt:=xlPart, _
    SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
    ReplaceFormat:=False
    
    For Each C In [q1:q200]
    
    If UCase(C.Value) = "" Or C.Value = " " Then
        C.Delete
    End If
    Next
        
    Range("R1").Select
    ActiveCell.FormulaR1C1 = "FeatureDescription2"
    Range("R2").Select
    ActiveCell.FormulaR1C1 = "=""{""&INDEX(C[-17],MATCH(RC[-12],C[-12],0))&""}"""
    lastRow = Range("A" & Rows.Count).End(xlUp).Row
    Range("R2").AutoFill Destination:=Range("R2:R" & lastRow)
    Columns("R:R").Select
    Selection.Replace What:="#N?A", Replacement:="", LookAt:=xlPart, _
    SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
    ReplaceFormat:=False
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
    :=False, Transpose:=False
    Columns("R:R").Select
    Selection.Replace What:="#N/A", Replacement:="", LookAt:=xlPart, _
    SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
    ReplaceFormat:=False
    
    For Each C In [R1:R200]
    
    If UCase(C.Value) = "" Or C.Value = " " Then
        C.Delete
    End If
    Next
    
    Range("s2").Select
    ActiveCell.FormulaR1C1 = _
        "FeatureDescription=""&lt;strong style=&quot;font-size: 13.3333px;&quot;&gt;"
    Range("t2").Select
    ActiveCell.FormulaR1C1 = _
        "&lt;/strong&gt;&lt;span style=&quot;font-size: 13.3333px;&quot;&gt;&amp;nbsp;"
    Range("s3").Select
    ActiveCell.FormulaR1C1 = _
        "&lt;/span&gt;&lt;div&gt;&lt;span style=&quot;font-size: 13.3333px;&quot;&gt;&lt;br/&gt;&lt;/span&gt;&lt;/div&gt;&lt;div&gt;&lt;strong style=&quot;font-size: 13.3333px;&quot;&gt;"
    Range("t3").Select
    ActiveCell.FormulaR1C1 = _
        "&lt;/strong&gt;&lt;span style=&quot;font-size: 13.3333px;&quot;&gt;&amp;nbsp;"
    Range("s4").Select
    ActiveCell.FormulaR1C1 = _
        "&lt;/span&gt;&lt;/div&gt;&lt;div&gt;&lt;span style=&quot;font-size: 13.3333px;&quot;&gt;&lt;br/&gt;&lt;/span&gt;&lt;/div&gt;&lt;div&gt;&lt;strong style=&quot;font-size: 13.3333px;&quot;&gt;"
    lastRow = Range("r" & Rows.Count).End(xlUp).Row
    Range("s4").AutoFill Destination:=Range("s4:s" & lastRow)
    Range("t4").Select
    ActiveCell.FormulaR1C1 = _
        "&lt;/strong&gt;&lt;span style=&quot;font-size: 13.3333px;&quot;&gt;&amp;nbsp;"
    lastRow = Range("R" & Rows.Count).End(xlUp).Row
    Range("T4").AutoFill Destination:=Range("T4:T" & lastRow)
   
    Range("U2").Select
    Application.CutCopyMode = False
    ActiveCell.FormulaR1C1 = "=RC[-2]&RC[-4]&RC[-1]&RC[-3]"
    lastRow = Range("Q" & Rows.Count).End(xlUp).Row
    Range("U2").AutoFill Destination:=Range("U2:U" & lastRow)
    
    lastRow = Range("U" & Rows.Count).End(xlUp).Row + 1
    ActiveSheet.Cells(lastRow, "U").Value = "&lt;/span&gt;&lt;/div&gt;"""

    Columns("U:U").Select
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
    :=False, Transpose:=False
    
    ' could delete the line below if the feature description isn't correct
    Columns("P:T").Select
    Selection.Delete Shift:=xlToLeft
    
    Range("P1").Select
    ActiveCell.FormulaR1C1 = "Copy Feature Description Below"
    
    lastRow = Range("A" & Rows.Count).End(xlUp).Row
    Range("S1").Select
    ActiveCell.FormulaR1C1 = "M_Visible"
    
    Range("S2").Select
    ActiveCell.FormulaR1C1 = _
        "=IF(COUNTIF(RC[-10],""*hid*""),""false"",""true"")"
    Range("S2").AutoFill Destination:=Range("S2:S" & lastRow)
    
    Range("T1").Select
    ActiveCell.FormulaR1C1 = "M_Editable"
    
    Range("T2").Select
    ActiveCell.FormulaR1C1 = _
        "=IF(RC[-19]=""CREATED_DATE"",""false"",IF(RC[-19]=""CREATED_USER"",""false"",IF(RC[-19]=""GLOBALID"",""false"",IF(RC[-19]=""LAST_EDITED_DATE"",""false"",IF(RC[-19]=""LAST_EDITED_USER"",""false"",IF(RC[-19]=""OBJECTID"",""false"",IF(RC[-19]=""UNIQUE_ID"",""false"",IF(RC[-19]=""NETWORK_ROUTE_ID"",""false"",IF(RC[-19]=""SENT_TO_FLOWCAL"",""false"",IF(RC[-19]=""TECHNICI" & _
        "AN_REQ_TO_DELETE_LF"",""false"",IF(RC[-19]=""MXASSETNUM"",""false"",IF(RC[-19]=""MXSITEID"",""false"",IF(RC[-19]=""GGS"",""false"",IF(RC[-19]=""SHAPE"",""false"",IF(RC[-19]=""SHAPE.LEN"",""false"",IF(RC[-19]=""SHAPE.AREA"",""false"",IF(RC[-19]=""DIVISION"",""false"",IF(RC[-19]=""ENCROACHMENT_ID_LEGACY"",""false"",IF(RC[-19]=""FRANCHISE"",""false"",IF(RC[-19]=""GGS_D" & _
        "ISTRICT_ID"",""false"",IF(RC[-19]=""INSTALLED_DATE"",""false"",IF(RC[-19]=""OBSERVATION_ID"",""false"",IF(RC[-19]=""PATROL_INTERVAL"",""false"",IF(RC[-19]=""PATROL_VENDOR"",""false"",IF(RC[-19]=""PATROL_YEAR"",""false"",IF(RC[-19]=""PRESERVE_RELATE_ID"",""false"",IF(RC[-19]=""PRESERVE_RELATE_INSTANCE"",""false"",IF(RC[-19]=""TO_DATE"",""false"",IF(RC[-19]=""FRANCHIS" & _
        "E_ID"",""false"",IF(RC[-19]=""IS_SAFETY_CRITICAL_EQUIPMENT"",""false"",IF(RC[-19]=""RATING"",""false"",IF(RC[-19]=""PIPELINE_ID"",""false"",""true""))))))))))))))))))))))))))))))))" & _
        ""
    Range("T2").AutoFill Destination:=Range("T2:T" & lastRow)
    
    Range("U1").Select
    ActiveCell.FormulaR1C1 = "M_Date_Format"
    
    Range("U2").Select
    ActiveCell.FormulaR1C1 = _
        "=IF(COUNTIF(RC[-18],""*date*""),""""""format""""""&"": {""&"" ""&""""""dateFormat""""""&"" :  ""&""""""shortDateShortTime""""""&"",""&""""""timezone""""""&"" : "" &""""""utc""""""&""}"","""")"
    Range("U2").AutoFill Destination:=Range("U2:U" & lastRow)
    
    Range("V1").Select
    ActiveCell.FormulaR1C1 = "M_Number_Format"
    
    Range("V2").Select
    ActiveCell.Formula2R1C1 = _
        "=IF(SUM(COUNTIF(RC[-19],{""*short*"",""*double*"",""*long*""})),""""""format""""""&"": {""&"" ""&""""""places""""""&"" :  ""&""2""&"",""&""""""digitSeparator""""""&"" : "" &""true""&""}"","""")"
    Range("V2").AutoFill Destination:=Range("V2:V" & lastRow)
    
    Range("W1").Select
    ActiveCell.FormulaR1C1 = "Field_Configs"
    
    Range("W2").Select
    ActiveCell.FormulaR1C1 = _
        "=IF(AND(RC[-1]="""",RC[-2]=""""),""""""fieldInfos""""""&"": [ ""&""{""&"" ""&""""""fieldName""""""&"": ""&""""""""&RC[-22]&""""""""&"",""&""""""visible""""""&"": "" &RC[-4]&"",""&""""""isEditable""""""&"": "" &RC[-3]&"",""&""""""label""""""&"": ""&""""""""&RC[-21]&""""""""&""}""&"","",IF(RC[-1]="""",""{""&"" ""&""""""fieldName""""""&"": ""&""""""""&RC[-22]&""""""""&" & _
        """,""&""""""visible""""""&"": "" &RC[-4]&"",""&""""""isEditable""""""&"": "" &RC[-3]&"",""&""""""label""""""&"": ""&""""""""&RC[-21]&""""""""&"",""& RC[-2]&""}""&"","",IF(RC[-2]="""",""{""&"" ""&""""""fieldName""""""&"": ""&""""""""&RC[-22]&""""""""&"",""&""""""visible""""""&"": "" &RC[-4]&"",""&""""""isEditable""""""&"": "" &RC[-3]&"",""&""""""label""""""&"": ""&""" & _
        """""""&RC[-21]&""""""""&"",""&RC[-1]&""}""&"","")))" & _
        ""
    Range("W3").Select
    ActiveCell.FormulaR1C1 = _
        "=IF(AND(RC[-1]="""",RC[-2]=""""),""{""&"" ""&""""""fieldName""""""&"": ""&""""""""&RC[-22]&""""""""&"",""&""""""visible""""""&"": "" &RC[-4]&"",""&""""""isEditable""""""&"": "" &RC[-3]&"",""&""""""label""""""&"": ""&""""""""&RC[-21]&""""""""&""}""&"","",IF(RC[-1]="""",""{""&"" ""&""""""fieldName""""""&"": ""&""""""""&RC[-22]&""""""""&"",""&""""""visible""""""&"": """ & _
        " &RC[-4]&"",""&""""""isEditable""""""&"": "" &RC[-3]&"",""&""""""label""""""&"": ""&""""""""&RC[-21]&""""""""&"",""& RC[-2]&""}""&"","",IF(RC[-2]="""",""{""&"" ""&""""""fieldName""""""&"": ""&""""""""&RC[-22]&""""""""&"",""&""""""visible""""""&"": "" &RC[-4]&"",""&""""""isEditable""""""&"": "" &RC[-3]&"",""&""""""label""""""&"": ""&""""""""&RC[-21]&""""""""&"",""&RC" & _
        "[-1]&""}""&"","")))" & _
        ""
    lastRow = Range("A" & Rows.Count).End(xlUp).Row - 1
    Range("W3").AutoFill Destination:=Range("W3:W" & lastRow)
    lastRow = Range("W" & Rows.Count).End(xlUp).Row + 1
    ActiveSheet.Cells(lastRow, "W").Select
    
    ActiveCell.FormulaR1C1 = _
        "=IF(AND(RC[-1]="""",RC[-2]=""""),""{""&"" ""&""""""fieldName""""""&"": ""&""""""""&RC[-22]&""""""""&"",""&""""""visible""""""&"": "" &RC[-4]&"",""&""""""isEditable""""""&"": "" &RC[-3]&"",""&""""""label""""""&"": ""&""""""""&RC[-21]&""""""""&""}],"",IF(RC[-1]="""",""{""&"" ""&""""""fieldName""""""&"": ""&""""""""&RC[-22]&""""""""&"",""&""""""visible""""""&"": "" &RC" & _
        "[-4]&"",""&""""""isEditable""""""&"": "" &RC[-3]&"",""&""""""label""""""&"": ""&""""""""&RC[-21]&""""""""&"",""& RC[-2]&""}]""&"","",IF(RC[-2]="""",""{""&"" ""&""""""fieldName""""""&"": ""&""""""""&RC[-22]&""""""""&"",""&""""""visible""""""&"": "" &RC[-4]&"",""&""""""isEditable""""""&"": "" &RC[-3]&"",""&""""""label""""""&"": ""&""""""""&RC[-21]&""""""""&"",""&RC[-1]" & _
        "&""}],"")))" & _
        ""

    
    End Sub


