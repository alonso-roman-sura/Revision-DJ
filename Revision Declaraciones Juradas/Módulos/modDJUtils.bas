Option Explicit

Public Function PickSourceFile(ByVal dialogTitle As String) As String
    Dim v As Variant
    
    v = Application.GetOpenFilename( _
        "Archivos Excel y HTML (*.xlsx;*.xls;*.xlsm;*.htm;*.html),*.xlsx;*.xls;*.xlsm;*.htm;*.html", _
        , dialogTitle)
    
    If VarType(v) = vbBoolean Then
        PickSourceFile = vbNullString
    Else
        PickSourceFile = CStr(v)
    End If
End Function

Public Function GetRequiredTable(ByVal tableName As String) As ListObject
    Dim ws As Worksheet
    Dim lo As ListObject
    
    For Each ws In ThisWorkbook.Worksheets
        For Each lo In ws.ListObjects
            If StrComp(lo.Name, tableName, vbTextCompare) = 0 Then
                Set GetRequiredTable = lo
                Exit Function
            End If
        Next lo
    Next ws
    
    Err.Raise vbObjectError + 1600, , "No se encontró la tabla '" & tableName & "'."
End Function

Public Function GetListColumn(ByVal lo As ListObject, ByVal headerName As String) As ListColumn
    Dim lc As ListColumn
    
    For Each lc In lo.ListColumns
        If CanonicalHeader(lc.Name) = CanonicalHeader(headerName) Then
            Set GetListColumn = lc
            Exit Function
        End If
    Next lc
End Function

Public Sub RemoveTableColumnIfExists(ByVal lo As ListObject, ByVal headerName As String)
    Dim lc As ListColumn
    
    Set lc = GetListColumn(lo, headerName)
    If Not lc Is Nothing Then lc.Delete
End Sub

Public Sub RemoveColumnsByPrefix(ByVal lo As ListObject, ByVal prefixText As String)
    Dim i As Long
    Dim prefixKey As String
    Dim currentKey As String
    
    prefixKey = CanonicalHeader(prefixText)
    
    For i = lo.ListColumns.Count To 1 Step -1
        currentKey = CanonicalHeader(lo.ListColumns(i).Name)
        If Left$(currentKey, Len(prefixKey)) = prefixKey Then
            lo.ListColumns(i).Delete
        End If
    Next i
End Sub

Public Sub DeleteSheetIfExists(ByVal sheetName As String)
    Dim ws As Worksheet
    
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(sheetName)
    On Error GoTo 0
    
    If Not ws Is Nothing Then
        ws.Delete
    End If
End Sub

Public Sub ForceTableColumnsAsText(ByVal lo As ListObject, ByVal headerList As Variant)
    Dim i As Long
    
    For i = LBound(headerList) To UBound(headerList)
        ForceTableColumnAsText lo, CStr(headerList(i))
    Next i
End Sub

Public Sub ForceTableColumnAsText(ByVal lo As ListObject, ByVal headerName As String)
    Dim lc As ListColumn
    Dim c As Range
    Dim v As Variant
    
    Set lc = GetListColumn(lo, headerName)
    If lc Is Nothing Then Exit Sub
    If lc.DataBodyRange Is Nothing Then Exit Sub
    
    lc.DataBodyRange.NumberFormat = "@"
    
    For Each c In lc.DataBodyRange.Cells
        v = c.Value2
        If Len(Trim$(CStr(v))) > 0 Then
            c.Value = CStr(v)
        Else
            c.Value = vbNullString
        End If
    Next c
End Sub

Public Sub CoerceTableDateColumns(ByVal lo As ListObject, ByVal headerList As Variant)
    Dim i As Long
    
    For i = LBound(headerList) To UBound(headerList)
        CoerceTableDateColumn lo, CStr(headerList(i))
    Next i
End Sub

Public Sub CoerceTableDateColumn(ByVal lo As ListObject, ByVal headerName As String)
    Dim lc As ListColumn
    Dim c As Range
    Dim dt As Date
    
    Set lc = GetListColumn(lo, headerName)
    If lc Is Nothing Then Exit Sub
    If lc.DataBodyRange Is Nothing Then Exit Sub
    
    For Each c In lc.DataBodyRange.Cells
        If Len(Trim$(CStr(c.Value))) > 0 Then
            If TryParseDateValue(c.Value, dt) Then
                c.Value = dt
            End If
        End If
    Next c
    
    lc.DataBodyRange.NumberFormat = "dd/mm/yyyy"
End Sub

Public Sub CoerceTableNumericColumns(ByVal lo As ListObject, ByVal headerList As Variant)
    Dim i As Long
    
    For i = LBound(headerList) To UBound(headerList)
        CoerceTableNumericColumn lo, CStr(headerList(i))
    Next i
End Sub

Public Sub CoerceTableNumericColumn(ByVal lo As ListObject, ByVal headerName As String)
    Dim lc As ListColumn
    Dim c As Range
    Dim s As String
    
    Set lc = GetListColumn(lo, headerName)
    If lc Is Nothing Then Exit Sub
    If lc.DataBodyRange Is Nothing Then Exit Sub
    
    For Each c In lc.DataBodyRange.Cells
        s = Trim$(CStr(c.Value))
        If Len(s) > 0 Then
            s = Replace$(s, ",", ".")
            If IsNumeric(s) Then c.Value = CDbl(s)
        End If
    Next c
    
    lc.DataBodyRange.NumberFormat = "General"
End Sub

Public Function TryParseDateValue(ByVal v As Variant, ByRef outDate As Date) As Boolean
    Dim s As String
    Dim parts() As String
    
    On Error GoTo Fail
    
    If IsDate(v) Then
        outDate = CDate(v)
        TryParseDateValue = True
        Exit Function
    End If
    
    If IsNumeric(v) Then
        If CDbl(v) > 0 Then
            outDate = DateSerial(1899, 12, 30) + CDbl(v)
            TryParseDateValue = True
            Exit Function
        End If
    End If
    
    s = Trim$(CStr(v))
    If Len(s) = 0 Then Exit Function
    
    If InStr(1, s, "T", vbTextCompare) > 0 Then
        s = Left$(s, InStr(1, s, "T", vbTextCompare) - 1)
    End If
    
    s = Replace$(s, ".", "/")
    s = Replace$(s, "-", "/")
    
    parts = Split(s, "/")
    If UBound(parts) = 2 Then
        If Len(parts(0)) = 4 Then
            outDate = DateSerial(CLng(parts(0)), CLng(parts(1)), CLng(parts(2)))
        Else
            outDate = DateSerial(CLng(parts(2)), CLng(parts(1)), CLng(parts(0)))
        End If
        TryParseDateValue = True
        Exit Function
    End If
    
Fail:
    TryParseDateValue = False
End Function

Public Function ToBoolean(ByVal v As Variant) As Boolean
    Dim s As String
    
    Select Case VarType(v)
        Case vbBoolean
            ToBoolean = CBool(v)
        Case vbByte, vbInteger, vbLong, vbSingle, vbDouble, vbCurrency, vbDecimal
            ToBoolean = (CDbl(v) <> 0)
        Case Else
            s = NormalizeText(CStr(v), True)
            ToBoolean = (s = "VERDADERO" Or s = "TRUE" Or s = "SI" Or s = "YES" Or s = "1")
    End Select
End Function

Public Function RowIsBlank(ByVal data As Variant, ByVal rowIndex As Long) As Boolean
    Dim c As Long
    
    For c = 1 To UBound(data, 2)
        If Len(Trim$(CStr(data(rowIndex, c)))) > 0 Then
            RowIsBlank = False
            Exit Function
        End If
    Next c
    
    RowIsBlank = True
End Function

Public Function LastUsedRow(ByVal ws As Worksheet) As Long
    Dim lastCell As Range
    
    Set lastCell = ws.Cells.Find(What:="*", After:=ws.Cells(1, 1), LookIn:=xlFormulas, SearchOrder:=xlByRows, SearchDirection:=xlPrevious)
    If lastCell Is Nothing Then
        LastUsedRow = 1
    Else
        LastUsedRow = lastCell.Row
    End If
End Function

Public Function LastUsedCol(ByVal ws As Worksheet) As Long
    Dim lastCell As Range
    
    Set lastCell = ws.Cells.Find(What:="*", After:=ws.Cells(1, 1), LookIn:=xlFormulas, SearchOrder:=xlByColumns, SearchDirection:=xlPrevious)
    If lastCell Is Nothing Then
        LastUsedCol = 1
    Else
        LastUsedCol = lastCell.Column
    End If
End Function

Public Function CanonicalHeader(ByVal txt As String) As String
    CanonicalHeader = NormalizeText(txt, True)
End Function

Public Function NormalizeText(ByVal txt As String, Optional ByVal compact As Boolean = False) As String
    Dim s As String
    Dim chars As Variant
    Dim i As Long
    
    s = UCase$(CStr(txt))
    
    s = Replace$(s, Chr$(160), " ")
    s = Replace$(s, vbCr, " ")
    s = Replace$(s, vbLf, " ")
    s = Replace$(s, vbTab, " ")
    
    s = Replace$(s, "Á", "A")
    s = Replace$(s, "À", "A")
    s = Replace$(s, "Ä", "A")
    s = Replace$(s, "Â", "A")
    
    s = Replace$(s, "É", "E")
    s = Replace$(s, "È", "E")
    s = Replace$(s, "Ë", "E")
    s = Replace$(s, "Ê", "E")
    
    s = Replace$(s, "Í", "I")
    s = Replace$(s, "Ì", "I")
    s = Replace$(s, "Ï", "I")
    s = Replace$(s, "Î", "I")
    
    s = Replace$(s, "Ó", "O")
    s = Replace$(s, "Ò", "O")
    s = Replace$(s, "Ö", "O")
    s = Replace$(s, "Ô", "O")
    
    s = Replace$(s, "Ú", "U")
    s = Replace$(s, "Ù", "U")
    s = Replace$(s, "Ü", "U")
    s = Replace$(s, "Û", "U")
    
    s = Replace$(s, "Ñ", "N")
    
    chars = Array(".", ",", ";", ":", "-", "_", "/", "\", "(", ")", "[", "]", "{", "}", "'", """")
    
    For i = LBound(chars) To UBound(chars)
        If compact Then
            s = Replace$(s, CStr(chars(i)), vbNullString)
        Else
            s = Replace$(s, CStr(chars(i)), " ")
        End If
    Next i
    
    On Error Resume Next
    s = Application.WorksheetFunction.Trim(s)
    On Error GoTo 0
    
    If compact Then s = Replace$(s, " ", vbNullString)
    
    NormalizeText = s
End Function

Public Function WorksheetExists(ByVal sheetName As String) As Boolean
    Dim ws As Worksheet
    
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(sheetName)
    On Error GoTo 0
    
    WorksheetExists = Not ws Is Nothing
End Function

Public Function TableExists(ByVal tableName As String) As Boolean
    Dim ws As Worksheet
    Dim lo As ListObject
    
    For Each ws In ThisWorkbook.Worksheets
        For Each lo In ws.ListObjects
            If StrComp(lo.Name, tableName, vbTextCompare) = 0 Then
                TableExists = True
                Exit Function
            End If
        Next lo
    Next ws
End Function

Public Function TableHasRows(ByVal lo As ListObject) As Boolean
    If lo Is Nothing Then Exit Function
    If lo.ListRows.Count = 0 Then Exit Function
    If lo.DataBodyRange Is Nothing Then Exit Function
    
    TableHasRows = True
End Function

Public Function TableHasRequiredColumns(ByVal lo As ListObject, ByVal requiredHeaders As Variant, Optional ByRef missingList As String = "") As Boolean
    Dim i As Long
    Dim headerName As String
    Dim missing As String
    
    If lo Is Nothing Then Exit Function
    
    For i = LBound(requiredHeaders) To UBound(requiredHeaders)
        headerName = CStr(requiredHeaders(i))
        
        If GetListColumn(lo, headerName) Is Nothing Then
            If Len(missing) > 0 Then missing = missing & ", "
            missing = missing & headerName
        End If
    Next i
    
    missingList = missing
    TableHasRequiredColumns = (Len(missing) = 0)
End Function