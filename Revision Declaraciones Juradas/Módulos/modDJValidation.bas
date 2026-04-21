Option Explicit

Public Function RunDeclarationCheck() As Boolean
    Dim startDate As Date
    Dim endDate As Date
    Dim statusHeader As String
    Dim loRep As ListObject
    Dim loCol As ListObject
    Dim dictStatus As Object
    
    ValidateCheckPrerequisites loRep, loCol
    
    If Not AskDateRange(startDate, endDate) Then Exit Function
    
    statusHeader = BuildStatusHeader(startDate, endDate)
    
    EnsureReportComputedColumns loRep, startDate, endDate, statusHeader
    Application.Calculate
    
    Set dictStatus = BuildReportStatusDictionary(loRep, "Nombres", "Apellidos", statusHeader)
    UpdateColaboradoresStatus loCol, dictStatus, statusHeader
    
    ClearSummaryBlock loRep
    WriteSummaryBlock loRep, statusHeader
    
    loRep.Parent.Cells.EntireColumn.AutoFit
    loCol.Parent.Cells.EntireColumn.AutoFit
    
    loRep.Parent.Activate
    loRep.Parent.Range("A1").Select
    
    RunDeclarationCheck = True
End Function

Private Sub ValidateCheckPrerequisites(ByRef loRep As ListObject, ByRef loCol As ListObject)
    Dim missing As String
    
    If Not TableExists(TABLE_COLAB) And Not TableExists(TABLE_REPORTE) Then
        Err.Raise vbObjectError + 2000, , "No hay datos cargados. Primero carga la base de colaboradores y el reporte DJ."
    End If
    
    If Not TableExists(TABLE_COLAB) Then
        Err.Raise vbObjectError + 2001, , "No se encontró la tabla de Colaboradores. Primero carga la base de colaboradores."
    End If
    
    If Not TableExists(TABLE_REPORTE) Then
        Err.Raise vbObjectError + 2002, , "No se encontró la tabla ReporteDJ. Primero carga el reporte de declaraciones juradas."
    End If
    
    Set loCol = GetRequiredTable(TABLE_COLAB)
    Set loRep = GetRequiredTable(TABLE_REPORTE)
    
    If Not TableHasRows(loCol) Then
        Err.Raise vbObjectError + 2003, , "La tabla Colaboradores no tiene registros válidos para evaluar."
    End If
    
    If Not TableHasRows(loRep) Then
        Err.Raise vbObjectError + 2004, , "La tabla ReporteDJ no tiene registros válidos para evaluar."
    End If
    
    If Not TableHasRequiredColumns(loCol, Array("Nombre Completo"), missing) Then
        Err.Raise vbObjectError + 2005, , "La tabla Colaboradores está incompleta. Faltan columnas: " & missing
    End If
    
    If Not TableHasRequiredColumns(loRep, Array("Nombres", "Apellidos", "Fecha de registro"), missing) Then
        Err.Raise vbObjectError + 2006, , "La tabla ReporteDJ está incompleta. Faltan columnas: " & missing
    End If
End Sub

Private Function AskDateRange(ByRef startDate As Date, ByRef endDate As Date) As Boolean
    Dim frm As frmDateRangeDJ
    
    Set frm = New frmDateRangeDJ
    frm.Show vbModal
    
    If frm.Cancelled Then
        Unload frm
        Set frm = Nothing
        Exit Function
    End If
    
    startDate = frm.StartDateValue
    endDate = frm.EndDateValue
    
    AskDateRange = True
    
    Unload frm
    Set frm = Nothing
End Function

Private Function BuildStatusHeader(ByVal startDate As Date, ByVal endDate As Date) As String
    If startDate = DateSerial(Year(startDate), 1, 1) _
       And endDate = Date Then
        BuildStatusHeader = "Llenado al " & Format$(endDate, "dd-mm-yyyy")
    ElseIf startDate = DateSerial(Year(startDate), 1, 1) _
       And endDate = DateSerial(Year(startDate), 12, 31) _
       And Year(startDate) = Year(endDate) Then
        BuildStatusHeader = "Llenado en " & Year(startDate)
    Else
        BuildStatusHeader = "Llenado entre " & Format$(startDate, "dd-mm-yyyy") & " y " & Format$(endDate, "dd-mm-yyyy")
    End If
End Function

Private Sub EnsureReportComputedColumns(ByVal lo As ListObject, ByVal startDate As Date, ByVal endDate As Date, ByVal statusHeader As String)
    Dim lcDoble As ListColumn
    Dim lcStatus As ListColumn
    
    RemoveTableColumnIfExists lo, "Doble Planilla"
    RemoveColumnsByPrefix lo, "Llenado en "
    RemoveColumnsByPrefix lo, "Llenado entre "
    RemoveColumnsByPrefix lo, "Llenado al "
    
    Set lcDoble = lo.ListColumns.Add
    lcDoble.Name = "Doble Planilla"
    
    Set lcStatus = lo.ListColumns.Add
    lcStatus.Name = statusHeader
    
    If Not lcDoble.DataBodyRange Is Nothing Then
        lcDoble.DataBodyRange.NumberFormat = "General"
        lcDoble.DataBodyRange.FormulaLocal = "=CONTAR.SI.CONJUNTO([Nombres];[@Nombres];[Apellidos];[@Apellidos])=2"
    End If
    
    If Not lcStatus.DataBodyRange Is Nothing Then
        lcStatus.DataBodyRange.NumberFormat = "General"
        lcStatus.DataBodyRange.FormulaLocal = BuildStatusFormulaLocal(startDate, endDate)
    End If
    
    ApplyOrangeHeaderFormat lo, "Doble Planilla"
    ApplyOrangeHeaderFormat lo, statusHeader
End Sub

Private Function BuildStatusFormulaLocal(ByVal startDate As Date, ByVal endDate As Date) As String
    BuildStatusFormulaLocal = _
        "=LET(" & _
        "inicio;FECHA(" & Year(startDate) & ";" & Month(startDate) & ";" & Day(startDate) & ");" & _
        "fin;FECHA(" & Year(endDate) & ";" & Month(endDate) & ";" & Day(endDate) & ");" & _
        "fecha;[@[Fecha de registro]];" & _
        "SI(" & _
        "[@[Doble Planilla]];" & _
        "CONTAR.SI.CONJUNTO([Nombres];[@Nombres];[Apellidos];[@Apellidos];[Fecha de registro];"">=""&inicio;[Fecha de registro];""<=""&fin)=2;" & _
        "Y(fecha>=inicio;fecha<=fin)" & _
        ")" & _
        ")"
End Function

Private Function BuildReportStatusDictionary(ByVal lo As ListObject, ByVal nombresHeader As String, ByVal apellidosHeader As String, ByVal statusHeader As String) As Object
    Dim dict As Object
    Dim lcNom As ListColumn
    Dim lcApe As ListColumn
    Dim lcStatus As ListColumn
    Dim arrNom As Variant
    Dim arrApe As Variant
    Dim arrStatus As Variant
    Dim i As Long
    Dim key As String
    Dim currentValue As Boolean
    
    Set dict = CreateObject("Scripting.Dictionary")
    dict.CompareMode = vbTextCompare
    
    Set lcNom = GetListColumn(lo, nombresHeader)
    Set lcApe = GetListColumn(lo, apellidosHeader)
    Set lcStatus = GetListColumn(lo, statusHeader)
    
    If lcNom Is Nothing Or lcApe Is Nothing Or lcStatus Is Nothing Then
        Err.Raise vbObjectError + 1800, , "Faltan columnas requeridas en ReporteDJ para construir la comprobación."
    End If
    
    If lcNom.DataBodyRange Is Nothing Then
        Set BuildReportStatusDictionary = dict
        Exit Function
    End If
    
    arrNom = lcNom.DataBodyRange.Value2
    arrApe = lcApe.DataBodyRange.Value2
    arrStatus = lcStatus.DataBodyRange.Value
    
    For i = 1 To UBound(arrNom, 1)
        key = NormalizeText(CStr(arrNom(i, 1)) & " " & CStr(arrApe(i, 1)), False)
        
        If Len(key) > 0 Then
            currentValue = ToBoolean(arrStatus(i, 1))
            
            If dict.Exists(key) Then
                dict(key) = (ToBoolean(dict(key)) Or currentValue)
            Else
                dict.Add key, currentValue
            End If
        End If
    Next i
    
    Set BuildReportStatusDictionary = dict
End Function

Private Sub UpdateColaboradoresStatus(ByVal lo As ListObject, ByVal dictStatus As Object, ByVal statusHeader As String)
    Dim lcNombre As ListColumn
    Dim lcKey As ListColumn
    Dim lcStatus As ListColumn
    Dim arrNombres As Variant
    Dim arrKeys() As Variant
    Dim arrStatus() As Variant
    Dim i As Long
    Dim n As Long
    Dim key As String
    
    RemoveTableColumnIfExists lo, "Clave Interna"
    RemoveColumnsByPrefix lo, "Llenado en "
    RemoveColumnsByPrefix lo, "Llenado entre "
    RemoveColumnsByPrefix lo, "Llenado al "
    
    Set lcNombre = GetListColumn(lo, "Nombre Completo")
    If lcNombre Is Nothing Then
        Err.Raise vbObjectError + 1900, , "No se encontró la columna 'Nombre Completo' en Colaboradores."
    End If
    
    Set lcKey = lo.ListColumns.Add
    lcKey.Name = "Clave Interna"
    
    Set lcStatus = lo.ListColumns.Add
    lcStatus.Name = statusHeader
    
    If lcNombre.DataBodyRange Is Nothing Then Exit Sub
    
    arrNombres = lcNombre.DataBodyRange.Value2
    n = UBound(arrNombres, 1)
    
    ReDim arrKeys(1 To n, 1 To 1)
    ReDim arrStatus(1 To n, 1 To 1)
    
    For i = 1 To n
        key = NormalizeText(CStr(arrNombres(i, 1)), False)
        arrKeys(i, 1) = key
        
        If dictStatus.Exists(key) Then
            arrStatus(i, 1) = ToBoolean(dictStatus(key))
        Else
            arrStatus(i, 1) = False
        End If
    Next i
    
    lcKey.DataBodyRange.NumberFormat = "@"
    lcKey.DataBodyRange.Value = arrKeys
    
    lcStatus.DataBodyRange.NumberFormat = "General"
    lcStatus.DataBodyRange.Value = arrStatus
    
    lcKey.Range.EntireColumn.Hidden = True
    
    HighlightFilledCollaborators lo, statusHeader
End Sub

Private Sub HighlightFilledCollaborators(ByVal lo As ListObject, ByVal statusHeader As String)
    Dim lcStatus As ListColumn
    Dim i As Long
    
    Set lcStatus = GetListColumn(lo, statusHeader)
    If lcStatus Is Nothing Then Exit Sub
    If lo.DataBodyRange Is Nothing Then Exit Sub
    
    ClearColaboradoresHighlight lo
    
    For i = 1 To lo.ListRows.Count
        If ToBoolean(lo.DataBodyRange.Cells(i, lcStatus.Index).Value) Then
            lo.DataBodyRange.Rows(i).Interior.Color = HighlightFillColor()
            lo.DataBodyRange.Rows(i).Font.Color = HighlightFontColor()
        End If
    Next i
End Sub

Private Sub ClearColaboradoresHighlight(ByVal lo As ListObject)
    If lo Is Nothing Then Exit Sub
    If lo.DataBodyRange Is Nothing Then Exit Sub
    
    lo.DataBodyRange.Interior.Pattern = xlNone
    lo.DataBodyRange.Font.ColorIndex = xlAutomatic
End Sub

Private Sub ApplyOrangeHeaderFormat(ByVal lo As ListObject, ByVal headerName As String)
    Dim lc As ListColumn
    
    Set lc = GetListColumn(lo, headerName)
    If lc Is Nothing Then Exit Sub
    
    With lc.Range.Cells(1, 1)
        .Interior.Color = OrangeAccentColor()
        .Font.Color = OrangeAccentFontColor()
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .Font.Bold = True
    End With
End Sub

Private Sub ClearSummaryBlock(ByVal loRep As ListObject)
    Dim ws As Worksheet
    Dim headerRow As Long
    Dim lastTableCol As Long
    
    Set ws = loRep.Parent
    headerRow = loRep.HeaderRowRange.Row
    lastTableCol = loRep.Range.Column + loRep.Range.Columns.Count - 1
    
    ws.Range(ws.Cells(headerRow, lastTableCol + 2), ws.Cells(headerRow, lastTableCol + 3)).Clear
End Sub

Private Sub WriteSummaryBlock(ByVal loRep As ListObject, ByVal statusHeader As String)
    Dim ws As Worksheet
    Dim headerRow As Long
    Dim lastTableCol As Long
    Dim lblCell As Range
    Dim valCell As Range
    
    Set ws = loRep.Parent
    headerRow = loRep.HeaderRowRange.Row
    lastTableCol = loRep.Range.Column + loRep.Range.Columns.Count - 1
    
    Set lblCell = ws.Cells(headerRow, lastTableCol + 2)
    Set valCell = ws.Cells(headerRow, lastTableCol + 3)
    
    lblCell.Value = "Faltan llenar"
    valCell.FormulaLocal = _
        "=LET(" & _
        "personas;Colaboradores[Clave Interna];" & _
        "faltantes;FILTRAR(personas;Colaboradores[" & statusHeader & "]=FALSO);" & _
        "SI.ERROR(CONTARA(UNICOS(faltantes));0)" & _
        ")"
    
    FormatSummaryCell lblCell, True
    FormatSummaryCell valCell, False
End Sub

Private Sub FormatSummaryCell(ByVal c As Range, ByVal makeBold As Boolean)
    Dim borderItem As Variant
    
    With c
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .Font.Bold = makeBold
    End With
    
    For Each borderItem In Array(xlEdgeLeft, xlEdgeTop, xlEdgeBottom, xlEdgeRight)
        With c.Borders(borderItem)
            .LineStyle = xlContinuous
            .Weight = xlThin
            .Color = SummaryBorderColor()
        End With
    Next borderItem
End Sub