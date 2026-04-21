Option Explicit

Public Function LoadColaboradoresTable() As ListObject
    Dim filePath As String
    Dim srcWb As Workbook
    Dim srcWs As Worksheet
    Dim hdrRow As Long
    Dim hdrMap As Object
    Dim arrOut As Variant
    Dim lo As ListObject
    Dim score As Long
    
    filePath = PickSourceFile("Seleccionar base de datos de colaboradores")
    If Len(filePath) = 0 Then Exit Function
    
    Set srcWb = Workbooks.Open(Filename:=filePath, ReadOnly:=True, Local:=True)
    
    On Error GoTo CleanFail
    
    Set srcWs = GetBestSourceSheet(srcWb, "ACTIVOS", ColabOutputHeaders())
    If srcWs Is Nothing Then
        Err.Raise vbObjectError + 1000, , "No se encontró una hoja válida para colaboradores."
    End If
    
    hdrRow = DetectHeaderRow(srcWs, ColabOutputHeaders())
    If hdrRow = 0 Then
        Err.Raise vbObjectError + 1001, , "No se pudo detectar la fila de encabezados en la base de colaboradores."
    End If
    
    Set hdrMap = HeaderMap(srcWs, hdrRow)
    
    If HeaderExists(hdrMap, "Fecha de salida") Then
        Err.Raise vbObjectError + 1002, , "El archivo no corresponde a la base requerida. Contiene la columna 'Fecha de salida'."
    End If
    
    score = HeaderScoreMap(hdrMap, ColabOutputHeaders())
    If score < MIN_COLAB_HEADER_MATCHES Then
        Err.Raise vbObjectError + 1003, , "La estructura del archivo de colaboradores no cumple la validación mínima."
    End If
    
    If Not HasEssentialHeaders(hdrMap, ColabEssentialHeaders()) Then
        Err.Raise vbObjectError + 1004, , "Faltan columnas esenciales en la base de colaboradores."
    End If
    
    arrOut = BuildOutputArray(srcWs, hdrRow, ColabOutputHeaders(), True, ColabTextHeaders())
    
    Set lo = WriteArrayToNewTable(arrOut, SHEET_COLAB, TABLE_COLAB, TableStyleColaboradores())
    
    ForceTableColumnsAsText lo, ColabTextHeaders()
    CoerceTableDateColumns lo, ColabDateHeaders()
    CoerceTableNumericColumns lo, ColabNumericHeaders()
    
    lo.Parent.Cells.EntireColumn.AutoFit
    
    Set LoadColaboradoresTable = lo
    
CleanExit:
    On Error Resume Next
    If Not srcWb Is Nothing Then srcWb.Close SaveChanges:=False
    Exit Function
    
CleanFail:
    On Error Resume Next
    If Not srcWb Is Nothing Then srcWb.Close SaveChanges:=False
    Err.Raise Err.Number, , Err.Description
End Function

Public Function LoadReporteDJTable() As ListObject
    Dim filePath As String
    Dim srcWb As Workbook
    Dim srcWs As Worksheet
    Dim hdrRow As Long
    Dim hdrMap As Object
    Dim arrOut As Variant
    Dim lo As ListObject
    Dim score As Long
    
    filePath = PickSourceFile("Seleccionar reporte de declaraciones juradas")
    If Len(filePath) = 0 Then Exit Function
    
    Set srcWb = Workbooks.Open(Filename:=filePath, ReadOnly:=True, Local:=True)
    
    On Error GoTo CleanFail
    
    Set srcWs = GetBestSourceSheet(srcWb, vbNullString, ReporteOutputHeaders())
    If srcWs Is Nothing Then
        Err.Raise vbObjectError + 1100, , "No se encontró una hoja válida para el reporte DJ."
    End If
    
    hdrRow = DetectHeaderRow(srcWs, ReporteOutputHeaders())
    If hdrRow = 0 Then
        Err.Raise vbObjectError + 1101, , "No se pudo detectar la fila de encabezados en el reporte DJ."
    End If
    
    Set hdrMap = HeaderMap(srcWs, hdrRow)
    
    score = HeaderScoreMap(hdrMap, ReporteOutputHeaders())
    If score < MIN_REPORTE_HEADER_MATCHES Then
        Err.Raise vbObjectError + 1102, , "La estructura del archivo del reporte DJ no cumple la validación mínima."
    End If
    
    If Not HasEssentialHeaders(hdrMap, ReporteEssentialHeaders()) Then
        Err.Raise vbObjectError + 1103, , "Faltan columnas esenciales en el reporte DJ."
    End If
    
    arrOut = BuildOutputArray(srcWs, hdrRow, ReporteOutputHeaders(), False, ReporteTextHeaders())
    
    Set lo = WriteArrayToNewTable(arrOut, SHEET_REPORTE, TABLE_REPORTE, TableStyleReporteDJ())
    
    ForceTableColumnsAsText lo, ReporteTextHeaders()
    CoerceTableDateColumns lo, ReporteDateHeaders()
    
    lo.Parent.Cells.EntireColumn.AutoFit
    
    Set LoadReporteDJTable = lo
    
CleanExit:
    On Error Resume Next
    If Not srcWb Is Nothing Then srcWb.Close SaveChanges:=False
    Exit Function
    
CleanFail:
    On Error Resume Next
    If Not srcWb Is Nothing Then srcWb.Close SaveChanges:=False
    Err.Raise Err.Number, , Err.Description
End Function

Private Function GetBestSourceSheet(ByVal wb As Workbook, ByVal preferredName As String, ByVal expectedHeaders As Variant) As Worksheet
    Dim ws As Worksheet
    Dim bestWs As Worksheet
    Dim score As Long
    Dim bestScore As Long
    
    On Error Resume Next
    If Len(preferredName) > 0 Then
        Set ws = wb.Worksheets(preferredName)
        If Not ws Is Nothing Then
            If BestHeaderScoreOnSheet(ws, expectedHeaders) > 0 Then
                Set GetBestSourceSheet = ws
                Exit Function
            End If
        End If
    End If
    On Error GoTo 0
    
    For Each ws In wb.Worksheets
        score = BestHeaderScoreOnSheet(ws, expectedHeaders)
        If score > bestScore Then
            bestScore = score
            Set bestWs = ws
        End If
    Next ws
    
    Set GetBestSourceSheet = bestWs
End Function

Private Function BestHeaderScoreOnSheet(ByVal ws As Worksheet, ByVal expectedHeaders As Variant) As Long
    Dim r As Long
    Dim maxRow As Long
    Dim score As Long
    Dim bestScore As Long
    
    maxRow = Application.Min(HEADER_SCAN_ROWS, LastUsedRow(ws))
    
    For r = 1 To maxRow
        score = CountHeaderMatchesInRow(ws, r, expectedHeaders)
        If score > bestScore Then bestScore = score
    Next r
    
    BestHeaderScoreOnSheet = bestScore
End Function

Private Function DetectHeaderRow(ByVal ws As Worksheet, ByVal expectedHeaders As Variant) As Long
    Dim r As Long
    Dim maxRow As Long
    Dim score As Long
    Dim bestScore As Long
    Dim bestRow As Long
    
    maxRow = Application.Min(HEADER_SCAN_ROWS, LastUsedRow(ws))
    
    For r = 1 To maxRow
        score = CountHeaderMatchesInRow(ws, r, expectedHeaders)
        If score > bestScore Then
            bestScore = score
            bestRow = r
        End If
    Next r
    
    DetectHeaderRow = bestRow
End Function

Private Function CountHeaderMatchesInRow(ByVal ws As Worksheet, ByVal hdrRow As Long, ByVal expectedHeaders As Variant) As Long
    Dim map As Object
    Set map = HeaderMap(ws, hdrRow)
    CountHeaderMatchesInRow = HeaderScoreMap(map, expectedHeaders)
End Function

Private Function HeaderMap(ByVal ws As Worksheet, ByVal hdrRow As Long) As Object
    Dim d As Object
    Dim lastCol As Long
    Dim c As Long
    Dim rawHeader As String
    Dim key As String
    
    Set d = CreateObject("Scripting.Dictionary")
    d.CompareMode = vbTextCompare
    
    lastCol = LastUsedCol(ws)
    
    For c = 1 To lastCol
        rawHeader = Trim$(CStr(ws.Cells(hdrRow, c).Value))
        If Len(rawHeader) > 0 Then
            key = CanonicalHeader(rawHeader)
            If Len(key) > 0 Then
                If Not d.Exists(key) Then d.Add key, c
            End If
        End If
    Next c
    
    Set HeaderMap = d
End Function

Private Function HeaderScoreMap(ByVal hdrMap As Object, ByVal expectedHeaders As Variant) As Long
    Dim i As Long
    Dim score As Long
    
    For i = LBound(expectedHeaders) To UBound(expectedHeaders)
        If ResolveHeaderIndex(hdrMap, CStr(expectedHeaders(i))) > 0 Then
            score = score + 1
        End If
    Next i
    
    HeaderScoreMap = score
End Function

Private Function HasEssentialHeaders(ByVal hdrMap As Object, ByVal essentialHeaders As Variant) As Boolean
    Dim i As Long
    
    For i = LBound(essentialHeaders) To UBound(essentialHeaders)
        If ResolveHeaderIndex(hdrMap, CStr(essentialHeaders(i))) = 0 Then
            HasEssentialHeaders = False
            Exit Function
        End If
    Next i
    
    HasEssentialHeaders = True
End Function

Private Function HeaderExists(ByVal hdrMap As Object, ByVal headerName As String) As Boolean
    HeaderExists = (ResolveHeaderIndex(hdrMap, headerName) > 0)
End Function

Private Function ResolveHeaderIndex(ByVal hdrMap As Object, ByVal targetHeader As String) As Long
    Dim aliases As Variant
    Dim i As Long
    Dim aliasKey As String
    
    aliases = GetHeaderAliases(targetHeader)
    
    For i = LBound(aliases) To UBound(aliases)
        aliasKey = CStr(aliases(i))
        If hdrMap.Exists(aliasKey) Then
            ResolveHeaderIndex = CLng(hdrMap(aliasKey))
            Exit Function
        End If
    Next i
End Function

Private Function BuildOutputArray(ByVal ws As Worksheet, ByVal hdrRow As Long, ByVal outHeaders As Variant, ByVal filterCountry As Boolean, ByVal textHeaders As Variant) As Variant
    Dim hdrMap As Object
    Dim lastRow As Long
    Dim lastCol As Long
    Dim data As Variant
    Dim outArr() As Variant
    Dim finalArr As Variant
    Dim r As Long
    Dim c As Long
    Dim outRow As Long
    Dim colCount As Long
    Dim srcCol As Long
    Dim countryCol As Long
    Dim keepRow As Boolean
    Dim headerName As String
    Dim cellValue As Variant
    
    Set hdrMap = HeaderMap(ws, hdrRow)
    
    lastRow = LastUsedRow(ws)
    lastCol = LastUsedCol(ws)
    
    If lastRow <= hdrRow Then
        Err.Raise vbObjectError + 1500, , "No se encontraron filas de datos debajo del encabezado."
    End If
    
    data = ws.Range(ws.Cells(hdrRow + 1, 1), ws.Cells(lastRow, lastCol)).Value2
    colCount = UBound(outHeaders) - LBound(outHeaders) + 1
    
    ReDim outArr(1 To UBound(data, 1) + 1, 1 To colCount)
    
    For c = LBound(outHeaders) To UBound(outHeaders)
        outArr(1, c - LBound(outHeaders) + 1) = CStr(outHeaders(c))
    Next c
    
    If filterCountry Then
        countryCol = ResolveHeaderIndex(hdrMap, "País")
        If countryCol = 0 Then
            Err.Raise vbObjectError + 1501, , "No se encontró la columna 'País' para filtrar Perú."
        End If
    End If
    
    outRow = 1
    
    For r = 1 To UBound(data, 1)
        If Not RowIsBlank(data, r) Then
            keepRow = True
            
            If filterCountry Then
                keepRow = (NormalizeText(CStr(ws.Cells(hdrRow + r, countryCol).Text), True) = "PERU")
            End If
            
            If keepRow Then
                outRow = outRow + 1
                
                For c = LBound(outHeaders) To UBound(outHeaders)
                    headerName = CStr(outHeaders(c))
                    srcCol = ResolveHeaderIndex(hdrMap, headerName)
                    
                    If srcCol > 0 Then
                        If IsHeaderInList(headerName, textHeaders) Then
                            cellValue = CStr(ws.Cells(hdrRow + r, srcCol).Text)
                        Else
                            cellValue = ws.Cells(hdrRow + r, srcCol).Value2
                        End If
                        
                        outArr(outRow, c - LBound(outHeaders) + 1) = cellValue
                    Else
                        outArr(outRow, c - LBound(outHeaders) + 1) = vbNullString
                    End If
                Next c
            End If
        End If
    Next r
    
    If outRow = 1 Then
        Err.Raise vbObjectError + 1502, , "No quedaron registros luego del filtrado aplicado."
    End If
    
    finalArr = Trim2DArrayRows(outArr, outRow, colCount)
    BuildOutputArray = finalArr
End Function

Private Function IsHeaderInList(ByVal headerName As String, ByVal headerList As Variant) As Boolean
    Dim i As Long
    
    For i = LBound(headerList) To UBound(headerList)
        If CanonicalHeader(headerName) = CanonicalHeader(CStr(headerList(i))) Then
            IsHeaderInList = True
            Exit Function
        End If
    Next i
End Function

Private Function Trim2DArrayRows(ByVal arr As Variant, ByVal usedRows As Long, ByVal usedCols As Long) As Variant
    Dim tmp() As Variant
    Dim r As Long
    Dim c As Long
    
    ReDim tmp(1 To usedRows, 1 To usedCols)
    
    For r = 1 To usedRows
        For c = 1 To usedCols
            tmp(r, c) = arr(r, c)
        Next c
    Next r
    
    Trim2DArrayRows = tmp
End Function

Private Function WriteArrayToNewTable(ByVal arr As Variant, ByVal sheetName As String, ByVal tableName As String, ByVal tableStyleName As String) As ListObject
    Dim ws As Worksheet
    Dim rg As Range
    Dim lo As ListObject
    
    Application.DisplayAlerts = False
    DeleteSheetIfExists sheetName
    Application.DisplayAlerts = True
    
    Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
    ws.Name = sheetName
    
    Set rg = ws.Range("A1").Resize(UBound(arr, 1), UBound(arr, 2))
    rg.Value = arr
    
    Set lo = ws.ListObjects.Add(SourceType:=xlSrcRange, Source:=rg, XlListObjectHasHeaders:=xlYes)
    lo.Name = tableName
    lo.TableStyle = tableStyleName
    
    Set WriteArrayToNewTable = lo
End Function
