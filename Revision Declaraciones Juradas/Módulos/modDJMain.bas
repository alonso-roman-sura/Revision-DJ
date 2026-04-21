Option Explicit

Public Sub CargarColaboradores()
    Dim lo As ListObject
    
    On Error GoTo ErrHandler
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    
    Set lo = LoadColaboradoresTable()
    If lo Is Nothing Then GoTo SafeExit
    
    MsgBox "Base de colaboradores cargada correctamente.", vbInformation
    
SafeExit:
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    Exit Sub
    
ErrHandler:
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    MsgBox "Error en CargarColaboradores: " & Err.Description, vbExclamation
End Sub

Public Sub CargarReporteDJ()
    Dim lo As ListObject
    
    On Error GoTo ErrHandler
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    
    Set lo = LoadReporteDJTable()
    If lo Is Nothing Then GoTo SafeExit
    
    MsgBox "Reporte DJ cargado correctamente.", vbInformation
    
SafeExit:
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    Exit Sub
    
ErrHandler:
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    MsgBox "Error en CargarReporteDJ: " & Err.Description, vbExclamation
End Sub

Public Sub EjecutarComprobacion()
    Dim completed As Boolean
    
    On Error GoTo ErrHandler
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    
    completed = RunDeclarationCheck()
    
    If completed Then
        MsgBox "Comprobación ejecutada correctamente.", vbInformation
    End If
    
SafeExit:
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    Exit Sub
    
ErrHandler:
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    MsgBox "Error en EjecutarComprobacion: " & Err.Description, vbExclamation
End Sub

Public Sub EliminarDatos()
    Dim hasColab As Boolean
    Dim hasReporte As Boolean
    Dim resp As VbMsgBoxResult
    
    On Error GoTo ErrHandler
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.DisplayAlerts = False
    
    hasColab = WorksheetExists(SHEET_COLAB) Or TableExists(TABLE_COLAB)
    hasReporte = WorksheetExists(SHEET_REPORTE) Or TableExists(TABLE_REPORTE)
    
    If Not hasColab And Not hasReporte Then
        Application.DisplayAlerts = True
        Application.ScreenUpdating = True
        Application.EnableEvents = True
        MsgBox "No hay hojas ni tablas cargadas para eliminar.", vbInformation
        Exit Sub
    End If
    
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    
    resp = MsgBox( _
        Prompt:="Se eliminarán las hojas de Colaboradores y ReporteDJ si existen." & vbCrLf & vbCrLf & _
                "¿Deseas continuar?", _
        Buttons:=vbQuestion + vbYesNo + vbDefaultButton2, _
        Title:="Confirmar eliminación")
    
    If resp <> vbYes Then Exit Sub
    
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.DisplayAlerts = False
    
    If WorksheetExists(SHEET_COLAB) Then DeleteSheetIfExists SHEET_COLAB
    If WorksheetExists(SHEET_REPORTE) Then DeleteSheetIfExists SHEET_REPORTE
    
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    
    MsgBox "Se eliminaron las hojas y tablas de trabajo.", vbInformation
    Exit Sub
    
ErrHandler:
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    MsgBox "Error en EliminarDatos: " & Err.Description, vbExclamation
End Sub
