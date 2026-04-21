Option Explicit

Public Const SHEET_COLAB As String = "Colaboradores"
Public Const TABLE_COLAB As String = "Colaboradores"

Public Const SHEET_REPORTE As String = "ReporteDJ"
Public Const TABLE_REPORTE As String = "ReporteDJ"

Public Const HEADER_SCAN_ROWS As Long = 25
Public Const MIN_COLAB_HEADER_MATCHES As Long = 8
Public Const MIN_REPORTE_HEADER_MATCHES As Long = 5

Public Function TableStyleColaboradores() As String
    TableStyleColaboradores = "TableStyleMedium1"
End Function

Public Function TableStyleReporteDJ() As String
    TableStyleReporteDJ = "TableStyleMedium2"
End Function

Public Function OrangeAccentColor() As Long
    OrangeAccentColor = RGB(237, 125, 49)
End Function

Public Function OrangeAccentFontColor() As Long
    OrangeAccentFontColor = RGB(255, 255, 255)
End Function

Public Function HighlightFillColor() As Long
    HighlightFillColor = RGB(146, 208, 80)
End Function

Public Function HighlightFontColor() As Long
    HighlightFontColor = RGB(0, 0, 0)
End Function

Public Function SummaryBorderColor() As Long
    SummaryBorderColor = RGB(68, 179, 225)
End Function

Public Function ColabOutputHeaders() As Variant
    ColabOutputHeaders = Array( _
        "País", _
        "Nombre Completo", _
        "Líder / Colaborador", _
        "Nombre Posición", _
        "Compañía", _
        "Nivel Organizacional 2", _
        "Nivel Organizacional 3", _
        "Nivel Organizacional 4", _
        "Nivel Organizacional 5", _
        "Fecha de ingreso", _
        "Fecha de antigüedad", _
        "Nombre Líder Directo", _
        "Correo Comunicados", _
        "Fecha de Nacimiento", _
        "Género", _
        "Ubicación Física", _
        "Administrativo / Ventas - gasto", _
        "Antigüedad", _
        "Edad", _
        "Generación", _
        "País donde se ejecuta el gasto (dato para finanzas)" _
    )
End Function

Public Function ColabEssentialHeaders() As Variant
    ColabEssentialHeaders = Array( _
        "País", _
        "Nombre Completo", _
        "Compañía" _
    )
End Function

Public Function ColabTextHeaders() As Variant
    ColabTextHeaders = Array( _
        "País", _
        "Nombre Completo", _
        "Líder / Colaborador", _
        "Nombre Posición", _
        "Compañía", _
        "Nivel Organizacional 2", _
        "Nivel Organizacional 3", _
        "Nivel Organizacional 4", _
        "Nivel Organizacional 5", _
        "Nombre Líder Directo", _
        "Correo Comunicados", _
        "Género", _
        "Ubicación Física", _
        "Administrativo / Ventas - gasto", _
        "Generación", _
        "País donde se ejecuta el gasto (dato para finanzas)" _
    )
End Function

Public Function ColabDateHeaders() As Variant
    ColabDateHeaders = Array( _
        "Fecha de ingreso", _
        "Fecha de antigüedad", _
        "Fecha de Nacimiento" _
    )
End Function

Public Function ColabNumericHeaders() As Variant
    ColabNumericHeaders = Array( _
        "Antigüedad", _
        "Edad" _
    )
End Function

Public Function ReporteOutputHeaders() As Variant
    ReporteOutputHeaders = Array( _
        "ID de usuario/empleado de Talentum", _
        "Nombres", _
        "Apellidos", _
        "Compañia", _
        "Fecha de registro", _
        "Adjunto declaración" _
    )
End Function

Public Function ReporteEssentialHeaders() As Variant
    ReporteEssentialHeaders = Array( _
        "ID de usuario/empleado de Talentum", _
        "Nombres", _
        "Apellidos", _
        "Compañia", _
        "Fecha de registro", _
        "Adjunto declaración" _
    )
End Function

Public Function ReporteTextHeaders() As Variant
    ReporteTextHeaders = Array( _
        "ID de usuario/empleado de Talentum", _
        "Nombres", _
        "Apellidos", _
        "Compañia", _
        "Adjunto declaración" _
    )
End Function

Public Function ReporteDateHeaders() As Variant
    ReporteDateHeaders = Array( _
        "Fecha de registro" _
    )
End Function

Public Function GetHeaderAliases(ByVal targetHeader As String) As Variant
    Dim key As String
    key = CanonicalHeader(targetHeader)
    
    Select Case key
        Case CanonicalHeader("Antigüedad")
            GetHeaderAliases = Array( _
                CanonicalHeader("Antigüedad"), _
                CanonicalHeader("Antigueüedad") _
            )
        
        Case CanonicalHeader("Fecha de antigüedad")
            GetHeaderAliases = Array( _
                CanonicalHeader("Fecha de antigüedad"), _
                CanonicalHeader("Fecha de antigueüedad") _
            )
        
        Case Else
            GetHeaderAliases = Array(key)
    End Select
End Function