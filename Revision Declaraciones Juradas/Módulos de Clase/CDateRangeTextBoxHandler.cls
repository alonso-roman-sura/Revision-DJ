Option Explicit

Public WithEvents txt As MSForms.TextBox
Public isYearBox As Boolean

Private Sub Txt_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    Select Case KeyAscii
        Case 8, 48 To 57
        Case Else
            KeyAscii = 0
    End Select
End Sub

Private Sub Txt_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    Dim v As String
    
    v = Trim$(txt.Value)
    If Len(v) = 0 Then Exit Sub
    If Not IsNumeric(v) Then Exit Sub
    
    If isYearBox Then
        txt.Value = CStr(CLng(v))
    Else
        txt.Value = Format$(CLng(v), "00")
    End If
End Sub