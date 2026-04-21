Option Explicit

Public WithEvents btn As MSForms.CommandButton
Public host As frmDateRangeDJ

Private Sub Btn_Click()
    If host Is Nothing Then Exit Sub
    host.HandleButtonClick btn.Name
End Sub