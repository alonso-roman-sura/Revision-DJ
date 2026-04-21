Option Explicit

Private Const UI_FONT_NAME As String = "Segoe UI"
Private Const UI_FONT_SIZE As Single = 9#

Private mCancelled As Boolean
Private mStartDate As Date
Private mEndDate As Date

Private mButtonHandlers As Collection
Private mTextBoxHandlers As Collection

Public Property Get Cancelled() As Boolean
    Cancelled = mCancelled
End Property

Public Property Get StartDateValue() As Date
    StartDateValue = mStartDate
End Property

Public Property Get EndDateValue() As Date
    EndDateValue = mEndDate
End Property

Private Sub UserForm_Initialize()
    Set mButtonHandlers = New Collection
    Set mTextBoxHandlers = New Collection
    
    mCancelled = True
    
    BuildLayout
    LoadDefaultDates
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    If CloseMode = vbFormControlMenu Then
        Cancel = True
        mCancelled = True
        Me.Hide
    End If
End Sub

Public Sub HandleButtonClick(ByVal buttonName As String)
    Select Case UCase$(buttonName)
        Case "CMDOK"
            ConfirmDates
        
        Case "CMDCANCEL"
            mCancelled = True
            Me.Hide
    End Select
End Sub

Private Sub BuildLayout()
    Dim fraRange As MSForms.Frame
    Dim topRow1 As Single
    Dim topRow2 As Single
    Dim leftLabel As Single
    Dim leftBoxes As Single
    Dim boxWDay As Single
    Dim boxWMonth As Single
    Dim boxWYear As Single
    Dim boxH As Single
    Dim sepW As Single
    Dim gap As Single
    Dim buttonTop As Single
    
    With Me
        .Caption = "Filtro de fechas"
        .Width = 305
        .Height = 175
        .StartUpPosition = 1
        .BackColor = RGB(240, 240, 240)
        .Font.Name = UI_FONT_NAME
        .Font.Size = UI_FONT_SIZE
    End With
    
    Set fraRange = Me.Controls.Add("Forms.Frame.1", "fraRange", True)
    With fraRange
        .Caption = "Rango de fechas"
        .Left = 8
        .Top = 8
        .Width = 286
        .Height = 96
        .Font.Name = UI_FONT_NAME
        .Font.Size = UI_FONT_SIZE
        .Font.Bold = False
        .BackColor = RGB(240, 240, 240)
        .foreColor = RGB(0, 0, 0)
        .SpecialEffect = fmSpecialEffectEtched
    End With
    
    topRow1 = 24
    topRow2 = 56
    
    leftLabel = 10
    leftBoxes = 150
    
    boxWDay = 28
    boxWMonth = 28
    boxWYear = 42
    boxH = 20
    sepW = 8
    gap = 2
    
    AddLabel fraRange, "lblStart", "Fecha Inicio Campaña DJ:", leftLabel, topRow1 + 2, 132, 16, UI_FONT_SIZE
    AddLabel fraRange, "lblEnd", "Fecha Fin Campaña DJ:", leftLabel, topRow2 + 2, 132, 16, UI_FONT_SIZE
    
    BuildDateRow fraRange, "txtStartDay", "txtStartMonth", "txtStartYear", _
                 leftBoxes, topRow1, boxWDay, boxWMonth, boxWYear, boxH, sepW, gap, 0
    
    BuildDateRow fraRange, "txtEndDay", "txtEndMonth", "txtEndYear", _
                 leftBoxes, topRow2, boxWDay, boxWMonth, boxWYear, boxH, sepW, gap, 3
    
    buttonTop = 116
    
    HookButton AddButton(Me, "cmdOK", "Aceptar", 96, buttonTop, 72, 24, 6, True, False)
    HookButton AddButton(Me, "cmdCancel", "Cancelar", 178, buttonTop, 72, 24, 7, False, True)
End Sub

Private Sub BuildDateRow(ByVal parentObj As Object, _
                         ByVal dayName As String, ByVal monthName As String, ByVal yearName As String, _
                         ByVal startLeft As Single, ByVal topPos As Single, _
                         ByVal boxWDay As Single, ByVal boxWMonth As Single, ByVal boxWYear As Single, _
                         ByVal boxH As Single, ByVal sepW As Single, ByVal gap As Single, _
                         ByVal tabBase As Long)
    
    Dim txt As MSForms.TextBox
    Dim lbl As MSForms.Label
    Dim pos1 As Single
    Dim pos2 As Single
    Dim pos3 As Single
    
    pos1 = startLeft
    pos2 = pos1 + boxWDay + sepW + gap
    pos3 = pos2 + boxWMonth + sepW + gap
    
    Set txt = AddTextBox(parentObj, dayName, pos1, topPos, boxWDay, boxH, 2, tabBase)
    HookTextBox txt, False
    
    Set lbl = AddLabel(parentObj, "lbl" & dayName & "Sep", "/", pos1 + boxWDay + 1, topPos + 2, sepW, boxH, UI_FONT_SIZE)
    lbl.TextAlign = fmTextAlignCenter
    
    Set txt = AddTextBox(parentObj, monthName, pos2, topPos, boxWMonth, boxH, 2, tabBase + 1)
    HookTextBox txt, False
    
    Set lbl = AddLabel(parentObj, "lbl" & monthName & "Sep", "/", pos2 + boxWMonth + 1, topPos + 2, sepW, boxH, UI_FONT_SIZE)
    lbl.TextAlign = fmTextAlignCenter
    
    Set txt = AddTextBox(parentObj, yearName, pos3, topPos, boxWYear, boxH, 4, tabBase + 2)
    HookTextBox txt, True
End Sub

Private Sub LoadDefaultDates()
    Dim defaultStart As Date
    Dim defaultEnd As Date
    Dim fra As MSForms.Frame
    
    Set fra = Me.Controls("fraRange")
    
    defaultStart = DateSerial(Year(Date), 1, 1)
    defaultEnd = Date
    
    fra.Controls("txtStartDay").Value = Format$(Day(defaultStart), "00")
    fra.Controls("txtStartMonth").Value = Format$(Month(defaultStart), "00")
    fra.Controls("txtStartYear").Value = CStr(Year(defaultStart))
    
    fra.Controls("txtEndDay").Value = Format$(Day(defaultEnd), "00")
    fra.Controls("txtEndMonth").Value = Format$(Month(defaultEnd), "00")
    fra.Controls("txtEndYear").Value = CStr(Year(defaultEnd))
End Sub

Private Sub ConfirmDates()
    Dim startD As Date
    Dim endD As Date
    Dim fra As MSForms.Frame
    
    Set fra = Me.Controls("fraRange")
    
    If Not ParseDateParts( _
        CStr(fra.Controls("txtStartDay").Value), _
        CStr(fra.Controls("txtStartMonth").Value), _
        CStr(fra.Controls("txtStartYear").Value), _
        startD) Then
        
        MsgBox "La fecha inicial no es válida.", vbExclamation
        fra.Controls("txtStartDay").SetFocus
        Exit Sub
    End If
    
    If Not ParseDateParts( _
        CStr(fra.Controls("txtEndDay").Value), _
        CStr(fra.Controls("txtEndMonth").Value), _
        CStr(fra.Controls("txtEndYear").Value), _
        endD) Then
        
        MsgBox "La fecha final no es válida.", vbExclamation
        fra.Controls("txtEndDay").SetFocus
        Exit Sub
    End If
    
    If startD > endD Then
        MsgBox "La fecha inicial no puede ser mayor que la fecha final.", vbExclamation
        fra.Controls("txtStartDay").SetFocus
        Exit Sub
    End If
    
    mStartDate = startD
    mEndDate = endD
    mCancelled = False
    
    Me.Hide
End Sub

Private Function ParseDateParts(ByVal dayText As String, ByVal monthText As String, ByVal yearText As String, ByRef outDate As Date) As Boolean
    Dim d As Long
    Dim m As Long
    Dim y As Long
    Dim dt As Date
    
    On Error GoTo Fail
    
    dayText = Trim$(dayText)
    monthText = Trim$(monthText)
    yearText = Trim$(yearText)
    
    If Len(dayText) = 0 Or Len(monthText) = 0 Or Len(yearText) = 0 Then Exit Function
    If Not IsNumeric(dayText) Or Not IsNumeric(monthText) Or Not IsNumeric(yearText) Then Exit Function
    
    d = CLng(dayText)
    m = CLng(monthText)
    y = CLng(yearText)
    
    If d < 1 Or d > 31 Then Exit Function
    If m < 1 Or m > 12 Then Exit Function
    If y < 1900 Or y > 9999 Then Exit Function
    
    dt = DateSerial(y, m, d)
    
    If Day(dt) <> d Or Month(dt) <> m Or Year(dt) <> y Then Exit Function
    
    outDate = dt
    ParseDateParts = True
    Exit Function

Fail:
    ParseDateParts = False
End Function

Private Function AddLabel(ByVal parentObj As Object, ByVal ctrlName As String, ByVal captionText As String, _
                          ByVal posLeft As Single, ByVal posTop As Single, _
                          ByVal ctrlWidth As Single, ByVal ctrlHeight As Single, _
                          ByVal fontSize As Single) As MSForms.Label
    Dim lbl As MSForms.Label
    
    Set lbl = parentObj.Controls.Add("Forms.Label.1", ctrlName, True)
    
    With lbl
        .Caption = captionText
        .Left = posLeft
        .Top = posTop
        .Width = ctrlWidth
        .Height = ctrlHeight
        .Font.Name = UI_FONT_NAME
        .Font.Size = fontSize
        .Font.Bold = False
        .BackStyle = fmBackStyleTransparent
        .foreColor = RGB(0, 0, 0)
    End With
    
    Set AddLabel = lbl
End Function

Private Function AddTextBox(ByVal parentObj As Object, ByVal ctrlName As String, _
                            ByVal posLeft As Single, ByVal posTop As Single, _
                            ByVal ctrlWidth As Single, ByVal ctrlHeight As Single, _
                            ByVal maxLen As Long, ByVal tabIndex As Long) As MSForms.TextBox
    Dim txt As MSForms.TextBox
    
    Set txt = parentObj.Controls.Add("Forms.TextBox.1", ctrlName, True)
    
    With txt
        .Left = posLeft
        .Top = posTop
        .Width = ctrlWidth
        .Height = ctrlHeight
        .MaxLength = maxLen
        .tabIndex = tabIndex
        .TextAlign = fmTextAlignCenter
        .BorderStyle = fmBorderStyleSingle
        .SpecialEffect = fmSpecialEffectFlat
        .Font.Name = UI_FONT_NAME
        .Font.Size = UI_FONT_SIZE
        .Font.Bold = False
        .BackColor = RGB(255, 255, 255)
        .foreColor = RGB(0, 0, 0)
    End With
    
    Set AddTextBox = txt
End Function

Private Function AddButton(ByVal parentObj As Object, ByVal ctrlName As String, ByVal captionText As String, _
                           ByVal posLeft As Single, ByVal posTop As Single, _
                           ByVal ctrlWidth As Single, ByVal ctrlHeight As Single, _
                           ByVal tabIndex As Long, ByVal isDefault As Boolean, ByVal isCancel As Boolean) As MSForms.CommandButton
    Dim btn As MSForms.CommandButton
    
    Set btn = parentObj.Controls.Add("Forms.CommandButton.1", ctrlName, True)
    
    With btn
        .Caption = captionText
        .Left = posLeft
        .Top = posTop
        .Width = ctrlWidth
        .Height = ctrlHeight
        .tabIndex = tabIndex
        .Font.Name = UI_FONT_NAME
        .Font.Size = UI_FONT_SIZE
        .Font.Bold = False
        .TakeFocusOnClick = False
        .BackColor = RGB(242, 242, 242)
        .Default = isDefault
        .Cancel = isCancel
    End With
    
    Set AddButton = btn
End Function

Private Sub HookButton(ByVal btn As MSForms.CommandButton)
    Dim handler As CDateRangeButtonHandler
    
    Set handler = New CDateRangeButtonHandler
    Set handler.btn = btn
    Set handler.host = Me
    
    mButtonHandlers.Add handler, btn.Name
End Sub

Private Sub HookTextBox(ByVal txt As MSForms.TextBox, ByVal isYearBox As Boolean)
    Dim handler As CDateRangeTextBoxHandler
    
    Set handler = New CDateRangeTextBoxHandler
    Set handler.txt = txt
    handler.isYearBox = isYearBox
    
    mTextBoxHandlers.Add handler, txt.Name
End Sub