VERSION 5.00
Begin VB.UserControl ProgressGuage 
   BackColor       =   &H00FFFFFF&
   CanGetFocus     =   0   'False
   ClientHeight    =   120
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   120
   FillColor       =   &H00FF0000&
   BeginProperty Font 
      Name            =   "Small Fonts"
      Size            =   6
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H00FF0000&
   ScaleHeight     =   8
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   8
   ToolboxBitmap   =   "PGuage.ctx":0000
End
Attribute VB_Name = "ProgressGuage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Enum BorderStyles
    [None] = 0
    [Fixed Single] = 1
End Enum

Private Const DEF_SUFFIX As String = "%"
Private Const DEF_BORDERSTYLE As Integer = 1

Private Type GUAGE_DATA
    Min As Long
    Max As Long
    Value As Long
    Alignment As AlignmentConstants
    Suffix As String
    MarkValue As Long
    UseMark As Boolean
    MarkColor As Long
End Type
Private mGuage As GUAGE_DATA

' -- API --

Private Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

Private Const DT_TOP = &H0
Private Const DT_LEFT = &H0
Private Const DT_CENTER = &H1
Private Const DT_RIGHT = &H2
Private Const DT_VCENTER = &H4
Private Const DT_BOTTOM = &H8
Private Const DT_SINGLELINE = &H20
Private Const DT_CALCRECT = &H400
Private Const DT_NOPREFIX = &H800

Private Declare Function GetSysColor Lib "user32" (ByVal nIndex As Long) As Long

Private Declare Function GetClientRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long
Private Declare Function InflateRect Lib "user32" (lpRect As RECT, ByVal X As Long, ByVal Y As Long) As Long

Private Declare Function DrawText Lib "user32" Alias "DrawTextA" (ByVal hDC As Long, ByVal lpStr As String, ByVal nCount As Long, lpRect As RECT, ByVal wFormat As Long) As Long

Private Declare Function CreateSolidBrush Lib "gdi32" (ByVal crColor As Long) As Long
Private Declare Function FillRect Lib "user32" (ByVal hDC As Long, lpRect As RECT, ByVal hBrush As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long

Public Event MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Public Event MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Public Event MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

Private Sub UserControl_InitProperties()
    With UserControl
        .BorderStyle = DEF_BORDERSTYLE
        .BackColor = vbWindowBackground
        .ForeColor = vbBlue
    End With
    With mGuage
        .Min = 0
        .Max = 100
        .Value = 0
        .Alignment = vbCenter
        .Suffix = DEF_SUFFIX
        .MarkValue = 0
        .UseMark = False
    End With
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    With PropBag
        UserControl.BorderStyle = .ReadProperty("BorderStyle", DEF_BORDERSTYLE)
        UserControl.BackColor = .ReadProperty("BackColor", vbWindowBackground)
        UserControl.ForeColor = .ReadProperty("ForeColor", vbBlue)
        UserControl.FillColor = .ReadProperty("FillColor", vbBlue)
        mGuage.MarkColor = .ReadProperty("MarkColor", vbBlack)
'        Set UserControl.Font = .ReadProperty("Font", Ambient.Font)

        mGuage.Min = .ReadProperty("Min", 0)
        mGuage.Max = .ReadProperty("Max", 100)
        mGuage.Value = .ReadProperty("Value", 0)
        mGuage.Alignment = .ReadProperty("Alignment", vbCenter)
        mGuage.Suffix = .ReadProperty("Suffix", DEF_SUFFIX)
        mGuage.MarkValue = .ReadProperty("MarkValue", 0)
        mGuage.UseMark = .ReadProperty("UseMark", False)
    End With

    ' Validate value range
    If mGuage.Value < mGuage.Min Then
        mGuage.Value = mGuage.Min
    ElseIf mGuage.Value > mGuage.Max Then
        mGuage.Value = mGuage.Max
    End If
    If mGuage.MarkValue < mGuage.Min Then
        mGuage.MarkValue = mGuage.Min
    ElseIf mGuage.MarkValue > mGuage.Max Then
        mGuage.MarkValue = mGuage.Max
    End If
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    With PropBag
        .WriteProperty "BorderStyle", UserControl.BorderStyle, DEF_BORDERSTYLE
        .WriteProperty "BackColor", UserControl.BackColor, vbWindowBackground
        .WriteProperty "ForeColor", UserControl.ForeColor, vbBlue
        .WriteProperty "FillColor", UserControl.FillColor, vbBlue
        .WriteProperty "MarkColor", mGuage.MarkColor, vbBlack
'        .WriteProperty "Font", UserControl.Font, Ambient.Font

        .WriteProperty "Min", mGuage.Min, 0
        .WriteProperty "Max", mGuage.Max, 100
        .WriteProperty "Value", mGuage.Value, 0
        .WriteProperty "Alignment", mGuage.Alignment, vbCenter
        .WriteProperty "Suffix", mGuage.Suffix, DEF_SUFFIX
        .WriteProperty "MarkValue", mGuage.MarkValue, 0
        .WriteProperty "UseMark", mGuage.UseMark, False
    End With
End Sub

Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseDown(Button, Shift, X, Y)
End Sub

Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseMove(Button, Shift, X, Y)
End Sub

Private Sub UserControl_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseUp(Button, Shift, X, Y)
End Sub

Private Sub UserControl_Paint()
    Dim tRect As RECT, tText As RECT, tMarkText As RECT
    Dim hWinDC As Long, nWidth As Long, nHeight As Long, wFormat As Long, _
        wMarkFormat As Long, nMarkDiff As Long, _
        hBrush As Long, nOffset As Long, nRight As Long, nColor As Long, _
        nForeColor As Long, nMark As Long
    Dim sTitle As String, sMarkTitle As String

    ' Set the text we will draw into a variable
    sTitle = Trim$(CStr(mGuage.Value) & mGuage.Suffix)

    ' Work out the co-ords for the percentage bar
    Call GetClientRect(UserControl.hwnd, tRect)
    nWidth = tRect.Right
    nHeight = tRect.Bottom
    '
    If mGuage.Value > mGuage.Min Then
        nRight = ((nWidth / mGuage.Max) * (mGuage.Value - mGuage.Min)) - 1
    Else
        nRight = -1
    End If

    If mGuage.UseMark Then
        Select Case mGuage.MarkValue
        Case mGuage.Min
            nMark = 0
        Case mGuage.Max
            nMark = nWidth
        Case mGuage.Min To mGuage.Max
            nMark = ((nWidth / mGuage.Max) * (mGuage.MarkValue - mGuage.Min)) - 1
        Case Else
            nMark = -1
        End Select

        nMarkDiff = mGuage.Value - mGuage.MarkValue
        If nMarkDiff <> 0 Then
            sMarkTitle = IIf(nMarkDiff > 0, "+", "") & Trim$(CStr(nMarkDiff))
            wMarkFormat = DT_SINGLELINE + DT_NOPREFIX + DT_VCENTER
            nOffset = 0
            '
            Call DrawText(UserControl.hDC, sMarkTitle, Len(sMarkTitle), tText, wMarkFormat + DT_CALCRECT)
            nOffset = ((tRect.Bottom - tText.Bottom) \ 2&) - 1
            tMarkText = tRect
            If nOffset > 0 Then Call InflateRect(tMarkText, -nOffset, 0&)
            If mGuage.Alignment <> vbRightJustify Then tMarkText.Left = tMarkText.Right - tText.Right
        End If
    End If

    ' Remember the DC handle
    hWinDC = UserControl.hDC

    ' Calculate text position and format
    nOffset = 0
    wFormat = DT_SINGLELINE + DT_NOPREFIX + DT_VCENTER
    '
    Select Case mGuage.Alignment
    Case vbRightJustify
        wFormat = wFormat + DT_RIGHT
    Case vbCenter
        wFormat = wFormat + DT_CENTER
    End Select
    '
    If mGuage.Alignment <> vbCenter Then
        Call DrawText(hWinDC, sTitle, Len(sTitle), tText, wFormat + DT_CALCRECT)
        nOffset = ((tRect.Bottom - tText.Bottom) \ 2&) - 1
    End If
    '
    tText = tRect
    If nOffset > 0 Then Call InflateRect(tText, -nOffset, 0&)

    nForeColor = UserControl.ForeColor

    ' Progress fill color
    If mGuage.Value = mGuage.Max Then UserControl.ForeColor = UserControl.FillColor
    nColor = (&HFFFFFF Xor GetColor(UserControl.ForeColor))

    UserControl.DrawMode = vbCopyPen

    ' Erase everything by redrawing the background
    hBrush = CreateSolidBrush(GetColor(UserControl.BackColor))
    Call FillRect(hWinDC, tRect, hBrush)
    Call DeleteObject(hBrush)

    ' Show the text
    Call DrawText(hWinDC, sTitle, Len(sTitle), tText, wFormat)

    If sMarkTitle <> "" Then Call DrawText(hWinDC, sMarkTitle, Len(sMarkTitle), tMarkText, wMarkFormat)

    ' Show the progress
    If nRight >= 0 Then
        UserControl.DrawMode = vbXorPen                         ' XOr the pen
        UserControl.Line (0, -1)-(nRight, nHeight), nColor, BF
    End If

    If mGuage.UseMark Then
        UserControl.DrawMode = vbCopyPen
        UserControl.Line (nMark, -1)-(nMark, nHeight), GetColor(mGuage.MarkColor), BF
    End If

    UserControl.ForeColor = nForeColor
End Sub

Private Function GetColor(ByVal nColor As Long) As Long
    Const SYSCOLOR_BIT As Long = &H80000000
    If (nColor And SYSCOLOR_BIT) = SYSCOLOR_BIT Then
        nColor = nColor And (Not SYSCOLOR_BIT)
        GetColor = GetSysColor(nColor)
    Else
        GetColor = nColor
    End If
End Function

' -----------------------------------------------------------------------------

Public Property Get BorderStyle() As BorderStyles
Attribute BorderStyle.VB_ProcData.VB_Invoke_Property = ";Appearance"
    BorderStyle = UserControl.BorderStyle
End Property
Public Property Let BorderStyle(ByVal NewValue As BorderStyles)
    UserControl.BorderStyle = NewValue
    PropertyChanged "BorderStyle"
End Property

Public Property Get BackColor() As Long
Attribute BackColor.VB_ProcData.VB_Invoke_Property = ";Appearance"
    BackColor = UserControl.BackColor
End Property
Public Property Let BackColor(ByVal NewValue As Long)
    UserControl.BackColor = NewValue
    PropertyChanged "BackColor"
    UserControl_Paint
End Property

Public Property Get ForeColor() As Long
Attribute ForeColor.VB_ProcData.VB_Invoke_Property = ";Appearance"
    ForeColor = UserControl.ForeColor
End Property
Public Property Let ForeColor(ByVal NewValue As Long)
    UserControl.ForeColor = NewValue
    PropertyChanged "ForeColor"
    UserControl_Paint
End Property

Public Property Get MaxColor() As Long
Attribute MaxColor.VB_ProcData.VB_Invoke_Property = ";Appearance"
    MaxColor = UserControl.FillColor
End Property
Public Property Let MaxColor(ByVal NewValue As Long)
    UserControl.FillColor = NewValue
    PropertyChanged "FillColor"
    UserControl_Paint
End Property

Public Property Get MarkColor() As Long
    MarkColor = mGuage.MarkColor
End Property
Public Property Let MarkColor(ByVal NewValue As Long)
    mGuage.MarkColor = NewValue
    PropertyChanged "MarkColor"
    If mGuage.UseMark Then UserControl_Paint
End Property

'Public Property Get Font() As StdFont
'    Set Font = UserControl.Font
'End Property
'Public Property Set Font(ByVal NewValue As StdFont)
'    Set UserControl.Font = NewValue
'    PropertyChanged "Font"
'    UserControl_Paint
'End Property

Public Property Get Alignment() As AlignmentConstants
Attribute Alignment.VB_ProcData.VB_Invoke_Property = ";Appearance"
    Alignment = mGuage.Alignment
End Property
Public Property Let Alignment(ByVal NewValue As AlignmentConstants)
    mGuage.Alignment = NewValue
    PropertyChanged "Alignment"
    UserControl_Paint
End Property

Public Property Get Suffix() As String
Attribute Suffix.VB_ProcData.VB_Invoke_Property = ";Text"
    Suffix = mGuage.Suffix
End Property
Public Property Let Suffix(ByVal NewValue As String)
    mGuage.Suffix = NewValue
    PropertyChanged "Suffix"
    UserControl_Paint
End Property

Public Property Get Min() As Long
Attribute Min.VB_ProcData.VB_Invoke_Property = ";Data"
    Min = mGuage.Min
End Property
Public Property Let Min(ByVal NewValue As Long)
    If NewValue > mGuage.Max Then Exit Property
    mGuage.Min = NewValue
    If mGuage.Value < mGuage.Min Then mGuage.Value = mGuage.Min
    PropertyChanged "Min"
    UserControl_Paint
End Property

Public Property Get Max() As Long
Attribute Max.VB_ProcData.VB_Invoke_Property = ";Data"
    Max = mGuage.Max
End Property
Public Property Let Max(ByVal NewValue As Long)
    If NewValue < mGuage.Min Then Exit Property
    mGuage.Max = NewValue
    If mGuage.Value > mGuage.Max Then mGuage.Value = mGuage.Max
    PropertyChanged "Max"
    UserControl_Paint
End Property

Public Property Get Value() As Long
Attribute Value.VB_ProcData.VB_Invoke_Property = ";Data"
Attribute Value.VB_UserMemId = 0
    Value = mGuage.Value
End Property
Public Property Let Value(ByVal NewValue As Long)
    If NewValue < mGuage.Min Then
        mGuage.Value = mGuage.Min
    ElseIf NewValue > mGuage.Max Then
        mGuage.Value = mGuage.Max
    Else
        mGuage.Value = NewValue
    End If
    PropertyChanged "Value"
    UserControl_Paint
End Property

Public Property Get MarkValue() As Long
    MarkValue = mGuage.MarkValue
End Property
Public Property Let MarkValue(ByVal NewValue As Long)
    mGuage.MarkValue = NewValue
    PropertyChanged "MarkValue"
    If mGuage.UseMark Then UserControl_Paint
End Property

Public Property Get UseMark() As Boolean
    UseMark = mGuage.UseMark
End Property
Public Property Let UseMark(ByVal NewValue As Boolean)
    mGuage.UseMark = NewValue
    PropertyChanged "UseMark"
    UserControl_Paint
End Property
