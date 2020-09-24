Attribute VB_Name = "modFx"
'##############################################
'#          Coded by Walter A. Narvasa        #
'#       POS2000 - Point of Sales System      #
'#                                            #
'#           area :  modFX                    #
'#    description :  Screen & GUI Effects     #
'#         e-mail :  walter@wancom.8k.com     #
'#            url :  http://wancom.8k.com     #
'#                                            #
'##############################################

'EXIT WINDOWS SETTINGS
Public Declare Function ExitWindows Lib "User32" (ByVal dwReturnCode As Long, ByVal uReserved As Integer) As Integer
Global Const EW_REBOOTSYSTEM = &H43
Global Const EW_RESTARTWINDOWS = &H42
Global Const EW_EXITWINDOWS = 0

' TO DISABLE/ENABLE CTRL-ALT-DELETE
Public Declare Function SystemParametersInfo Lib "User32" Alias "SystemParametersInfoA" (ByVal uAction As Long, ByVal uParam As Long, ByVal lpvParam As Any, ByVal fuWinIni As Long) As Long

' FOR INI SETTINGS
#If Win16 Then
   Declare Function WritePrivateProfileString Lib "Kernel" (ByVal AppName As String, ByVal Keyname As String, ByVal NewString As String, ByVal filename As String) As Integer
   Declare Function GetPrivateProfileString Lib "Kernel" Alias "GetPrivateProfilestring" (ByVal AppName As String, ByVal Keyname As Any, ByVal default As String, ByVal ReturnedString As String, ByVal MAXSIZE As Integer, ByVal filename As String) As Integer
#Else
   Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As String, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
   Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As String, ByVal lpString As Any, ByVal lpFileName As String) As Long
#End If

' SET FORM ON TOP Declare our API functions
Declare Function SetWindowPos Lib "User32" ( _
                                ByVal hwnd As Long, _
                                ByVal hWndInsertAfter As Long, _
                                ByVal X As Long, ByVal Y As Long, _
                                ByVal cx As Long, ByVal cy As Long, _
                                ByVal wFlags As Long) As Long

' FOR RESOLUTION CHANGER
Private Declare Function EnumDisplaySettings Lib "User32" Alias "EnumDisplaySettingsA" (ByVal lpszDeviceName As Long, ByVal iModeNum As Long, lpDevMode As Any) As Boolean
Public Declare Function ChangeDisplaySettings Lib "User32" Alias "ChangeDisplaySettingsA" (lpDevMode As Any, ByVal dwflags As Long) As Long
    Private Const CCDEVICENAME = 32
    Private Const CCFORMNAME = 32
    Private Const DM_BITSPERPEL = &H60000
    Private Const DM_PELSWIDTH = &H80000
    Private Const DM_PELSHEIGHT = &H100000
Public Type DEVMODE
    dmDeviceName As String * CCDEVICENAME
    dmSpecVersion As Integer
    dmDriverVersion As Integer
    dmSize As Integer
    dmDriverExtra As Integer
    dmFields As Long
    dmOrientation As Integer
    dmPaperSize As Integer
    dmPaperLength As Integer
    dmPaperWidth As Integer
    dmScale As Integer
    dmCopies As Integer
    dmDefaultSource As Integer
    dmPrintQuality As Integer
    dmColor As Integer
    dmDuplex As Integer
    dmYResolution As Integer
    dmTTOption As Integer
    dmCollate As Integer
    dmFormName As String * CCFORMNAME
    dmUnusedPadding As Integer
    dmBitsPerPel As Integer
    dmPelsWidth As Long
    dmPelsHeight As Long
    dmDisplayFlags As Long
    dmDisplayFrequency As Long
    End Type

' FOR SYSTEM TRAY ICON
Public Declare Function Shell_NotifyIcon Lib "shell32" Alias "Shell_NotifyIconA" (ByVal dwMessage As Long, pnid As NOTIFYICONDATA) As Boolean
Public Declare Function SetForegroundWindow Lib "User32" (ByVal hwnd As Long) As Long
Public Const HWND_TOPMOST = -1
Public Const HWND_NOTOPMOST = -2
Public Const NIM_ADD = &H0
Public Const NIM_MODIFY = &H1
Public Const NIM_DELETE = &H2
Public Const NIF_MESSAGE = &H1
Public Const NIF_ICON = &H2
Public Const NIF_TIP = &H4
Public Const WM_MOUSEMOVE = &H200
Public Const WM_LBUTTONDOWN = &H201
Public Const WM_LBUTTONUP = &H202
Public Const WM_LBUTTONDBLCLK = &H203
Public Const WM_RBUTTONDOWN = &H204
Public Const WM_RBUTTONUP = &H205
Public Const WM_RBUTTONDBLCLK = &H206
Public Const SWP_NOMOVE = &H2
Public Const SWP_NOSIZE = &H1
Public Const flags = SWP_NOMOVE Or SWP_NOSIZE
Public nid As NOTIFYICONDATA
Dim hKey As Long, MainKeyHandle As Long
Dim rtn As Long, lBuffer As Long, sBuffer As String
Dim lBufferSize As Long
Public Type POINTAPI
        X As Long
        Y As Long
End Type
Public Type NOTIFYICONDATA
        cbSize As Long
        hwnd As Long
        uId As Long
        uFlags As Long
        uCallBackMessage As Long
        hIcon As Long
        szTip As String * 64
End Type
Const DisplayErrorMsg = False

' DISABLE RIGHT MOUSE CLICK
Public Declare Function SetWindowsHookEx Lib "User32" Alias "SetWindowsHookExA" (ByVal idHook As Long, ByVal lpfn As Long, ByVal hmod As Long, ByVal dwThreadId As Long) As Long
Public Declare Function UnhookWindowsHookEx Lib "User32" (ByVal hHook As Long) As Long
Public Declare Function CallNextHookEx Lib "User32" (ByVal hHook As Long, ByVal nCode As Long, ByVal wParam As Long, lparam As Any) As Long
Public Declare Function GetCurrentThreadId Lib "kernel32" () As Long
Public Declare Function GetClassName Lib "User32" Alias "GetClassNameA" (ByVal hwnd As Long, ByVal lpClassName As String, ByVal nMaxCount As Long) As Long
'Public Const WM_RBUTTONUP = &H205
'Public Const WH_MOUSE = 7
'Type POINTAPI
'    x As Long
'    y As Long
'End Type
Type MOUSEHOOKSTRUCT
    pt As POINTAPI
    hwnd As Long
    wHitTestCode As Long
    dwExtraInfo As Long
    End Type
Public l_hMouseHook As Long

' Drag Form Declaration
Public Declare Function ReleaseCapture Lib "User32" () As Long
Public Declare Function SendMessage Lib "User32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lparam As Any) As Long
Public Const WM_NCLBUTTONDOWN = &HA1

'BITBIT FUNCTION & declare SRCCOPY
Public Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Public Const SRCCOPY = &HCC0020

' MACOS Titlebar
Public Function CreateMacOSTitleBar(pict As PictureBox, title As String)
    pict.FontTransparent = False
    pict.AutoRedraw = True
    pict.ScaleMode = 3
    pict.BackColor = &HCCCCCC
    pict.BorderStyle = 0
    pict.ForeColor = QBColor(0)
    pict.Font = "Chicago"
    pict.FontBold = False
    pict.FontSize = 10
    If (pict.ScaleWidth / 2) - (pict.TextWidth(title) / 2) <= 8 Then title = ""
    If title = "" Then
        lhs_left = 8
        lhs_right = pict.ScaleWidth - 8
        l_top = pict.ScaleHeight / 2 - 6
        dorhs = False
            dolhs = True
                GoTo drawit
            End If
            l_top = pict.ScaleHeight / 2 - 6
            lhs_left = 8
            sc = pict.ScaleWidth
            lhs_right = ((sc / 2) - (pict.TextWidth(title) / 2)) - 4
            lhs_right = Int(lhs_right)
            rhs_left = ((sc / 2) + (pict.TextWidth(title) / 2)) + 4
            rhs_left = Int(rhs_left)
            rhs_right = pict.ScaleWidth - 8
            dolhs = True
                dorhs = True
drawit:
                    If dolhs = True Then
                        For X = l_top To l_top + 10 Step 2
                            pict.Line (lhs_left - 1, X)-(lhs_right, X), &HFFFFFF
                            pict.Line (lhs_left, X + 1.5)-(lhs_right + 1, X + 1.5), &H666666
                        Next X
                    End If
                    If dorhs = True Then
                        For X = l_top To l_top + 10 Step 2
                            pict.Line (rhs_left - 1, X)-(rhs_right, X), &HFFFFFF
                            pict.Line (rhs_left, X + 1.5)-(rhs_right + 1, X + 1.5), &H666666
                        Next X
                    End If
                    pict.Line (0, pict.ScaleHeight - 1)-(pict.ScaleWidth, pict.ScaleHeight - 1), &H666666
                    maclefttext = (pict.ScaleWidth / 2) - (pict.TextWidth(title) / 2)
                    pict.CurrentX = maclefttext
                    mactoptext = (pict.ScaleHeight / 2) - (pict.TextHeight(title) / 2)
                    pict.CurrentY = mactoptext
                    pict.Print title
End Function

' MACOS Button
Function MacButton(xCaption As String, xDestination As PictureBox, _
                    nTop As Integer, nLeft As Integer, nWidth As Integer, _
                    nHeight As Integer, xSource As PictureBox, _
                    xTop As Integer, xLeft As Integer, xType As Integer)
    Call BitBlt(xDestination.hDC, nTop, nLeft, nWidth, nHeight, xSource.hDC, xTop, xLeft, SRCCOPY)
    xDestination.Refresh
    xDestination.FontBold = False
    xDestination.FontSize = 9
    xDestination.CurrentX = 117
    If xType = 1 Then
        xDestination.Font = "System"
        xDestination.CurrentY = 159
    ElseIf xType = 2 Then
        xDestination.Font = "System"
        xDestination.CurrentY = 112
    ElseIf xType = 3 Then
        xDestination.Font = "Wingdings 3"
        xDestination.CurrentY = 40
    End If
    xDestination.Print xCaption
End Function

' EXIT WINDOWS FUNCTION
Function DoExitWindows()
    On Error Resume Next
    Dim RetVal As Integer
    RetVal = ExitWindows(EW_EXITWINDOWS, 0)
End Function

' Drag Form Function
Function DragForm(frm As Form)
  Dim ret As Long
  ret = ReleaseCapture()
  ret = SendMessage(frm.hwnd, WM_NCLBUTTONDOWN, 2&, 0&)
End Function

' 3D FORM SETTINGS
Function ColForm(Obj As Object, r%, G%, B%, Step%)
    Dim R1%, G1%, B1%, R2%, G2%, B2%
    Obj.ScaleMode = 3
    Obj.AutoRedraw = True
    Obj.BackColor = RGB(r%, G%, B%)
    R1% = r% + Step%: If R1% > 255 Then R1% = 255
    G1% = G% + Step%: If G1% > 255 Then G1% = 255
    B1% = B% + Step%: If B1% > 255 Then B1% = 255
    R2% = r% - Step%: If R2% < 0 Then R2% = 0
    G2% = G% - Step%: If G2% < 0 Then G2% = 0
    B2% = B% - Step%: If B2% < 0 Then B2% = 0
    Obj.Line (2, 2)-(Obj.ScaleWidth - 2, Obj.ScaleHeight - 2), RGB(R1%, G1%, B1%), B
    Obj.Line (Obj.ScaleWidth - 2, 2)-(Obj.ScaleWidth - 2, Obj.ScaleHeight - 1), RGB(R2%, G2%, B2%)
    Obj.Line (1, Obj.ScaleHeight - 2)-(Obj.ScaleWidth - 2, Obj.ScaleHeight - 2), RGB(R2%, G2%, B2%)
    Obj.Line (5, 5)-(Obj.ScaleWidth - 5, Obj.ScaleHeight - 5), RGB(R2%, G2%, B2%), B
    Obj.Line (Obj.ScaleWidth - 5, 6)-(Obj.ScaleWidth - 5, Obj.ScaleHeight - 4), RGB(R1%, G1%, B1%)
    Obj.Line (5, Obj.ScaleHeight - 5)-(Obj.ScaleWidth - 5, Obj.ScaleHeight - 5), RGB(R1%, G1%, B1%)
End Function

' SET FORM ON TOP
Public Sub SetFormOnTop(myForm As Object)
     SetWindowPos myForm.hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE
End Sub

' FOR EXIT GUI EFFECTS
Public Sub ExitFx(frm As Form)
    On Error Resume Next
    Dim GotoVal, GoInto
    GotoVal = frm.Height / 2
    For GoInto = 1 To GotoVal
        DoEvents
        frm.Height = frm.Height - 100
        frm.Top = (Screen.Height - frm.Height) \ 2
        If frm.Height <= 500 Then Exit For
            Next GoInto
horiz:
    frm.Height = 100
    GotoVal = frm.Width / 2
        For GoInto = 1 To GotoVal
            DoEvents
            frm.Width = frm.Width - 100
            frm.Left = (Screen.Width - frm.Width) \ 2
        If frm.Width <= 20 Then Exit For
            Next GoInto
End Sub

' FOR UNLOADING ALL FORMS
Public Sub UnloadAllForms(Optional sFormName As String = "")
    Dim Form As Form
    For Each Form In Forms
        If Form.Name <> sFormName Then
            Unload Form
            Set Form = Nothing
        End If
    Next Form
End Sub

' FOR MINIMIZING ALL FORMS
Public Sub MinimizeAllForms()
    Dim objTemp As Object
        For Each objTemp In Forms
            objTemp.WindowState = 1
        Next
End Sub

' FOR HIDING CHILD FORMS
Public Sub HideAllForms()
    Dim objTemp As Object
        For Each objTemp In Forms
            objTemp.Hide
        Next
End Sub
' SYNTAX: HideChildForms

' FOR INI SETTINGS
Function ReadINI(Section, Keyname, filename As String) As String
    Dim sRet As String
    sRet = String(255, Chr(0))
    ReadINI = Left(sRet, GetPrivateProfileString(Section, ByVal Keyname, "", sRet, Len(sRet), filename))
End Function
Function WriteINI(sSection As String, sKeyName As String, sNewString As String, sFileName) As Integer
    Dim r
    r = WritePrivateProfileString(sSection, sKeyName, sNewString, sFileName)
End Function

' FOR RESOLUTION VERIFIER
Function IsResolution(Width As Integer, Height As Integer) As Boolean
    If (Screen.Width / Screen.TwipsPerPixelX = Width) And (Screen.Height / Screen.TwipsPerPixelY = Height) Then
        IsResolution = True
    Else
        IsResolution = False
    End If
End Function

' FOR RESOLUTION CHANGER
Function ChangeRes(Width As Single, Height As Single, BPP As Integer) As Integer
    On Error GoTo ERROR_HANDLER
    Dim DevM As DEVMODE, i As Integer, ReturnVal As Boolean, _
    RetValue, OldWidth As Single, OldHeight As Single, _
    OldBPP As Integer
    Call EnumDisplaySettings(0&, -1, DevM)
    OldWidth = DevM.dmPelsWidth
    OldHeight = DevM.dmPelsHeight
    OldBPP = DevM.dmBitsPerPel
    i = 0
    Do
        ReturnVal = EnumDisplaySettings(0&, i, DevM)
        i = i + 1
    Loop Until (ReturnVal = False)
    DevM.dmFields = DM_PELSWIDTH Or DM_PELSHEIGHT Or DM_BITSPERPEL
    DevM.dmPelsWidth = Width
    DevM.dmPelsHeight = Height
    DevM.dmBitsPerPel = BPP
    Call ChangeDisplaySettings(DevM, 1)
    If RetValue = vbCancel Then
        DevM.dmPelsWidth = OldWidth
        DevM.dmPelsHeight = OldHeight
        DevM.dmBitsPerPel = OldBPP
        Call ChangeDisplaySettings(DevM, 1)
        ChangeRes = 0
    Else
        ChangeRes = 1
    End If
    Exit Function
ERROR_HANDLER:
    ChangeRes = 0
End Function

' TO DISABLE/ENABLE CTRL-ALT-DELETE
Function DisableCtrlAltDelete(bDisabled As Boolean)
    Dim X As Long
    X = SystemParametersInfo(97, bDisabled, CStr(1), 0)
End Function
        
' DISABLE RIGHT MOUSE CLICK
Public Function MouseHookProc(ByVal nCode As Long, ByVal wParam As Long, mhs As MOUSEHOOKSTRUCT) As Long
    If (nCode >= 0 And wParam = WM_RBUTTONUP) Then
        Dim sClassName As String
        Dim sTestClass As String
        sTestClass = "HTML_Internet Explorer"
        sClassName = String$(256, 0)
        If GetClassName(mhs.hwnd, sClassName, Len(sClassName)) > 0 Then
            If Left$(sClassName, Len(sTestClass)) = sTestClass Then
                MouseHookProc = 1
                Exit Function
            End If
        End If
    End If
    MouseHookProc = CallNextHookEx(l_hMouseHook, nCode, wParam, mhs)
End Function
Public Sub BeginRightMouseTrap()
    l_hMouseHook = SetWindowsHookEx(WH_MOUSE, AddressOf MouseHookProc, App.hInstance, GetCurrentThreadId)
End Sub
Public Sub EndRightMouseTrap()
    UnhookWindowsHookEx l_hMouseHook
End Sub
