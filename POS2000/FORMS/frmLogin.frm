VERSION 5.00
Begin VB.Form frmLogin 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00CCCCCC&
   BorderStyle     =   0  'None
   Caption         =   "frmLogin"
   ClientHeight    =   2985
   ClientLeft      =   3495
   ClientTop       =   3105
   ClientWidth     =   5565
   Icon            =   "frmLogin.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2985
   ScaleWidth      =   5565
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox BoxContainer 
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   0  'None
      FillColor       =   &H00C0C0C0&
      Height          =   2890
      Left            =   40
      ScaleHeight     =   2895
      ScaleWidth      =   5460
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   40
      Width           =   5460
      Begin VB.PictureBox titleBar 
         BorderStyle     =   0  'None
         Height          =   375
         Left            =   120
         ScaleHeight     =   375
         ScaleWidth      =   5205
         TabIndex        =   6
         TabStop         =   0   'False
         Top             =   90
         Width           =   5200
         Begin VB.PictureBox Help 
            Appearance      =   0  'Flat
            AutoRedraw      =   -1  'True
            BackColor       =   &H00CCCCCC&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   280
            Left            =   4920
            MouseIcon       =   "frmLogin.frx":030A
            MousePointer    =   99  'Custom
            ScaleHeight     =   285
            ScaleWidth      =   315
            TabIndex        =   8
            TabStop         =   0   'False
            Top             =   60
            Width           =   313
         End
         Begin VB.PictureBox Closed 
            Appearance      =   0  'Flat
            AutoRedraw      =   -1  'True
            BackColor       =   &H00CCCCCC&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   280
            Left            =   0
            MouseIcon       =   "frmLogin.frx":0614
            MousePointer    =   99  'Custom
            ScaleHeight     =   285
            ScaleWidth      =   315
            TabIndex        =   7
            TabStop         =   0   'False
            Top             =   60
            Width           =   315
         End
      End
      Begin VB.TextBox txtUserName 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   2040
         TabIndex        =   0
         Top             =   960
         Width           =   2415
      End
      Begin VB.TextBox txtPassword 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         IMEMode         =   3  'DISABLE
         Left            =   2040
         PasswordChar    =   "*"
         TabIndex        =   1
         Top             =   1320
         Width           =   2415
      End
      Begin VB.PictureBox cmdOk 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H00CCCCCC&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   570
         Left            =   1560
         MouseIcon       =   "frmLogin.frx":091E
         MousePointer    =   99  'Custom
         ScaleHeight     =   570
         ScaleWidth      =   1095
         TabIndex        =   2
         TabStop         =   0   'False
         Top             =   2040
         Width           =   1100
      End
      Begin VB.PictureBox cmdCancel 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H00CCCCCC&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   570
         Left            =   2880
         MouseIcon       =   "frmLogin.frx":0C28
         MousePointer    =   99  'Custom
         ScaleHeight     =   570
         ScaleWidth      =   1095
         TabIndex        =   3
         TabStop         =   0   'False
         Top             =   2040
         Width           =   1100
      End
      Begin VB.Frame Frame1 
         BackColor       =   &H00000000&
         Caption         =   "Expire Information"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   2775
         Left            =   1680
         TabIndex        =   11
         Top             =   600
         Visible         =   0   'False
         Width           =   3525
         Begin VB.Timer Timer1 
            Interval        =   1000
            Left            =   2160
            Top             =   1680
         End
         Begin VB.Label Label7 
            BackStyle       =   0  'Transparent
            ForeColor       =   &H0000FF00&
            Height          =   375
            Left            =   2640
            TabIndex        =   28
            Top             =   1440
            Width           =   615
         End
         Begin VB.Label lblToday 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            ForeColor       =   &H0000FF00&
            Height          =   195
            Left            =   1560
            TabIndex        =   27
            Top             =   2040
            Width           =   45
         End
         Begin VB.Label lblE 
            BackStyle       =   0  'Transparent
            Caption         =   "Date Today:"
            ForeColor       =   &H0000FF00&
            Height          =   255
            Left            =   120
            TabIndex        =   26
            Top             =   2040
            Width           =   975
         End
         Begin VB.Label lblLeft 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            ForeColor       =   &H0000FF00&
            Height          =   195
            Left            =   2520
            TabIndex        =   25
            Top             =   0
            Width           =   765
         End
         Begin VB.Label lblExpired 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            ForeColor       =   &H0000FF00&
            Height          =   195
            Left            =   1560
            TabIndex        =   24
            Top             =   1680
            Width           =   45
         End
         Begin VB.Label lblTrial 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            ForeColor       =   &H0000FF00&
            Height          =   195
            Left            =   1560
            TabIndex        =   23
            Top             =   1320
            Width           =   45
         End
         Begin VB.Label lblF 
            BackStyle       =   0  'Transparent
            Caption         =   "Days Left:"
            ForeColor       =   &H0000FF00&
            Height          =   255
            Left            =   120
            TabIndex        =   22
            Top             =   2400
            Width           =   855
         End
         Begin VB.Label lblD 
            BackStyle       =   0  'Transparent
            Caption         =   "Trial Expired:"
            ForeColor       =   &H0000FF00&
            Height          =   255
            Left            =   120
            TabIndex        =   21
            Top             =   1680
            Width           =   975
         End
         Begin VB.Label lblC 
            BackStyle       =   0  'Transparent
            Caption         =   "Begin Trial:"
            ForeColor       =   &H0000FF00&
            Height          =   255
            Left            =   120
            TabIndex        =   20
            Top             =   1320
            Width           =   855
         End
         Begin VB.Label lblB 
            BackStyle       =   0  'Transparent
            Caption         =   "App Used:"
            ForeColor       =   &H0000FF00&
            Height          =   255
            Left            =   120
            TabIndex        =   19
            Top             =   960
            Width           =   855
         End
         Begin VB.Label lblTimes 
            BackStyle       =   0  'Transparent
            ForeColor       =   &H0000FF00&
            Height          =   255
            Left            =   1560
            TabIndex        =   18
            Top             =   960
            Width           =   1575
         End
         Begin VB.Label lblA 
            BackStyle       =   0  'Transparent
            Caption         =   "App Started at:"
            ForeColor       =   &H0000FF00&
            Height          =   255
            Left            =   120
            TabIndex        =   17
            Top             =   600
            Width           =   1335
         End
         Begin VB.Label lblStart 
            BackStyle       =   0  'Transparent
            ForeColor       =   &H0000FF00&
            Height          =   375
            Left            =   1560
            TabIndex        =   16
            Top             =   600
            Width           =   1815
         End
         Begin VB.Label Label5 
            BackStyle       =   0  'Transparent
            Caption         =   "Times"
            ForeColor       =   &H0000FF00&
            Height          =   255
            Left            =   2280
            TabIndex        =   15
            Top             =   960
            Width           =   615
         End
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   "This Trial Version will shutdown in:"
            ForeColor       =   &H0000FF00&
            Height          =   255
            Left            =   90
            TabIndex        =   14
            Top             =   240
            Width           =   2415
         End
         Begin VB.Label Label2 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "120"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H0000FFFF&
            Height          =   255
            Left            =   2520
            TabIndex        =   13
            Top             =   240
            Width           =   375
         End
         Begin VB.Label Label3 
            BackStyle       =   0  'Transparent
            Caption         =   "Secs."
            ForeColor       =   &H0000FF00&
            Height          =   255
            Left            =   2400
            TabIndex        =   12
            Top             =   480
            Width           =   630
         End
      End
      Begin VB.Label UserName 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "User Name"
         BeginProperty Font 
            Name            =   "Chicago"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   870
         TabIndex        =   10
         Top             =   960
         Width           =   1050
      End
      Begin VB.Label Password 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Password"
         BeginProperty Font 
            Name            =   "Chicago"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   960
         TabIndex        =   9
         Top             =   1320
         Width           =   960
      End
   End
   Begin VB.PictureBox Source 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00CCCCCC&
      BorderStyle     =   0  'None
      Height          =   4530
      Left            =   120
      Picture         =   "frmLogin.frx":0F32
      ScaleHeight     =   4411.822
      ScaleMode       =   0  'User
      ScaleWidth      =   5820
      TabIndex        =   4
      Top             =   3360
      Visible         =   0   'False
      Width           =   5820
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   2
      Height          =   2950
      Left            =   15
      Top             =   15
      Width           =   5520
   End
End
Attribute VB_Name = "frmLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'##############################################
'#          Coded by Walter A. Narvasa        #
'#       POS2000 - Point of Sales System      #
'#                                            #
'#           area :  frmLogin                 #
'#    description :  Password Login           #
'#         e-mail :  walter@wancom.8k.com     #
'#            url :  http://wancom.8k.com     #
'#                                            #
'##############################################

Option Explicit

Public db As DAO.Database
Dim dummy As DAO.Recordset
Dim dummy2 As DAO.Recordset
Dim ctr As Integer
Private Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Private Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Private Declare Function ShowWindow& Lib "user32" (ByVal hwnd&, ByVal nCmdShow&)
Private Const SW_HIDE = 0
Private Const SW_NORMAL = 1
Const SWP_HIDEWINDOW = &H80
Const SWP_SHOWWINDOW = &H40
Dim hDesk As Long
Dim Thwnd As Long

Private Sub Closed_Click()
    Me.Hide
    UnloadAllForms
    End
End Sub

Private Sub Closed_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call BitBlt(Closed.hDC, 0, 0, 73, 50, Source.hDC, 18, 107, SRCCOPY)
    Closed.Refresh
End Sub

Private Sub Closed_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call BitBlt(Closed.hDC, 0, 0, 73, 50, Source.hDC, 0, 107, SRCCOPY)
    Closed.Refresh
End Sub

Private Sub cmdCancel_Click()
    Me.Hide
    'UnloadAllForms
    Call DoExitWindows
End Sub

Private Sub cmdCancel_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call MacButton(" Cancel", cmdCancel, 0, 0, 73, 50, Source, 74, 0, 1)
End Sub

Private Sub cmdCancel_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call MacButton(" Cancel", cmdCancel, 0, 0, 73, 50, Source, 0, 0, 1)
End Sub

Private Sub cmdOk_Click()
    Dim strs As String
    If Get_User(txtUserName, txtPassword) Then
        'frmLogin.db.Close 'DISABLED TO ENABLED MULTI-LOG IN
        Me.Hide
        frmMain.Show
        frmMain.Company_Name = IIf(dummy2("COMPANY_NAME") = "", "DEMO COPY", Trim(dummy2("COMPANY_NAME")))
        Today = Now
        frmMain.StatusMessage = " Current User: " + txtUserName + _
                                "           " + Format(Today, "dddd " & "dd " & "mmmm " & "yyyy ")
        frmMain.MenuList.SetFocus
    Else
        ctr = ctr + 1
        If ctr = 4 Then
           End
        Else
           Call MessageBox("frmLogin", "Invalid User!!!!   Try Again....  You have" + str(4 - ctr) + " tries left", 0)
           SendKeys "{Home}+{End}"
        End If
   End If
End Sub

Private Sub cmdOk_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call MacButton("     Ok", cmdOk, 0, 0, 73, 50, Source, 74, 0, 1)
End Sub

Private Sub cmdOk_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call MacButton("     Ok", cmdOk, 0, 0, 73, 50, Source, 0, 0, 1)
End Sub

Private Sub Form_Load()
    On Error Resume Next
    Dim Thwnd As Long
    Dim RetValue
    RetValue = ChangeRes(800, 600, 32)
    Call CreateMacOSTitleBar(titleBar, " System Login ")
    Call MacButton("     Ok", cmdOk, 0, 0, 73, 50, Source, 0, 0, 1)
    Call MacButton(" Cancel", cmdCancel, 0, 0, 73, 50, Source, 0, 0, 1)
    Call BitBlt(Help.hDC, 0, 0, 73, 50, Source.hDC, 0, 90, SRCCOPY)
    Help.Refresh
    Call BitBlt(Closed.hDC, 0, 0, 73, 50, Source.hDC, 0, 107, SRCCOPY)
    Closed.Refresh
    Call ColForm(BoxContainer, 217, 211, 213, 125)
    frmWallpaper.Show
    KeyPreview = True
    Set db = Workspaces(0).OpenDatabase(App.Path & "\DATABASE\POSDATA.MDB")
    Set dummy2 = frmLogin.db.OpenRecordset("select * from SETUP order by COMPANY_NAME")
    If dummy2.RecordCount = 0 Then
        dummy2.AddNew
        dummy2("COMPANY_NAME") = ""
        dummy2.Update
    Else
        frmWallpaper.Company_Name = Trim(dummy2("COMPANY_NAME"))
        frmAbout.Company_Name = Trim(dummy2("COMPANY_NAME"))
    End If
    If dummy2("OPTION_ALAS") = True Then 'AUTO-LOADING AT STARTUP
        FileCopy App.Path & "\POS2000.LNK", Mid$(App.Path, 1, 3) & "WINDOWS\START MENU\PROGRAMS\STARTUP\POS2000.LNK"
    Else
        Kill Mid$(App.Path, 1, 3) & "WINDOWS\START MENU\PROGRAMS\STARTUP\POS2000.LNK"
    End If
    If dummy2("OPTION_HWT") = True Then 'HIDE WINDOWS TASKBAR
        Thwnd = FindWindow("Shell_traywnd", "")
        Call SetWindowPos(Thwnd, 0, 0, 0, 0, 0, SWP_HIDEWINDOW)
    Else
        Thwnd = FindWindow("Shell_traywnd", "")
        Call SetWindowPos(Thwnd, 0, 0, 0, 0, 0, SWP_SHOWWINDOW)
    End If
    If dummy2("OPTION_HDI") = True Then 'HIDE DESKTOP ICONS
        hDesk = FindWindow("progman", vbNullString)
        Call ShowWindow(hDesk, SW_HIDE)
    Else
        hDesk = FindWindow("progman", vbNullString)
        Call ShowWindow(hDesk, SW_NORMAL)
    End If
    If dummy2("OPTION_E3DT") = True Then 'ENABLE 30 DAYS TRIALER
        Call frmLogin.TrialerActivation
    End If
    If dummy2("OPTION_DRCM") = True Then 'DISABLE RIGHT-CLICK MOUSE
        BeginRightMouseTrap
    Else
        EndRightMouseTrap
    End If
    If dummy2("OPTION_DCADATS") = True Then 'DISABLE CTRL-ALT-DELETE/ALT-TAB SHORTCUT
        Call DisableCtrlAltDelete(True)
    Else
        Call DisableCtrlAltDelete(False)
    End If
    If dummy2("OPTION_DPW") = True Then 'DISABLE POS2000 WALLPAPER
        frmWallpaper.Hide
    Else
        frmWallpaper.Show
    End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyEscape:
                Me.Hide
                UnloadAllForms
                End
    End Select
    If (Shift = vbAltMask) Then
        Select Case KeyCode
            Case vbKeyF4
                    KeyCode = 0
        End Select
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Error Resume Next
    If dummy2("OPTION_E3DT") = True Then 'ENABLE 30 DAYS TRIALER
        SetStringValue "HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\WANCOM SYSTEMS", "Todays Date", frmLogin.lblStart.Caption
    Else
    End If
End Sub

Private Sub Help_Click()
'
End Sub

Private Sub Help_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call BitBlt(Help.hDC, 0, 0, 73, 50, Source.hDC, 19, 90, SRCCOPY)
    Help.Refresh
End Sub

Private Sub Help_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call BitBlt(Help.hDC, 0, 0, 73, 50, Source.hDC, 0, 90, SRCCOPY)
    Help.Refresh
End Sub

Private Sub Timer1_Timer()
    On Error Resume Next
    Dim xCount
    If dummy2("OPTION_E3DT") = True Then 'ENABLE 30 DAYS TRIALER
        xCount = IIf(IsNull(dummy2("COMPANY_EXPIRYCOUNT")), "0", dummy2("COMPANY_EXPIRYCOUNT"))
        'BY NUMBER OF TIMES USED TIME BOMB
        If Int(lblTimes.Caption) >= Int(xCount) Then
            'Call MessageBox("frmLogin", "Your trial version is expired!", 0)
            'frmMessageBox.SetFocus
            Kill App.Path + "\POS2000.EXE"
            Kill App.Path + "\DATABASE\POSDATA.MDB"
            End
        End If
        'BY 30 DAYS TRIAL
        If lblLeft <= 0 Then
            'Call MessageBox("frmLogin", "Your trial version is expired!", 0)
            'frmMessageBox.SetFocus
            Kill App.Path + "\POS2000.EXE"
            Kill App.Path + "\DATABASE\POSDATA.MDB"
            End
        End If
    End If
End Sub

Function TrialerActivation()
    On Error Resume Next
    Label7.Caption = GetStringValue("HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion", "RegisteredOwner")
    If GetStringValue("HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\WANCOM SYSTEMS", "Times Opened") = "Error" Then
        CreateKey "HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\WANCOM SYSTEMS"
        SetStringValue "HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\WANCOM SYSTEMS", "Times Opened", "1"
    End If
    If GetStringValue("HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\WANCOM SYSTEMS", "Times Opened") = "" Then
        SetStringValue "HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\WANCOM SYSTEMS", "Times Opened", "1"
    End If
        lblTimes.Caption = GetStringValue("HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\WANCOM SYSTEMS", "Times Opened") & "  times"
    If GetStringValue("HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\WANCOM SYSTEMS", "Expire Date") = "Error" Then
        CreateKey "HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\WANCOM SYSTEMS"
        SetStringValue "HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\WANCOM SYSTEMS", "Expire Date", Now + 29
        SetStringValue "HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\WANCOM SYSTEMS", "Trial Start Date", Now
    End If
    If GetStringValue("HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\WANCOM SYSTEMS", "Expire Date") = "" Then
        SetStringValue "HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\WANCOM SYSTEMS", "Expire Date", Now + 29
        SetStringValue "HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\WANCOM SYSTEMS", "Trial Start Date", Now
    End If
        lblTrial.Caption = GetStringValue("HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\WANCOM SYSTEMS", "Trial Start Date")
        lblTimes.Caption = GetStringValue("HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\WANCOM SYSTEMS", "Times Opened")
        CreateKey "HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\WANCOM SYSTEMS"
        SetStringValue "HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\WANCOM SYSTEMS", "Times Opened", lblTimes.Caption + 1
        SetStringValue "HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\WANCOM SYSTEMS", "Todays Date", Now
        lblStart.Caption = GetStringValue("HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\WANCOM SYSTEMS", "Todays Date")
        SetStringValue "HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\WANCOM SYSTEMS", "Todays Date", lblStart.Caption
        lblExpired.Caption = GetStringValue("HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\WANCOM SYSTEMS", "Expire Date")
        lblLeft = Format$(days360(lblStart.Caption, lblExpired.Caption), "###,###") '& " Days Remaining"
        SetStringValue "HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\WANCOM SYSTEMS", "Days", lblLeft.Caption
        lblLeft.Caption = GetStringValue("HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\WANCOM SYSTEMS", "Days")
        SetStringValue "HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\WANCOM SYSTEMS", "Copyright", App.LegalCopyright
        SetStringValue "HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\WANCOM SYSTEMS", "Trade Mark", App.LegalTrademarks
        SetStringValue "HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\WANCOM SYSTEMS", "Version", App.Major & "." & App.Minor & "." & App.Revision
    If lblLeft = 1 & " Days Remaining" Then
        lblLeft = 1 & " Day Remaining"
    End If
    If Val(lblLeft.Caption) < 1 Then
        lblLeft = "Last Day"
    End If
    If Val(lblLeft.Caption) < 0 Then
        Call MessageBox("frmLogin", "Your trial version is expired!", 0)
        frmMessageBox.SetFocus
        End
    End If
    If Val(lblLeft.Caption) > 30 Then
        Call MessageBox("frmLogin", "Do not adjust Date/Time. Your trial version is expired!", 0)
        frmMessageBox.SetFocus
        End
    Else
        'ProgressBar1.Value = Val(lblLeft)
    End If
    Exit Function
End Function

Public Function days360(dt1 As Date, dt2 As Date) As Long
    Dim z1 As Long, z2 As Long
    Dim d1 As Long, d2 As Long
    Dim m1 As Long, m2 As Long
    Dim Y1 As Long, Y2 As Long
    d1 = Day(dt1)
    m1 = Month(dt1)
    Y1 = Year(dt1)
    d2 = Day(dt2)
    m2 = Month(dt2)
    Y2 = Year(dt2)
    If d1 = 31 Then
        z1 = 30
    Else
        z1 = d1
    End If
    If d2 = 31 And d1 >= 30 Then
        z2 = 30
    Else
        z2 = d2
    End If
    days360 = (Y2 - Y1) * 360 + (m2 - m1) * 30 + (z2 - z1)
End Function

Private Sub titleBar_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call DragForm(Me)
End Sub

Function Get_User(p_user As String, p_pass As String) As Boolean
    Dim strs As String
    strs = ""
    strs = "select * from USER_PASSWORD where USER_NAME = '" & p_user & "'" _
            & " and USER_PASSWORD = '" & Decode_Pass(p_pass) & "'"
    Set dummy = Me.db.OpenRecordset(strs)
    If dummy.AbsolutePosition <> -1 Then
        Get_User = True
        frmMain.MenuList.Clear
        If dummy("USER_ALLOW_SM") = True Then frmMain.MenuList.AddItem "Supplier Masterfile"
        If dummy("USER_ALLOW_PM") = True Then frmMain.MenuList.AddItem "Product Masterfile"
        If dummy("USER_ALLOW_CM") = True Then frmMain.MenuList.AddItem "Category Masterfile"
        If dummy("USER_ALLOW_ST") = True Then frmMain.MenuList.AddItem "Selling Transaction"
        If dummy("USER_ALLOW_RT") = True Then frmMain.MenuList.AddItem "Receiving Transaction"
        If dummy("USER_ALLOW_SRR") = True Then frmMain.MenuList.AddItem "Stocks Re-order Report"
        If dummy("USER_ALLOW_SHR") = True Then frmMain.MenuList.AddItem "Selling History Report"
        If dummy("USER_ALLOW_RHR") = True Then frmMain.MenuList.AddItem "Receiving History Report"
        If dummy("USER_ALLOW_SPSR") = True Then frmMain.MenuList.AddItem "Stocks Per Supplier Report"
        If dummy("USER_ALLOW_PLR") = True Then frmMain.MenuList.AddItem "Product Listing Report"
        If dummy("USER_ALLOW_SLR") = True Then frmMain.MenuList.AddItem "Supplier Listing Report"
        If dummy("USER_ALLOW_BRF") = True Then frmMain.MenuList.AddItem "Backup/Restore Files"
        If dummy("USER_ALLOW_PS") = True Then frmMain.MenuList.AddItem "Password Security"
        If dummy("USER_ALLOW_CFS") = True Then frmMain.MenuList.AddItem "Code File Setup"
        If dummy("USER_ALLOW_SS") = True Then frmMain.MenuList.AddItem "Software Setup"
    Else
         Get_User = False
    End If
End Function

Function Get_BirthDate() As Boolean
    On Error Resume Next
    Dim strs As String
    Dim i As String
    i = InputBox("Enter End-User Birthdate", "Message:")
    If i = "" Then i = "3/7/73"
        strs = "select * from USER_PASSWORD where  USER_BIRTHDATE = #" & i & "#"
        Set dummy = db.OpenRecordset(strs)
    If dummy.AbsolutePosition >= 0 Then
        Get_BirthDate = True
        txtUserName = dummy("USER_NAME")
        txtPassword = dummy("USER_PASSWORD")
    Else
        Get_BirthDate = False
   End If
End Function

Private Sub txtPassword_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then cmdOk_Click
End Sub
Private Sub txtUserName_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then txtPassword.SetFocus
End Sub

