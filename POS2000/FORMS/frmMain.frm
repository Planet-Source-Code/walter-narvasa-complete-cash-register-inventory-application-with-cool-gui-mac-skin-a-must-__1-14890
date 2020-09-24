VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "CRYSTL32.OCX"
Begin VB.Form frmMain 
   BackColor       =   &H00CCCCCC&
   BorderStyle     =   0  'None
   Caption         =   "frmMain"
   ClientHeight    =   9000
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   12000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   9000
   ScaleWidth      =   12000
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox BoxContainer 
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   0  'None
      FillColor       =   &H00C0C0C0&
      Height          =   8880
      Left            =   40
      ScaleHeight     =   8880
      ScaleWidth      =   11895
      TabIndex        =   0
      Top             =   40
      Width           =   11890
      Begin VB.PictureBox Wallpaper 
         BackColor       =   &H00C0C0C0&
         Height          =   7460
         Left            =   3360
         Picture         =   "frmMain.frx":0000
         ScaleHeight     =   7395
         ScaleWidth      =   8295
         TabIndex        =   13
         Top             =   600
         Width           =   8350
         Begin Project1.ProgressGuage guageSystem 
            Height          =   210
            Left            =   4560
            Top             =   6360
            Width           =   3300
            _ExtentX        =   5821
            _ExtentY        =   370
         End
         Begin Project1.ProgressGuage guageUser 
            Height          =   210
            Left            =   4560
            Top             =   6600
            Width           =   3300
            _ExtentX        =   5821
            _ExtentY        =   370
         End
         Begin Project1.ProgressGuage guageGDI 
            Height          =   210
            Left            =   4560
            Top             =   6840
            Width           =   3300
            _ExtentX        =   5821
            _ExtentY        =   370
         End
         Begin VB.Label Company_Name 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   36
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00404040&
            Height          =   1935
            Left            =   0
            TabIndex        =   41
            Top             =   3360
            Width           =   8280
         End
         Begin VB.Label lblFree 
            BackStyle       =   0  'Transparent
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00404040&
            Height          =   255
            Left            =   6600
            TabIndex        =   40
            Top             =   5760
            Visible         =   0   'False
            Width           =   855
         End
         Begin VB.Label lblTotal 
            BackStyle       =   0  'Transparent
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00404040&
            Height          =   255
            Left            =   6600
            TabIndex        =   39
            Top             =   6000
            Visible         =   0   'False
            Width           =   855
         End
         Begin VB.Label lblUsed 
            BackStyle       =   0  'Transparent
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00404040&
            Height          =   255
            Left            =   6600
            TabIndex        =   38
            Top             =   5520
            Visible         =   0   'False
            Width           =   855
         End
         Begin VB.Label txtDiskSize 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "n/a"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00404040&
            Height          =   225
            Left            =   5400
            TabIndex        =   37
            Top             =   6000
            Width           =   1095
         End
         Begin VB.Label txtFreeSpace 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "n/a"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00404040&
            Height          =   225
            Left            =   5400
            TabIndex        =   36
            Top             =   5760
            Width           =   1095
         End
         Begin VB.Label txtUsedKB 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "n/a"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00404040&
            Height          =   225
            Left            =   5400
            TabIndex        =   35
            Top             =   5520
            Width           =   1095
         End
         Begin VB.Label Label13 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Total Disk Space"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00404040&
            Height          =   210
            Left            =   3840
            TabIndex        =   34
            Top             =   6007
            Width           =   1545
         End
         Begin VB.Label Label12 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Free Disk Space"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00404040&
            Height          =   210
            Left            =   3915
            TabIndex        =   33
            Top             =   5767
            Width           =   1470
         End
         Begin VB.Label Label11 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Disk Space Used"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00404040&
            Height          =   210
            Left            =   3840
            TabIndex        =   32
            Top             =   5527
            Width           =   1545
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "n/a"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00404040&
            Height          =   225
            Index           =   3
            Left            =   2400
            TabIndex        =   31
            Top             =   6195
            Width           =   1095
         End
         Begin VB.Label Label10 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Available Virtual Memory:"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00404040&
            Height          =   225
            Left            =   480
            TabIndex        =   30
            Top             =   6195
            Width           =   1815
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "n/a"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00404040&
            Height          =   225
            Index           =   6
            Left            =   2400
            TabIndex        =   29
            Top             =   6855
            Width           =   1095
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "n/a"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00404040&
            Height          =   225
            Index           =   5
            Left            =   2400
            TabIndex        =   28
            Top             =   6645
            Width           =   1095
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "n/a"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00404040&
            Height          =   225
            Index           =   4
            Left            =   2400
            TabIndex        =   27
            Top             =   6420
            Width           =   1095
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "n/a"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00404040&
            Height          =   225
            Index           =   2
            Left            =   2400
            TabIndex        =   26
            Top             =   5970
            Width           =   1095
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "n/a"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00404040&
            Height          =   225
            Index           =   1
            Left            =   2400
            TabIndex        =   25
            Top             =   5745
            Width           =   1095
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "n/a"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00404040&
            Height          =   225
            Index           =   0
            Left            =   2400
            TabIndex        =   24
            Top             =   5520
            Width           =   1095
         End
         Begin VB.Label Label9 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Memory Load:"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00404040&
            Height          =   225
            Left            =   960
            TabIndex        =   23
            Top             =   6855
            Width           =   1335
         End
         Begin VB.Label Label8 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Available Page File:"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00404040&
            Height          =   225
            Left            =   360
            TabIndex        =   22
            Top             =   6645
            Width           =   1935
         End
         Begin VB.Label Label7 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Total Page File:"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00404040&
            Height          =   225
            Left            =   960
            TabIndex        =   21
            Top             =   6420
            Width           =   1335
         End
         Begin VB.Label Label6 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Total Virtual Memory:"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00404040&
            Height          =   225
            Left            =   480
            TabIndex        =   20
            Top             =   5970
            Width           =   1815
         End
         Begin VB.Label Label5 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Available Physical Memory:"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00404040&
            Height          =   225
            Left            =   360
            TabIndex        =   19
            Top             =   5745
            Width           =   1935
         End
         Begin VB.Label Label4 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Total Physical Memory:"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00404040&
            Height          =   225
            Left            =   600
            TabIndex        =   18
            Top             =   5520
            Width           =   1695
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "GDI"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   11.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00404040&
            Height          =   270
            Left            =   3960
            TabIndex        =   17
            Top             =   6825
            Width           =   450
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "User"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   11.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00404040&
            Height          =   270
            Left            =   3885
            TabIndex        =   16
            Top             =   6570
            Width           =   525
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "System"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   11.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00404040&
            Height          =   270
            Left            =   3615
            TabIndex        =   15
            Top             =   6330
            Width           =   795
         End
      End
      Begin VB.PictureBox Applets 
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   0  'None
         Height          =   495
         Left            =   3310
         ScaleHeight     =   495
         ScaleWidth      =   8385
         TabIndex        =   10
         Top             =   8160
         Width           =   8390
         Begin VB.Timer tmrUpdate 
            Interval        =   1000
            Left            =   3360
            Top             =   0
         End
         Begin VB.Label StatusMessage 
            BackStyle       =   0  'Transparent
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   240
            TabIndex        =   12
            Top             =   135
            Width           =   6735
         End
         Begin VB.Label Clock 
            BackStyle       =   0  'Transparent
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   6960
            TabIndex        =   11
            Top             =   140
            Width           =   1455
         End
      End
      Begin VB.PictureBox titleBar 
         BorderStyle     =   0  'None
         Height          =   375
         Left            =   120
         ScaleHeight     =   375
         ScaleWidth      =   11655
         TabIndex        =   6
         Top             =   90
         Width           =   11655
         Begin VB.PictureBox Closed 
            Appearance      =   0  'Flat
            AutoRedraw      =   -1  'True
            BackColor       =   &H00CCCCCC&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   280
            Left            =   0
            MouseIcon       =   "frmMain.frx":531E
            MousePointer    =   99  'Custom
            ScaleHeight     =   285
            ScaleWidth      =   315
            TabIndex        =   9
            Top             =   60
            Width           =   315
         End
         Begin VB.PictureBox Minimized 
            Appearance      =   0  'Flat
            AutoRedraw      =   -1  'True
            BackColor       =   &H00CCCCCC&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   280
            Left            =   11335
            MouseIcon       =   "frmMain.frx":5628
            MousePointer    =   99  'Custom
            ScaleHeight     =   285
            ScaleWidth      =   315
            TabIndex        =   8
            Top             =   60
            Width           =   315
         End
         Begin VB.PictureBox Maximized 
            Appearance      =   0  'Flat
            AutoRedraw      =   -1  'True
            BackColor       =   &H00CCCCCC&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   280
            Left            =   11065
            MouseIcon       =   "frmMain.frx":5932
            MousePointer    =   99  'Custom
            ScaleHeight     =   285
            ScaleWidth      =   315
            TabIndex        =   7
            Top             =   60
            Width           =   315
         End
      End
      Begin VB.PictureBox MenuContainer 
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   0  'None
         Height          =   8055
         Left            =   240
         ScaleHeight     =   8055
         ScaleWidth      =   3015
         TabIndex        =   1
         Top             =   600
         Width           =   3015
         Begin VB.Timer Timer1 
            Interval        =   1000
            Left            =   2040
            Top             =   0
         End
         Begin VB.PictureBox cmdAbout 
            Appearance      =   0  'Flat
            AutoRedraw      =   -1  'True
            BackColor       =   &H00CCCCCC&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   450
            Left            =   360
            MouseIcon       =   "frmMain.frx":5C3C
            MousePointer    =   99  'Custom
            ScaleHeight     =   450
            ScaleWidth      =   2355
            TabIndex        =   14
            Top             =   6840
            Width           =   2355
         End
         Begin VB.PictureBox cmdShutdown 
            Appearance      =   0  'Flat
            AutoRedraw      =   -1  'True
            BackColor       =   &H00CCCCCC&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   450
            Left            =   360
            MouseIcon       =   "frmMain.frx":5F46
            MousePointer    =   99  'Custom
            ScaleHeight     =   450
            ScaleWidth      =   2355
            TabIndex        =   5
            Top             =   7440
            Width           =   2355
         End
         Begin VB.PictureBox cmdChangeUser 
            Appearance      =   0  'Flat
            AutoRedraw      =   -1  'True
            BackColor       =   &H00CCCCCC&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   450
            Left            =   360
            MouseIcon       =   "frmMain.frx":6250
            MousePointer    =   99  'Custom
            ScaleHeight     =   450
            ScaleWidth      =   2355
            TabIndex        =   4
            Top             =   6240
            Width           =   2355
         End
         Begin VB.ListBox MenuList 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   5520
            Left            =   240
            TabIndex        =   3
            Top             =   600
            Width           =   2535
         End
         Begin VB.PictureBox MenuHeader 
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   0  'None
            Height          =   375
            Left            =   120
            ScaleHeight     =   375
            ScaleWidth      =   2775
            TabIndex        =   2
            Top             =   90
            Width           =   2775
            Begin Crystal.CrystalReport ReportIt 
               Left            =   600
               Top             =   0
               _ExtentX        =   741
               _ExtentY        =   741
               _Version        =   348160
               WindowLeft      =   230
               WindowTop       =   45
               WindowWidth     =   552
               WindowHeight    =   494
               WindowBorderStyle=   0
               WindowControlBox=   0   'False
               WindowMaxButton =   0   'False
               WindowMinButton =   0   'False
               PrintFileLinesPerPage=   60
            End
         End
      End
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   2
      Height          =   8985
      Left            =   20
      Top             =   20
      Width           =   11990
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'##############################################
'#          Coded by Walter A. Narvasa        #
'#       POS2000 - Point of Sales System      #
'#                                            #
'#           area :  frmMain                  #
'#    description :  Main Menu-Command Center #
'#         e-mail :  walter@wancom.8k.com     #
'#            url :  http://wancom.8k.com     #
'#                                            #
'##############################################

Public What_Rpt As String
Private Declare Function GetDiskFreeSpace Lib "kernel32" Alias "GetDiskFreeSpaceA" (ByVal lpRootPathName As String, lpSectorsPerCluster As Long, lpBytesPerSector As Long, lpNumberOfFreeClusters As Long, lpTotalNumberOfClusters As Long) As Long
Private Declare Function GetVolumeInformation Lib "kernel32" Alias "GetVolumeInformationA" (ByVal lpRootPathName As String, ByVal lpVolumeNameBuffer As String, ByVal nVolumeNameSize As Long, lpVolumeSerialNumber As Long, lpMaximumComponentLength As Long, lpFileSystemFlags As Long, ByVal lpFileSystemNameBuffer As String, ByVal nFileSystemNameSize As Long) As Long
Private Declare Function GetDriveType Lib "kernel32" Alias "GetDriveTypeA" (ByVal nDrive As String) As Long
Dim i, intDiskSize, intfreeKB, intUsedKB As Long
Dim nReturnValue, SectorsPerCluster, BytesPerSector As Long
Dim TotalClusters, FreeClusters As Long
'Dim pnlError As Panel
Dim sDriveLetter, sDrive As String
Dim fs, abc, dc As Integer
Private Const SR = 0
Private Const GDI = 1
Private Const USR = 2
Private Const VER_PLATFORM_WIN32s = 0
Private Const VER_PLATFORM_WIN32_WINDOWS = 1
Private Const VER_PLATFORM_WIN32_NT = 2
Private Type OSVERSIONINFO
    dwOSVersionInfoSize As Long
    dwMajorVersion      As Long
    dwMinorVersion      As Long
    dwBuildNumber       As Long
    dwPlatformId        As Long
    szCSDVersion        As String * 128
End Type
Private Type MEMORYSTATUS
    dwLength        As Long
    dwMemoryLoad    As Long
    dwTotalPhys     As Long
    dwAvailPhys     As Long
    dwTotalPageFile As Long
    dwAvailPageFile As Long
    dwTotalVirtual  As Long
    dwAvailVirtual  As Long
End Type
Private Declare Sub GlobalMemoryStatus Lib "kernel32" (lpBuffer As MEMORYSTATUS)
Private Declare Function pBGetFreeSystemResources Lib "rsrc32.dll" Alias "_MyGetFreeSystemResources32@4" (ByVal iResType As Integer) As Integer
Private Declare Function GetVersionEx Lib "kernel32" Alias "GetVersionExA" (lpVersionInformation As OSVERSIONINFO) As Long
Private mIsWin32 As Boolean
Dim datprimary As DAO.Recordset
Dim datsecondary As DAO.Recordset
Dim datthirdary As DAO.Recordset

Private Sub Closed_Click()
    UnloadAllForms
End Sub

Private Sub Closed_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call BitBlt(frmMain.Closed.hDC, 0, 0, 73, 50, frmLogin.Source.hDC, 18, 107, SRCCOPY)
    frmMain.Closed.Refresh
End Sub

Private Sub Closed_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call BitBlt(frmMain.Closed.hDC, 0, 0, 73, 50, frmLogin.Source.hDC, 0, 107, SRCCOPY)
    frmMain.Closed.Refresh
End Sub

Private Sub cmdAbout_Click()
    frmAbout.Show
End Sub

Private Sub cmdAbout_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call MacButton("            About", frmMain.cmdAbout, 0, 0, 170, 30, frmLogin.Source, 182, 30, 2)
End Sub

Private Sub cmdAbout_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call MacButton("            About", frmMain.cmdAbout, 0, 0, 170, 30, frmLogin.Source, 147, 0, 2)
End Sub

Private Sub cmdChangeUser_Click()
    frmLogin.txtUserName = ""
    frmLogin.txtPassword = ""
    frmLogin.Show
    frmLogin.txtUserName.SetFocus
End Sub

Private Sub cmdChangeUser_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call MacButton("       Change User", frmMain.cmdChangeUser, 0, 0, 170, 30, frmLogin.Source, 182, 30, 2)
End Sub

Private Sub cmdChangeUser_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call MacButton("       Change User", frmMain.cmdChangeUser, 0, 0, 170, 30, frmLogin.Source, 147, 0, 2)
End Sub

Private Sub cmdShutdown_Click()
    'UnloadAllForms
    Call DoExitWindows
End Sub

Private Sub cmdShutdown_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call MacButton("         Shutdown", frmMain.cmdShutdown, 0, 0, 170, 30, frmLogin.Source, 182, 30, 2)
End Sub

Private Sub cmdShutdown_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call MacButton("         Shutdown", frmMain.cmdShutdown, 0, 0, 170, 30, frmLogin.Source, 147, 0, 2)
End Sub


Private Sub Form_Load()
    Dim VolName As String, fSys As String
    Dim Drive As String, DriveType As Long, erg As Long
    Dim OSInfo As OSVERSIONINFO
        OSInfo.dwOSVersionInfoSize = Len(OSInfo)
        Call GetVersionEx(OSInfo)
        mIsWin32 = (OSInfo.dwPlatformId = VER_PLATFORM_WIN32_WINDOWS)
        tmrUpdate_Timer
        tmrUpdate.Enabled = True
    sDriveLetter = "C:\"
    nReturnValue = GetDiskFreeSpace(sDriveLetter, SectorsPerCluster, BytesPerSector, FreeClusters, TotalClusters)
    If nReturnValue = "0" Then
        txtLabel = ""
        txtFsys = ""
        txtDriveType = ""
        txtFreeSpace = ""
        txtDiskSize = ""
        txtUsedKB = ""
        'lblFree.Visible = False
        'lblTotal.Visible = False
        'lblUsed.Visible = False
   Else
        VolName = Space(127)
        fSys = Space(127)
        Drive = sDriveLetter
        DriveType& = GetDriveType(Drive$)
        'erg& = GetVolumeInformation(Drive$, VolName$, 127, 0, 0, 0, fSys, 127)
        txtLabel = VolName$
        txtFsys = fSys$
    Select Case DriveType
        Case 1:
            txtDriveType = "Unknown"
        Case 2:
            txtDriveType = "Removable"
        Case 3:
            txtDriveType = "Fixed"
        Case 4:
            txtDriveType = "Remote"
        Case 5:
            txtDriveType = "CD-Rom"
        Case 6:
            txtDriveType = "RAMdisk"
        Case Else:
            txtDriveType = "*ERROR*"
    End Select
    intDiskSize = SectorsPerCluster * BytesPerSector * (TotalClusters \ 1024) \ 1024
    intfreeKB = SectorsPerCluster * BytesPerSector * (FreeClusters \ 1024) \ 1024
    intUsedKB = intDiskSize - intfreeKB
    If intDiskSize > 1024 Then
       intDiskSize = Format((intDiskSize / 1024), "##,##0.00")
       lblTotal = "Gigabytes"
       lblTotal.Visible = True
     Else
       lblTotal = "MBytes"
       lblTotal.Visible = True
    End If
    
    If intfreeKB > 1024 Then
       intfreeKB = Format((intfreeKB / 1024), "##,##0.00")
       lblFree = "Gigabytes"
       lblFree.Visible = True
     Else
       lblFree = "MBytes"
       lblFree.Visible = True
    End If
    If intUsedKB > 1024 Then
       intUsedKB = Format((intUsedKB / 1024), "##,##0.00")
       lblUsed = "Gigabytes"
       lblUsed.Visible = True
     Else
       lblUsed = "MBytes"
       lblUsed.Visible = True
    End If
    txtUsedKB = intUsedKB
    txtFreeSpace = intfreeKB
    txtDiskSize = intDiskSize
  End If
    Call ColForm(BoxContainer, 217, 211, 213, 125)
    Call ColForm(MenuContainer, 217, 211, 213, 125)
    Call ColForm(Applets, 217, 211, 213, 125)
    Call CreateMacOSTitleBar(titleBar, " Point-of-Sales System 2000-Pharmaceutical Version ")
    Call CreateMacOSTitleBar(MenuHeader, " Main Menu ")
    Call MacButton("       Change User", frmMain.cmdChangeUser, 0, 0, 170, 30, frmLogin.Source, 147, 0, 2)
    Call MacButton("            About", frmMain.cmdAbout, 0, 0, 170, 30, frmLogin.Source, 147, 0, 2)
    Call MacButton("         Shutdown", frmMain.cmdShutdown, 0, 0, 170, 30, frmLogin.Source, 147, 0, 2)
    Call BitBlt(frmMain.Closed.hDC, 0, 0, 73, 50, frmLogin.Source.hDC, 0, 107, SRCCOPY)
    frmMain.Closed.Refresh
    Call BitBlt(frmMain.Maximized.hDC, 0, 0, 73, 50, frmLogin.Source.hDC, 0, 72, SRCCOPY)
    frmMain.Maximized.Refresh
    Call BitBlt(frmMain.Minimized.hDC, 0, 0, 73, 50, frmLogin.Source.hDC, 0, 124, SRCCOPY)
    frmMain.Minimized.Refresh
    KeyPreview = True
    StatusMessage.Caption = " Today is " + Date$
    Set datsecondary = frmLogin.db.OpenRecordset("select * from INVOICE order by INVOICE_NO")
    Set datthirdary = frmLogin.db.OpenRecordset("select * from INVOICE_DETAIL order by INVOICE_NOD")
    Set datprimary = frmLogin.db.OpenRecordset("select * from SETUP order by COMPANY_NAME")
End Sub

Public Sub Operation_CleanUp()
    'On Error Resume Next
    If datsecondary.RecordCount <> 0 Then
        datsecondary.MoveFirst
        Do While Not datsecondary.EOF
            datsecondary.Delete
            datsecondary.MoveNext
        Loop
    End If
    If datthirdary.RecordCount <> 0 Then
        datthirdary.MoveFirst
        Do While Not datthirdary.EOF
            datthirdary.Delete
            datthirdary.MoveNext
        Loop
    End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    On Error Resume Next
    Dim AltDown
    AltDown = (Shift And vbAltMask) > 0
    Select Case KeyCode
        Case vbKeyEscape:
                UnloadAllForms
                End
        Case vbKeyC:
                If AltDown Then
                    Call MacButton("       Change User", frmMain.cmdChangeUser, 0, 0, 170, 30, frmLogin.Source, 182, 30, 2)
                    cmdChangeUser_Click
                    Call MacButton("       Change User", frmMain.cmdChangeUser, 0, 0, 170, 30, frmLogin.Source, 147, 0, 2)
                End If
        Case vbKeyA:
                If AltDown Then
                    Call MacButton("            About", frmMain.cmdAbout, 0, 0, 170, 30, frmLogin.Source, 182, 30, 2)
                    cmdAbout_Click
                    Call MacButton("            About", frmMain.cmdAbout, 0, 0, 170, 30, frmLogin.Source, 147, 0, 2)
                End If
        Case vbKeyS:
                If AltDown Then
                    Call MacButton("         Shutdown", frmMain.cmdShutdown, 0, 0, 170, 30, frmLogin.Source, 182, 30, 2)
                    cmdShutdown_Click
                    Call MacButton("         Shutdown", frmMain.cmdShutdown, 0, 0, 170, 30, frmLogin.Source, 147, 0, 2)
                End If
        Case vbKeyReturn:
                Call MenuFunction
    End Select
End Sub
    
Function MenuFunction()
                If MenuList = "Supplier Masterfile" Then
                    frmSupplier.Show
                ElseIf MenuList = "Product Masterfile" Then
                    frmProduct.Show
                ElseIf MenuList = "Category Masterfile" Then
                    frmCategory.Show
                ElseIf MenuList = "Selling Transaction" Then
                    frmSelling.Show
                ElseIf MenuList = "Receiving Transaction" Then
                    frmReceiving.Show
                ElseIf MenuList = "Stocks Re-order Report" Then
                    ReportIt.ReportFileName = App.Path & "\REPORTS\PRODLEFT.RPT"
                    ReportIt.WindowTitle = MenuList
                    ReportIt.Action = 1
                ElseIf MenuList = "Selling History Report" Then
                    frmPrint.Show
                    frmPrint.PrintIt.WindowTitle = MenuList
                    What_Rpt = MenuList
                ElseIf MenuList = "Receiving History Report" Then
                    frmPrint.Show
                    frmPrint.PrintIt.WindowTitle = MenuList
                    What_Rpt = MenuList
                ElseIf MenuList = "Stocks Per Supplier Report" Then
                    frmPrint.Show
                    frmPrint.PrintIt.WindowTitle = MenuList
                    What_Rpt = MenuList
                ElseIf MenuList = "Product Listing Report" Then
                    ReportIt.ReportFileName = App.Path & "\REPORTS\PRODLIST.RPT"
                    ReportIt.WindowTitle = MenuList
                    ReportIt.Action = 1
                ElseIf MenuList = "Supplier Listing Report" Then
                    ReportIt.ReportFileName = App.Path & "\REPORTS\SUPPLIER.RPT"
                    ReportIt.WindowTitle = MenuList
                    ReportIt.Action = 1
                ElseIf MenuList = "Backup/Restore Files" Then
                    frmBackup.Show
                ElseIf MenuList = "Password Security" Then
                    frmPassword.Show
                ElseIf MenuList = "Code File Setup" Then
                    frmCodeFile.Show
                ElseIf MenuList = "Software Setup" Then
                    frmSetup.Show
                End If
End Function

Private Sub Maximized_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call BitBlt(frmMain.Maximized.hDC, 0, 0, 73, 50, frmLogin.Source.hDC, 18, 72, SRCCOPY)
    frmMain.Maximized.Refresh
End Sub

Private Sub Maximized_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call BitBlt(frmMain.Maximized.hDC, 0, 0, 73, 50, frmLogin.Source.hDC, 0, 72, SRCCOPY)
    frmMain.Maximized.Refresh
End Sub

Private Sub MenuList_Click()
'    Call MenuFunction
End Sub

Private Sub Minimized_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call BitBlt(frmMain.Minimized.hDC, 0, 0, 73, 50, frmLogin.Source.hDC, 18, 124, SRCCOPY)
    frmMain.Minimized.Refresh
End Sub

Private Sub Minimized_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call BitBlt(frmMain.Minimized.hDC, 0, 0, 73, 50, frmLogin.Source.hDC, 0, 124, SRCCOPY)
    frmMain.Minimized.Refresh
End Sub

Private Sub Timer1_Timer()
    On Error Resume Next
    If datsecondary.RecordCount = 0 And datthirdary.RecordCount = 0 Then
    Else
        If datprimary("SELLING_CLEANUPDATE") <= Date Then
            Call MessageBox("frmMain", "You are scheduled for Invoice Cleanup." & vbCrLf & _
                            "Do you want to proceed with Invoice Cleanup ?", 1)
            frmMessageBox2.Show
            frmMessageBox2.SetFocus
        End If
    End If
End Sub

Private Sub titleBar_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call DragForm(Me)
End Sub

Private Sub tmrUpdate_Timer()
    On Local Error Resume Next
    Dim MS As MEMORYSTATUS
    Today = Now
    Clock.Caption = Format(Today, "hh:mm:ss ampm ")
    MS.dwLength = Len(MS)
    Call GlobalMemoryStatus(MS)
    With MS
        lbl(0) = Format$(.dwTotalPhys / 1024, "#,###") & " Kb"
        lbl(1) = Format$(.dwAvailPhys / 1024, "#,###") & " Kb"
        lbl(2) = Format$(.dwTotalVirtual / 1024, "#,###") & " Kb"
        lbl(3) = Format$(.dwAvailVirtual / 1024, "#,###") & " Kb"
        lbl(4) = Format$(.dwTotalPageFile / 1024, "#,###") & " Kb"
        lbl(5) = Format$(.dwAvailPageFile / 1024, "#,###") & " Kb"
        lbl(6) = Format$(.dwMemoryLoad, "##0") & "%"
    End With
    If mIsWin32 Then
        Dim nIndex As Integer
            guageSystem.Value = pBGetFreeSystemResources(SR)
            guageSystem.ForeColor = AssignColour(guageSystem.Value)
            guageUser.Value = pBGetFreeSystemResources(USR)
            guageUser.ForeColor = AssignColour(guageUser.Value)
            guageGDI.Value = pBGetFreeSystemResources(GDI)
            guageGDI.ForeColor = AssignColour(guageGDI.Value)
            Select Case guageSystem.Value
                Case Is > 99
                Case Is < 1
                Case Else
                    nIndex = (100 - guageSystem.Value) * 0.13
                Select Case nIndex
                    Case Is < 0
                        nIndex = 0
                    Case Is > 12
                        nIndex = 12
                End Select
        End Select
    End If
End Sub

Private Function AssignColour(nPercent As Integer) As Long
    Select Case nPercent
        Case Is < 10
            AssignColour = vbRed
        Case Is < 25
            AssignColour = vbMagenta
        Case Else
            AssignColour = vbBlue
    End Select
End Function

