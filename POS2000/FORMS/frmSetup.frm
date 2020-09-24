VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmSetup 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   0  'None
   Caption         =   "frmProduct"
   ClientHeight    =   6375
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   9330
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6375
   ScaleWidth      =   9330
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox BoxContainer 
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   0  'None
      FillColor       =   &H00C0C0C0&
      Height          =   6300
      Left            =   40
      ScaleHeight     =   6300
      ScaleWidth      =   9225
      TabIndex        =   15
      TabStop         =   0   'False
      Top             =   40
      Width           =   9220
      Begin VB.PictureBox Logo 
         BackColor       =   &H00000000&
         FillColor       =   &H00808080&
         ForeColor       =   &H00808080&
         Height          =   975
         Left            =   7680
         ScaleHeight     =   915
         ScaleWidth      =   1215
         TabIndex        =   36
         TabStop         =   0   'False
         Top             =   720
         Width           =   1275
      End
      Begin VB.TextBox txtField 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   2
         Left            =   2460
         MultiLine       =   -1  'True
         TabIndex        =   2
         Top             =   1440
         Width           =   2535
      End
      Begin VB.PictureBox CheckContainer 
         BackColor       =   &H00808080&
         BorderStyle     =   0  'None
         Height          =   3015
         Left            =   240
         ScaleHeight     =   3015
         ScaleWidth      =   8775
         TabIndex        =   16
         TabStop         =   0   'False
         Top             =   1920
         Width           =   8775
         Begin VB.CheckBox txtCheck 
            BackColor       =   &H00C0C0C0&
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
            Index           =   6
            Left            =   360
            TabIndex        =   9
            Top             =   1680
            Width           =   200
         End
         Begin VB.TextBox txtField 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Index           =   3
            Left            =   2280
            MultiLine       =   -1  'True
            TabIndex        =   10
            Top             =   2040
            Width           =   615
         End
         Begin VB.TextBox txtField 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   825
            Index           =   5
            Left            =   4560
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   13
            Top             =   1920
            Width           =   3855
         End
         Begin VB.TextBox txtField 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   825
            Index           =   4
            Left            =   4560
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   12
            Top             =   840
            Width           =   3855
         End
         Begin VB.CheckBox txtCheck 
            BackColor       =   &H00C0C0C0&
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
            Index           =   5
            Left            =   360
            TabIndex        =   8
            Top             =   1440
            Width           =   200
         End
         Begin VB.CheckBox txtCheck 
            BackColor       =   &H00C0C0C0&
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
            Index           =   4
            Left            =   360
            TabIndex        =   7
            Top             =   1200
            Width           =   200
         End
         Begin VB.CheckBox txtCheck 
            BackColor       =   &H00C0C0C0&
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
            Index           =   3
            Left            =   360
            TabIndex        =   6
            Top             =   960
            Width           =   200
         End
         Begin VB.CheckBox txtCheck 
            BackColor       =   &H00C0C0C0&
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
            Index           =   2
            Left            =   360
            TabIndex        =   5
            Top             =   720
            Width           =   200
         End
         Begin VB.CheckBox txtCheck 
            BackColor       =   &H00C0C0C0&
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
            Index           =   1
            Left            =   360
            TabIndex        =   4
            Top             =   480
            Width           =   200
         End
         Begin VB.CheckBox txtCheck 
            BackColor       =   &H00C0C0C0&
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
            Index           =   0
            Left            =   360
            TabIndex        =   3
            Top             =   240
            Width           =   200
         End
         Begin MSComCtl2.DTPicker DatePick 
            Height          =   315
            Index           =   0
            Left            =   2880
            TabIndex        =   14
            TabStop         =   0   'False
            Top             =   2040
            Visible         =   0   'False
            Width           =   1455
            _ExtentX        =   2566
            _ExtentY        =   556
            _Version        =   393216
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            CustomFormat    =   "MM/dd/yyyy"
            Format          =   24772611
            CurrentDate     =   36511
         End
         Begin MSComCtl2.DTPicker DatePick 
            Height          =   315
            Index           =   1
            Left            =   6840
            TabIndex        =   11
            Top             =   195
            Width           =   1455
            _ExtentX        =   2566
            _ExtentY        =   556
            _Version        =   393216
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            CustomFormat    =   "MM/dd/yyyy"
            Format          =   24772611
            CurrentDate     =   36511
         End
         Begin MSComDlg.CommonDialog cdlPicture 
            Left            =   0
            Top             =   0
            _ExtentX        =   847
            _ExtentY        =   847
            _Version        =   393216
         End
         Begin VB.Label Label 
            AutoSize        =   -1  'True
            BackColor       =   &H00808080&
            BackStyle       =   0  'Transparent
            Caption         =   "Disable POS2000 Wallpaper"
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
            Index           =   15
            Left            =   600
            TabIndex        =   43
            Top             =   1680
            Width           =   2640
         End
         Begin VB.Label Label 
            AutoSize        =   -1  'True
            BackColor       =   &H00808080&
            BackStyle       =   0  'Transparent
            Caption         =   "Invoice Clean-up Date"
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
            Index           =   14
            Left            =   4560
            TabIndex        =   41
            Top             =   240
            Width           =   2190
         End
         Begin VB.Label Label 
            AutoSize        =   -1  'True
            BackColor       =   &H00808080&
            BackStyle       =   0  'Transparent
            Caption         =   "System Expiration"
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
            Index           =   13
            Left            =   360
            TabIndex        =   40
            Top             =   2040
            Width           =   1785
         End
         Begin VB.Label Label 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackColor       =   &H00808080&
            BackStyle       =   0  'Transparent
            Caption         =   "Invoice Selling Footer"
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
            Index           =   10
            Left            =   4560
            TabIndex        =   39
            Top             =   1680
            Width           =   2145
         End
         Begin VB.Label Label 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackColor       =   &H00808080&
            BackStyle       =   0  'Transparent
            Caption         =   "Invoice Selling Header"
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
            Index           =   2
            Left            =   4560
            TabIndex        =   38
            Top             =   600
            Width           =   2190
         End
         Begin VB.Label Label 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackColor       =   &H00808080&
            BackStyle       =   0  'Transparent
            Caption         =   "F1=Help"
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
            Index           =   11
            Left            =   480
            TabIndex        =   35
            Top             =   2520
            Width           =   765
         End
         Begin VB.Label Label 
            AutoSize        =   -1  'True
            BackColor       =   &H00808080&
            BackStyle       =   0  'Transparent
            Caption         =   "Enable 30 Days Trialer/Expiration"
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
            Index           =   9
            Left            =   600
            TabIndex        =   34
            Top             =   960
            Width           =   3270
         End
         Begin VB.Label Label 
            AutoSize        =   -1  'True
            BackColor       =   &H00808080&
            BackStyle       =   0  'Transparent
            Caption         =   "Auto-Loading at Startup"
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
            Index           =   6
            Left            =   600
            TabIndex        =   33
            Top             =   240
            Width           =   2385
         End
         Begin VB.Label Label 
            AutoSize        =   -1  'True
            BackColor       =   &H00808080&
            BackStyle       =   0  'Transparent
            Caption         =   "Disable Right-Click Mouse"
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
            Index           =   5
            Left            =   600
            TabIndex        =   32
            Top             =   1200
            Width           =   2550
         End
         Begin VB.Label Label 
            AutoSize        =   -1  'True
            BackColor       =   &H00808080&
            BackStyle       =   0  'Transparent
            Caption         =   "Disable Ctr-Alt-Del/Alt-Tab Shortcut"
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
            Index           =   4
            Left            =   600
            TabIndex        =   31
            Top             =   1440
            Width           =   3555
         End
         Begin VB.Label Label 
            AutoSize        =   -1  'True
            BackColor       =   &H00808080&
            BackStyle       =   0  'Transparent
            Caption         =   "Hide Desktop Icons"
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
            Index           =   1
            Left            =   600
            TabIndex        =   30
            Top             =   720
            Width           =   1875
         End
         Begin VB.Label Label 
            AutoSize        =   -1  'True
            BackColor       =   &H00808080&
            BackStyle       =   0  'Transparent
            Caption         =   "Hide Windows Taskbar"
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
            Index           =   0
            Left            =   600
            TabIndex        =   29
            Top             =   480
            Width           =   2190
         End
      End
      Begin VB.TextBox txtField 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   1
         Left            =   2460
         MultiLine       =   -1  'True
         TabIndex        =   1
         Top             =   1080
         Width           =   4455
      End
      Begin VB.TextBox txtField 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   0
         Left            =   2460
         Locked          =   -1  'True
         TabIndex        =   0
         Top             =   720
         Width           =   4455
      End
      Begin VB.PictureBox titleBar 
         BorderStyle     =   0  'None
         Height          =   375
         Left            =   120
         ScaleHeight     =   375
         ScaleWidth      =   9015
         TabIndex        =   22
         TabStop         =   0   'False
         Top             =   90
         Width           =   9015
         Begin VB.PictureBox Maximized 
            Appearance      =   0  'Flat
            AutoRedraw      =   -1  'True
            BackColor       =   &H00CCCCCC&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   280
            Left            =   8430
            MouseIcon       =   "frmSetup.frx":0000
            MousePointer    =   99  'Custom
            ScaleHeight     =   285
            ScaleWidth      =   315
            TabIndex        =   25
            TabStop         =   0   'False
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
            Left            =   8700
            MouseIcon       =   "frmSetup.frx":030A
            MousePointer    =   99  'Custom
            ScaleHeight     =   285
            ScaleWidth      =   315
            TabIndex        =   24
            TabStop         =   0   'False
            Top             =   60
            Width           =   315
         End
         Begin VB.PictureBox Closed 
            Appearance      =   0  'Flat
            AutoRedraw      =   -1  'True
            BackColor       =   &H00CCCCCC&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   280
            Left            =   0
            MouseIcon       =   "frmSetup.frx":0614
            MousePointer    =   99  'Custom
            ScaleHeight     =   285
            ScaleWidth      =   315
            TabIndex        =   23
            TabStop         =   0   'False
            Top             =   60
            Width           =   315
         End
      End
      Begin VB.PictureBox ButtonContainer 
         BackColor       =   &H00808080&
         BorderStyle     =   0  'None
         Height          =   975
         Left            =   240
         ScaleHeight     =   975
         ScaleWidth      =   8775
         TabIndex        =   17
         TabStop         =   0   'False
         Top             =   5040
         Width           =   8775
         Begin VB.PictureBox cmdPicture 
            Appearance      =   0  'Flat
            AutoRedraw      =   -1  'True
            BackColor       =   &H00CCCCCC&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   450
            Left            =   600
            MouseIcon       =   "frmSetup.frx":091E
            MousePointer    =   99  'Custom
            ScaleHeight     =   450
            ScaleWidth      =   2355
            TabIndex        =   42
            TabStop         =   0   'False
            Top             =   300
            Width           =   2355
         End
         Begin VB.PictureBox cmdExit 
            Appearance      =   0  'Flat
            AutoRedraw      =   -1  'True
            BackColor       =   &H00CCCCCC&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   570
            Left            =   7080
            MouseIcon       =   "frmSetup.frx":0C28
            MousePointer    =   99  'Custom
            ScaleHeight     =   570
            ScaleWidth      =   1095
            TabIndex        =   21
            TabStop         =   0   'False
            Top             =   240
            Width           =   1100
         End
         Begin VB.PictureBox cmdEdit 
            Appearance      =   0  'Flat
            AutoRedraw      =   -1  'True
            BackColor       =   &H00CCCCCC&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   570
            Left            =   3480
            MouseIcon       =   "frmSetup.frx":0F32
            MousePointer    =   99  'Custom
            ScaleHeight     =   570
            ScaleWidth      =   1095
            TabIndex        =   20
            TabStop         =   0   'False
            Top             =   240
            Width           =   1100
         End
         Begin VB.PictureBox cmdSave 
            Appearance      =   0  'Flat
            AutoRedraw      =   -1  'True
            BackColor       =   &H00CCCCCC&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   570
            Left            =   4680
            MouseIcon       =   "frmSetup.frx":123C
            MousePointer    =   99  'Custom
            ScaleHeight     =   570
            ScaleWidth      =   1095
            TabIndex        =   19
            TabStop         =   0   'False
            Top             =   240
            Width           =   1100
         End
         Begin VB.PictureBox cmdUndo 
            Appearance      =   0  'Flat
            AutoRedraw      =   -1  'True
            BackColor       =   &H00CCCCCC&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   570
            Left            =   5880
            MouseIcon       =   "frmSetup.frx":1546
            MousePointer    =   99  'Custom
            ScaleHeight     =   570
            ScaleWidth      =   1095
            TabIndex        =   18
            TabStop         =   0   'False
            Top             =   240
            Width           =   1100
         End
      End
      Begin VB.Label Label 
         AutoSize        =   -1  'True
         BackColor       =   &H00808080&
         BackStyle       =   0  'Transparent
         Caption         =   "Logo"
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
         Index           =   12
         Left            =   7080
         TabIndex        =   37
         Top             =   720
         Width           =   465
      End
      Begin VB.Label Label 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H00808080&
         BackStyle       =   0  'Transparent
         Caption         =   "Address"
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
         Index           =   7
         Left            =   1530
         TabIndex        =   28
         Top             =   1080
         Width           =   780
      End
      Begin VB.Label Label 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H00808080&
         BackStyle       =   0  'Transparent
         Caption         =   "Registered to: Name"
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
         Index           =   3
         Left            =   300
         TabIndex        =   27
         Top             =   720
         Width           =   2010
      End
      Begin VB.Label Label 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H00808080&
         BackStyle       =   0  'Transparent
         Caption         =   "Telephone Number"
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
         Index           =   8
         Left            =   495
         TabIndex        =   26
         Top             =   1440
         Width           =   1815
      End
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   2
      Height          =   6360
      Left            =   15
      Top             =   15
      Width           =   9300
   End
End
Attribute VB_Name = "frmSetup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'##############################################
'#          Coded by Walter A. Narvasa        #
'#       POS2000 - Point of Sales System      #
'#                                            #
'#           area :  frmSetup                 #
'#    description :  Software Setup           #
'#         e-mail :  walter@wancom.8k.com     #
'#            url :  http://wancom.8k.com     #
'#                                            #
'##############################################

Dim dummy As DAO.Recordset
Dim datprimary As DAO.Recordset
Dim rec_Isnew As Boolean
Dim p_add, p_edit, p_save, p_undo, p_top, p_prev, p_next, p_last, p_del
Dim p_isadding, p_isediting, p_isnavigate

Private Sub cmdEdit_Click()
'    On Error GoTo ErrorEdit
    EditMode = True
    Press_Buttons ("Edit")
    txtField(0).SetFocus
'ErrorEdit:
'    Call MessageBox("frmSetup", Err.Description, 0)
End Sub

Private Sub cmdEdit_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call MacButton("    Edit", frmSetup.cmdEdit, 0, 0, 73, 50, frmLogin.Source, 74, 0, 1)
End Sub

Private Sub cmdEdit_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call MacButton("    Edit", frmSetup.cmdEdit, 0, 0, 73, 50, frmLogin.Source, 0, 0, 1)
End Sub

Private Sub cmdExit_Click()
    EditMode = False
    Unload Me
End Sub

Private Sub cmdExit_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call MacButton("    Exit", frmSetup.cmdExit, 0, 0, 73, 50, frmLogin.Source, 74, 0, 1)
End Sub

Private Sub cmdExit_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call MacButton("    Exit", frmSetup.cmdExit, 0, 0, 73, 50, frmLogin.Source, 0, 0, 1)
End Sub

Private Sub cmdPicture_Click()
    On Error GoTo CancelLoad
        cdlPicture.ShowOpen
    On Error GoTo BadLoad
        Logo = LoadPicture(cdlPicture.filename)
    On Error GoTo 0
    On Error GoTo 0
    Exit Sub
CancelLoad:
    If Err.Number <> cdlCancel Then
        Call MessageBox("frmSetup", Err.Description, 0)
        frmMessageBox.Show
    Else
        Exit Sub
    End If
BadLoad:
        Call MessageBox("frmSetup", Err.Description, 0)
        frmMessageBox.Show
    Exit Sub
End Sub

Private Sub cmdPicture_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call MacButton("      Change Logo", frmSetup.cmdPicture, 0, 0, 170, 30, frmLogin.Source, 182, 30, 2)
End Sub

Private Sub cmdPicture_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call MacButton("      Change Logo", frmSetup.cmdPicture, 0, 0, 170, 30, frmLogin.Source, 147, 0, 2)
End Sub

Private Sub cmdSave_Click()
    On Error Resume Next
    Call MacButton("    Edit", frmSetup.cmdEdit, 0, 0, 73, 50, frmLogin.Source, 0, 0, 1)
    Press_Buttons ("Save")
    Call MacButton("   Save", frmSetup.cmdSave, 0, 0, 73, 50, frmLogin.Source, 0, 0, 1)
End Sub

Private Sub cmdSave_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call MacButton("   Save", frmSetup.cmdSave, 0, 0, 73, 50, frmLogin.Source, 74, 0, 1)
End Sub

Private Sub cmdSave_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call MacButton("   Save", frmSetup.cmdSave, 0, 0, 73, 50, frmLogin.Source, 0, 0, 1)
End Sub

Private Sub cmdUndo_Click()
    Call MacButton("    Edit", frmSetup.cmdEdit, 0, 0, 73, 50, frmLogin.Source, 0, 0, 1)
    Press_Buttons ("Undo")
    Call MacButton("   Undo", frmSetup.cmdUndo, 0, 0, 73, 50, frmLogin.Source, 0, 0, 1)
End Sub

Private Sub cmdUndo_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call MacButton("   Undo", frmSetup.cmdUndo, 0, 0, 73, 50, frmLogin.Source, 74, 0, 1)
End Sub

Private Sub cmdUndo_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call MacButton("   Undo", frmSetup.cmdUndo, 0, 0, 73, 50, frmLogin.Source, 0, 0, 1)
End Sub

Private Sub Form_Activate()
    On Error Resume Next
    If EditMode = False Then
        strs = "select * from SETUP order by COMPANY_NAME"
        Set datprimary = frmLogin.db.OpenRecordset(strs)
        If datprimary.AbsolutePosition <> -1 Then
            Display_Fields
        End If
            Enable_Fields (True)
            Object_Tab_Trigger (False)
            Enable_Buttons
            Enable_Buttons
            rec_Isnew = False
    End If
End Sub

Private Sub Form_Load()
    On Error Resume Next
    Call ColForm(BoxContainer, 217, 211, 213, 125)
    Call ColForm(ButtonContainer, 217, 211, 213, 125)
    Call ColForm(CheckContainer, 217, 211, 213, 125)
    Call CreateMacOSTitleBar(titleBar, " Software Setup ")
    Call BitBlt(frmSetup.Closed.hDC, 0, 0, 73, 50, frmLogin.Source.hDC, 0, 107, SRCCOPY)
    frmSetup.Closed.Refresh
    Call BitBlt(frmSetup.Maximized.hDC, 0, 0, 73, 50, frmLogin.Source.hDC, 0, 72, SRCCOPY)
    frmSetup.Maximized.Refresh
    Call BitBlt(frmSetup.Minimized.hDC, 0, 0, 73, 50, frmLogin.Source.hDC, 0, 124, SRCCOPY)
    frmSetup.Minimized.Refresh
    KeyPreview = True
    Call MacButton("      Change Logo", frmSetup.cmdPicture, 0, 0, 170, 30, frmLogin.Source, 147, 0, 2)
    Call MacButton("    Edit", frmSetup.cmdEdit, 0, 0, 73, 50, frmLogin.Source, 0, 0, 1)
    Call MacButton("   Save", frmSetup.cmdSave, 0, 0, 73, 50, frmLogin.Source, 0, 0, 1)
    Call MacButton("   Undo", frmSetup.cmdUndo, 0, 0, 73, 50, frmLogin.Source, 0, 0, 1)
    Call MacButton("    Exit", frmSetup.cmdExit, 0, 0, 73, 50, frmLogin.Source, 0, 0, 1)
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    On Error Resume Next
    Dim AltDown
    AltDown = (Shift And vbAltMask) > 0
    Select Case KeyCode
        Case vbKeyEscape:
                Me.Hide
        Case vbKeyL:
                If AltDown Then
                    Call MacButton("      Change Logo", frmSetup.cmdPicture, 0, 0, 170, 30, frmLogin.Source, 147, 0, 2)
                End If
        Case vbKeyE:
                If AltDown Then
                    Call MacButton("    Edit", frmSetup.cmdEdit, 0, 0, 73, 50, frmLogin.Source, 0, 0, 1)
                End If
        Case vbKeyS:
                If AltDown Then
                    Call MacButton("   Save", frmSetup.cmdSave, 0, 0, 73, 50, frmLogin.Source, 0, 0, 1)
                End If
        Case vbKeyU:
                If AltDown Then
                    Call MacButton("   Undo", frmSetup.cmdUndo, 0, 0, 73, 50, frmLogin.Source, 0, 0, 1)
                End If
        Case vbKeyX:
                If AltDown Then
                    Call MacButton("    Exit", frmSetup.cmdExit, 0, 0, 73, 50, frmLogin.Source, 0, 0, 1)
                End If
    End Select
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    On Error Resume Next
    Dim AltDown
    AltDown = (Shift And vbAltMask) > 0
    Select Case KeyCode
        Case vbKeyEscape:
                Me.Hide
        Case vbKeyF1:
                frmHelp.Show
                frmHelp.Help_Values = Space(1) & vbCrLf & _
                                      "Note: The following are Software Setup Key Shortcuts." & vbCrLf & _
                                      Space(1) & vbCrLf & _
                                      "ALT-E=Edit, ALT-S=Save, ALT-U=Undo, ALT-X=Exit" & vbCrLf & _
                                      "ALT-L=Change Logo"
        Case vbKeyL:
                If AltDown Then
                    Call MacButton("      Change Logo", frmSetup.cmdPicture, 0, 0, 170, 30, frmLogin.Source, 182, 30, 2)
                    cmdPicture_Click
                End If
        Case vbKeyE:
                If AltDown Then
                    Call MacButton("    Edit", frmSetup.cmdEdit, 0, 0, 73, 50, frmLogin.Source, 74, 0, 1)
                    cmdEdit_Click
                End If
        Case vbKeyS:
                If AltDown Then
                    Call MacButton("   Save", frmSetup.cmdSave, 0, 0, 73, 50, frmLogin.Source, 74, 0, 1)
                    cmdSave_Click
                End If
        Case vbKeyU:
                If AltDown Then
                    Call MacButton("   Undo", frmSetup.cmdUndo, 0, 0, 73, 50, frmLogin.Source, 74, 0, 1)
                    cmdUndo_Click
                End If
        Case vbKeyX:
                If AltDown Then
                    Call MacButton("    Exit", frmSetup.cmdExit, 0, 0, 73, 50, frmLogin.Source, 74, 0, 1)
                    cmdExit_Click
                End If
    End Select
End Sub

Private Sub titleBar_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call DragForm(Me)
End Sub

Private Sub Closed_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call BitBlt(frmSetup.Closed.hDC, 0, 0, 73, 50, frmLogin.Source.hDC, 18, 107, SRCCOPY)
    frmSetup.Closed.Refresh
End Sub

Private Sub Closed_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call BitBlt(frmSetup.Closed.hDC, 0, 0, 73, 50, frmLogin.Source.hDC, 0, 107, SRCCOPY)
    frmSetup.Closed.Refresh
End Sub

Private Sub Maximized_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call BitBlt(frmSetup.Maximized.hDC, 0, 0, 73, 50, frmLogin.Source.hDC, 18, 72, SRCCOPY)
    frmSetup.Maximized.Refresh
End Sub

Private Sub Maximized_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call BitBlt(frmSetup.Maximized.hDC, 0, 0, 73, 50, frmLogin.Source.hDC, 0, 72, SRCCOPY)
    frmSetup.Maximized.Refresh
End Sub

Private Sub Minimized_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call BitBlt(frmSetup.Minimized.hDC, 0, 0, 73, 50, frmLogin.Source.hDC, 18, 124, SRCCOPY)
    frmSetup.Minimized.Refresh
End Sub

Private Sub Minimized_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call BitBlt(frmSetup.Minimized.hDC, 0, 0, 73, 50, frmLogin.Source.hDC, 0, 124, SRCCOPY)
    frmSetup.Minimized.Refresh
End Sub

Private Sub Display_Fields()
    On Error Resume Next
    If datprimary.AbsolutePosition <> -1 Then
        txtField(0) = IIf(IsNull(datprimary("COMPANY_NAME")), "", datprimary("COMPANY_NAME"))
        txtField(1) = IIf(IsNull(datprimary("COMPANY_ADDRESS")), "", datprimary("COMPANY_ADDRESS"))
        txtField(2) = IIf(IsNull(datprimary("COMPANY_TELEPHONE")), "", datprimary("COMPANY_TELEPHONE"))
        txtCheck(0).Value = IIf(IsNull(datprimary("OPTION_ALAS")), 0, IIf(datprimary("OPTION_ALAS") = 0, 0, 1))
        txtCheck(1).Value = IIf(IsNull(datprimary("OPTION_HWT")), 0, IIf(datprimary("OPTION_HWT") = 0, 0, 1))
        txtCheck(2).Value = IIf(IsNull(datprimary("OPTION_HDI")), 0, IIf(datprimary("OPTION_HDI") = 0, 0, 1))
        txtCheck(3).Value = IIf(IsNull(datprimary("OPTION_E3DT")), 0, IIf(datprimary("OPTION_E3DT") = 0, 0, 1))
        txtCheck(4).Value = IIf(IsNull(datprimary("OPTION_DRCM")), 0, IIf(datprimary("OPTION_DRCM") = 0, 0, 1))
        txtCheck(5).Value = IIf(IsNull(datprimary("OPTION_DCADATS")), 0, IIf(datprimary("OPTION_DCADATS") = 0, 0, 1))
        txtCheck(6).Value = IIf(IsNull(datprimary("OPTION_DPW")), 0, IIf(datprimary("OPTION_DPW") = 0, 0, 1))
        txtField(3) = IIf(IsNull(datprimary("COMPANY_EXPIRYCOUNT")), "0", datprimary("COMPANY_EXPIRYCOUNT"))
        DatePick(0).Value = IIf(IsDate(datprimary("COMPANY_EXPIRYDATE")), datprimary("COMPANY_EXPIRYDATE"), Date)
        DatePick(1).Value = IIf(IsDate(datprimary("SELLING_CLEANUPDATE")), datprimary("SELLING_CLEANUPDATE"), Date)
        txtField(4) = IIf(IsNull(datprimary("SELLING_HEADER")), "", datprimary("SELLING_HEADER"))
        txtField(5) = IIf(IsNull(datprimary("SELLING_FOOTER")), "", datprimary("SELLING_FOOTER"))
        Call ResetLogo(Logo)
        Call DisplayLogo(datprimary("COMPANY_LOGO"))
    Else
        Clear_Fields
    End If
End Sub

Private Sub Clear_Fields()
    For i = 0 To 5
        txtField(i) = ""
    Next i
    For i = 0 To 6
        txtCheck(i) = 0
    Next i
End Sub

Private Sub Update_Fields(isNew As Boolean)
    On Error Resume Next
    Dim SaveToPicture
    datprimary.Edit
    datprimary("COMPANY_NAME") = txtField(0)
    datprimary("COMPANY_ADDRESS") = txtField(1)
    datprimary("COMPANY_TELEPHONE") = txtField(2)
    datprimary("OPTION_ALAS") = IIf(txtCheck(0).Value = 0, 0, -1)
    datprimary("OPTION_HWT") = IIf(txtCheck(1).Value = 0, 0, -1)
    datprimary("OPTION_HDI") = IIf(txtCheck(2).Value = 0, 0, -1)
    datprimary("OPTION_E3DT") = IIf(txtCheck(3).Value = 0, 0, -1)
    datprimary("OPTION_DRCM") = IIf(txtCheck(4).Value = 0, 0, -1)
    datprimary("OPTION_DCADATS") = IIf(txtCheck(5).Value = 0, 0, -1)
    datprimary("OPTION_DPW") = IIf(txtCheck(6).Value = 0, 0, -1)
    datprimary("COMPANY_EXPIRYCOUNT") = IIf(txtField(3) = "", 0, txtField(3))
    datprimary("COMPANY_EXPIRYDATE") = DatePick(0).Value
    datprimary("SELLING_CLEANUPDATE") = DatePick(1).Value
    datprimary("SELLING_HEADER") = txtField(4)
    datprimary("SELLING_FOOTER") = txtField(5)
    SaveToPicture = CopyFileToField(cdlPicture.filename, datprimary("COMPANY_LOGO"))
    datprimary.Update
End Sub

Private Sub Enable_Fields(isLock As Boolean)
    On Error Resume Next
    For i = 0 To 5
        txtField(i).Enabled = Not isLock
    Next i
    For i = 0 To 6
        txtCheck(i).Enabled = Not isLock
    Next i
    For i = 0 To 1
        DatePick(i).Enabled = Not isLock
    Next i
    Object_Tab_Trigger (Not isLock)
End Sub

Public Sub Press_Buttons(p_type As String)
    On Error Resume Next
    Select Case p_type
            Case "Edit"
                p_isediting = True
                p_save = True
                p_isadding = False
            Case "Save"
                Update_Fields (p_isadding)
                p_save = False
                p_isadding = False
                p_isediting = False
            Case "Undo"
                p_save = False
                p_isadding = False
                p_isediting = False
    End Select
    Enable_Fields (Not p_save)
    Enable_Buttons
    If Not p_isadding Then Display_Fields
End Sub

Private Sub Enable_Buttons()
    On Error Resume Next
    Dim cur_rec, fst_rec, lst_rec, rec_cnt As Integer
    Dim mark_rec As Variant
    rec_cnt = datprimary.RecordCount
    If rec_cnt > 0 Then
        If Not datprimary.BOF Or Not datprimary.EOF Then
            cur_rec = datprimary.AbsolutePosition + 1
            mark_rec = datprimary.Bookmark
        End If
        datprimary.MoveFirst
        fst_rec = datprimary.AbsolutePosition + 1
        datprimary.MoveLast
        lst_rec = datprimary.AbsolutePosition + 1
        If Not datprimary.BOF Or Not datprimary.EOF Then
            datprimary.Bookmark = mark_rec
        End If
        If fst_rec = cur_rec Then
            p_top = False
            p_prev = False
            p_next = True
            p_last = True
        End If
        If lst_rec = cur_rec Then
            p_top = True
            p_prev = True
            p_next = False
            p_last = False
        End If
        If (rec_cnt >= 0 And rec_cnt <= 1) Then
            p_top = False
            p_prev = False
            p_next = False
            p_last = False
        End If
        If cur_rec <> fst_rec And cur_rec <> lst_rec Then
            p_top = True
            p_prev = True
            p_next = True
            p_last = True
        End If
    End If
    If rec_cnt = 0 Then 'And Not p_isadding Then
        p_add = True
        p_edit = False
        p_undo = False
        p_top = False
        p_prev = False
        p_next = False
        p_last = False
        p_del = False
    End If
    If rec_cnt > 0 And (Not p_isediting And Not p_isadding) Then
        p_add = True
        p_edit = True
        p_del = True
    End If
    If Not p_isediting And Not p_isadding Then
        p_save = False
        p_undo = False
    Else
        p_save = True
        p_undo = True
        p_add = False
        p_edit = False
        p_top = False
        p_prev = False
        p_next = False
        p_last = False
        p_del = False
    End If
    cmdEdit.Enabled = p_edit
    cmdSave.Enabled = p_save
    cmdUndo.Enabled = p_undo
    cmdExit.Enabled = Not cmdSave.Enabled
End Sub
                            
Private Sub Object_Tab_Trigger(isTab As Boolean)
    On Error Resume Next
    For i = 0 To 5
        txtField(i).TabStop = isTab
    Next i
    For i = 0 To 1
        DatePick(i).TabStop = isTab
    Next i
    For i = 0 To 6
        txtCheck(i).TabStop = isTab
    Next i
End Sub

Function DisplayLogo(xField As Field)
    On Error Resume Next
    Dim DataFile As Integer, Fl As Long, Chunks As Integer
    Dim Fragment As Integer, Chunk() As Byte, i As Integer
    Const ChunkSize As Integer = 16384
    Dim MediaTemp As String
    Dim lngOffset As Long
    Dim lngTotalSize As Long
    Dim strChunk As String
    If Not IsNull(xField) Then
        MediaTemp = App.Path & "\TEMP\PICTURE.TMP"
        DataFile = 1
        Open MediaTemp For Binary Access Write As DataFile
            lngTotalSize = xField.FieldSize
            Chunks = lngTotalSize \ ChunkSize
            Fragment = lngTotalSize Mod ChunkSize
        ReDim Chunk(ChunkSize)
            Chunk() = xField.GetChunk(lngOffset, ChunkSize)
        Put DataFile, , Chunk()
            lngOffset = lngOffset + ChunkSize
        Do While lngOffset < lngTotalSize
            Chunk() = xField.GetChunk(lngOffset, ChunkSize)
            Put DataFile, , Chunk()
            lngOffset = lngOffset + ChunkSize
        Loop
        Close DataFile
        filename = MediaTemp
        Logo = LoadPicture(filename)
    End If
End Function

Function ResetLogo(xPicture As PictureBox)
    On Error Resume Next
    filename = ""
    xPicture.Picture = LoadPicture("")
    xPicture.Refresh
End Function

