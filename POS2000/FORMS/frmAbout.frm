VERSION 5.00
Begin VB.Form frmAbout 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00CCCCCC&
   BorderStyle     =   0  'None
   Caption         =   "frmAbout"
   ClientHeight    =   3990
   ClientLeft      =   4500
   ClientTop       =   4005
   ClientWidth     =   5565
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   3990
   ScaleWidth      =   5565
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox BoxContainer 
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   0  'None
      FillColor       =   &H00C0C0C0&
      Height          =   3890
      Left            =   40
      ScaleHeight     =   3885
      ScaleWidth      =   5460
      TabIndex        =   0
      Top             =   40
      Width           =   5460
      Begin VB.PictureBox AuthorContainer 
         BackColor       =   &H00808080&
         BorderStyle     =   0  'None
         FillColor       =   &H00C0C0C0&
         Height          =   1575
         Left            =   360
         ScaleHeight     =   1575
         ScaleWidth      =   1500
         TabIndex        =   4
         Top             =   720
         Width           =   1500
         Begin VB.Image Image1 
            Height          =   1365
            Left            =   120
            Picture         =   "frmAbout.frx":0000
            Top             =   120
            Width           =   1305
         End
      End
      Begin VB.PictureBox cmdOk 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   570
         Left            =   2280
         MouseIcon       =   "frmAbout.frx":5E1A
         MousePointer    =   99  'Custom
         ScaleHeight     =   570
         ScaleWidth      =   1095
         TabIndex        =   1
         Top             =   3120
         Width           =   1100
      End
      Begin VB.PictureBox titleBar 
         BorderStyle     =   0  'None
         Height          =   375
         Left            =   120
         ScaleHeight     =   375
         ScaleWidth      =   5205
         TabIndex        =   2
         Top             =   90
         Width           =   5200
      End
      Begin VB.Label Company_Name 
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Chicago"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Left            =   2160
         TabIndex        =   8
         Top             =   2400
         Width           =   3075
         WordWrap        =   -1  'True
      End
      Begin VB.Label MessageList 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "License to:"
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
         Left            =   2160
         TabIndex        =   7
         Top             =   2160
         Width           =   2115
         WordWrap        =   -1  'True
      End
      Begin VB.Label MessageList 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Copyright (c) 2000 Walter A. Narvasa  -Author"
         BeginProperty Font 
            Name            =   "Chicago"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   675
         Index           =   2
         Left            =   2160
         TabIndex        =   6
         Top             =   1320
         Width           =   2475
         WordWrap        =   -1  'True
      End
      Begin VB.Label MessageList 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "for Pharmaceutical"
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
         Left            =   2160
         TabIndex        =   5
         Top             =   960
         Width           =   2115
         WordWrap        =   -1  'True
      End
      Begin VB.Label MessageList 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Point-of-Sales 2000"
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
         Left            =   2160
         TabIndex        =   3
         Top             =   720
         Width           =   2115
         WordWrap        =   -1  'True
      End
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   2
      Height          =   3950
      Left            =   15
      Top             =   15
      Width           =   5520
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'##############################################
'#          Coded by Walter A. Narvasa        #
'#       POS2000 - Point of Sales System      #
'#                                            #
'#           area :  frmAbout                 #
'#    description :  About the Author         #
'#         e-mail :  walter@wancom.8k.com     #
'#            url :  http://wancom.8k.com     #
'#                                            #
'##############################################


Private Sub cmdOk_Click()
    Unload Me
End Sub

Private Sub cmdOk_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call MacButton("     Ok", cmdOk, 0, 0, 73, 50, frmLogin.Source, 74, 0, 1)
End Sub

Private Sub cmdOk_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call MacButton("     Ok", cmdOk, 0, 0, 73, 50, frmLogin.Source, 0, 0, 1)
End Sub

Private Sub Form_Load()
    Call CreateMacOSTitleBar(titleBar, " About ")
    Call MacButton("     Ok", cmdOk, 0, 0, 73, 50, frmLogin.Source, 0, 0, 1)
    Call ColForm(BoxContainer, 217, 211, 213, 125)
    Call ColForm(AuthorContainer, 217, 211, 213, 125)
    KeyPreview = True
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn:
                Unload Me
    End Select
End Sub

Private Sub titleBar_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call DragForm(Me)
End Sub

