VERSION 5.00
Begin VB.Form frmHelp 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00CCCCCC&
   BorderStyle     =   0  'None
   Caption         =   "frmHelp"
   ClientHeight    =   4395
   ClientLeft      =   4500
   ClientTop       =   4005
   ClientWidth     =   6765
   LinkTopic       =   "Form1"
   ScaleHeight     =   4395
   ScaleWidth      =   6765
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox BoxContainer 
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   0  'None
      FillColor       =   &H00C0C0C0&
      Height          =   4290
      Left            =   40
      ScaleHeight     =   4290
      ScaleWidth      =   6660
      TabIndex        =   0
      Top             =   40
      Width           =   6660
      Begin VB.PictureBox cmdOk 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   570
         Left            =   2760
         MouseIcon       =   "frmHelp.frx":0000
         MousePointer    =   99  'Custom
         ScaleHeight     =   570
         ScaleWidth      =   1095
         TabIndex        =   1
         Top             =   3480
         Width           =   1100
      End
      Begin VB.PictureBox titleBar 
         BorderStyle     =   0  'None
         Height          =   375
         Left            =   120
         ScaleHeight     =   375
         ScaleWidth      =   6405
         TabIndex        =   2
         Top             =   90
         Width           =   6405
      End
      Begin VB.Label Help_Values 
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
         Height          =   2745
         Left            =   240
         TabIndex        =   4
         Top             =   600
         Width           =   6195
         WordWrap        =   -1  'True
      End
      Begin VB.Label MessageList 
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
         Height          =   2145
         Index           =   0
         Left            =   240
         TabIndex        =   3
         Top             =   720
         Width           =   6195
         WordWrap        =   -1  'True
      End
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   2
      Height          =   4350
      Left            =   15
      Top             =   15
      Width           =   6720
   End
End
Attribute VB_Name = "frmHelp"
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
    Call CreateMacOSTitleBar(titleBar, " Help ")
    Call MacButton("     Ok", cmdOk, 0, 0, 73, 50, frmLogin.Source, 0, 0, 1)
    Call ColForm(BoxContainer, 217, 211, 213, 125)
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

