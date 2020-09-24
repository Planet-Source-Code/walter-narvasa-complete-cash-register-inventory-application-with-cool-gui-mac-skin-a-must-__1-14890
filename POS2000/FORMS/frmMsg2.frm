VERSION 5.00
Begin VB.Form frmMessageBox2 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00CCCCCC&
   BorderStyle     =   0  'None
   Caption         =   "frmMessageBox"
   ClientHeight    =   2985
   ClientLeft      =   4500
   ClientTop       =   4005
   ClientWidth     =   5565
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
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
      TabIndex        =   0
      Top             =   40
      Width           =   5460
      Begin VB.PictureBox cmdCancel 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   570
         Left            =   2880
         MouseIcon       =   "frmMsg2.frx":0000
         MousePointer    =   99  'Custom
         ScaleHeight     =   570
         ScaleWidth      =   1095
         TabIndex        =   4
         Top             =   2040
         Width           =   1100
      End
      Begin VB.PictureBox cmdOk 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   570
         Left            =   1560
         MouseIcon       =   "frmMsg2.frx":030A
         MousePointer    =   99  'Custom
         ScaleHeight     =   570
         ScaleWidth      =   1095
         TabIndex        =   1
         Top             =   2040
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
      Begin VB.Label MessageList 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Put Message Here!!"
         BeginProperty Font 
            Name            =   "Chicago"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   825
         Left            =   720
         TabIndex        =   3
         Top             =   870
         Width           =   3915
         WordWrap        =   -1  'True
      End
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
Attribute VB_Name = "frmMessageBox2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'##############################################
'#          Coded by Walter A. Narvasa        #
'#       POS2000 - Point of Sales System      #
'#                                            #
'#           area :  frmLogin                 #
'#    description :  Message Box2 (Ok/Cancel) #
'#         e-mail :  walter@wancom.8k.com     #
'#            url :  http://wancom.8k.com     #
'#                                            #
'##############################################

Private Sub cmdCancel_Click()
    Call CancelValidation
End Sub

Private Sub cmdCancel_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call MacButton("  Cancel", cmdCancel, 0, 0, 73, 50, frmLogin.Source, 74, 0, 1)
End Sub

Private Sub cmdCancel_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call MacButton("  Cancel", cmdCancel, 0, 0, 73, 50, frmLogin.Source, 0, 0, 1)
End Sub

Private Sub cmdOk_Click()
    Call MessageValidation
End Sub

Private Sub cmdOk_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call MacButton("     Ok", cmdOk, 0, 0, 73, 50, frmLogin.Source, 74, 0, 1)
End Sub

Private Sub cmdOk_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call MacButton("     Ok", cmdOk, 0, 0, 73, 50, frmLogin.Source, 0, 0, 1)
End Sub

Private Sub Form_Load()
    Call CreateMacOSTitleBar(titleBar, " Readme Message ")
    Call MacButton("     Ok", cmdOk, 0, 0, 73, 50, frmLogin.Source, 0, 0, 1)
    Call MacButton("  Cancel", cmdCancel, 0, 0, 73, 50, frmLogin.Source, 0, 0, 1)
    Call ColForm(BoxContainer, 217, 211, 213, 125)
    KeyPreview = True
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn:
                Call MessageValidation
        Case vbKeyEscape:
                cmdCancel_Click
    End Select
End Sub

Private Sub titleBar_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call DragForm(Me)
End Sub

Function MessageValidation()
    If CmdType = "frmLogin" Then
        frmLogin.txtUserName.SetFocus
    ElseIf CmdType = "frmProduct" Then
        frmProduct.Press_Buttons ("Delete")
    ElseIf CmdType = "frmSupplier" Then
        frmSupplier.Press_Buttons ("Delete")
    ElseIf CmdType = "frmCategory" Then
        frmCategory.Press_Buttons ("Delete")
    ElseIf CmdType = "frmSelling" Then
        frmSelling.Press_Button ("DELETE")
    ElseIf CmdType = "frmPassword" Then
        frmPassword.Press_Buttons ("Delete")
    ElseIf CmdType = "frmCodeFile" Then
        frmCodeFile.Press_Buttons ("Delete")
    ElseIf CmdType = "frmMain" Then
        frmMain.Operation_CleanUp
    ElseIf CmdType = "frmBackup" Then
        frmBackup.Restore_Database
    End If
        Unload Me
End Function

Function CancelValidation()
    If CmdType = "frmLogin" Then
    ElseIf CmdType = "frmProduct" Then
    ElseIf CmdType = "frmSupplier" Then
    ElseIf CmdType = "frmCategory" Then
    ElseIf CmdType = "frmSelling" Then
    ElseIf CmdType = "frmPassword" Then
    ElseIf CmdType = "frmCodeFile" Then
    ElseIf CmdType = "frmMain" Then
        frmMain.Timer1.Enabled = False
    ElseIf CmdType = "frmBackup" Then
    End If
        Unload Me
End Function

