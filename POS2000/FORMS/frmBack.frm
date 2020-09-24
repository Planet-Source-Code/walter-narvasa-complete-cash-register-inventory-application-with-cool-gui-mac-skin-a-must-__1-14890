VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmBackup 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00CCCCCC&
   BorderStyle     =   0  'None
   Caption         =   "frmBackup"
   ClientHeight    =   2985
   ClientLeft      =   4500
   ClientTop       =   4005
   ClientWidth     =   5565
   LinkTopic       =   "Form1"
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
      Begin MSComDlg.CommonDialog Dialog 
         Left            =   480
         Top             =   1800
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin VB.PictureBox cmdRestore 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H00CCCCCC&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   450
         Left            =   1560
         MouseIcon       =   "frmBack.frx":0000
         MousePointer    =   99  'Custom
         ScaleHeight     =   450
         ScaleWidth      =   2355
         TabIndex        =   4
         TabStop         =   0   'False
         Top             =   1320
         Width           =   2355
      End
      Begin VB.PictureBox cmdBackup 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H00CCCCCC&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   450
         Left            =   1560
         MouseIcon       =   "frmBack.frx":030A
         MousePointer    =   99  'Custom
         ScaleHeight     =   450
         ScaleWidth      =   2355
         TabIndex        =   3
         TabStop         =   0   'False
         Top             =   720
         Width           =   2355
      End
      Begin VB.PictureBox cmdOk 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   570
         Left            =   2160
         MouseIcon       =   "frmBack.frx":0614
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
Attribute VB_Name = "frmBackup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'##############################################
'#          Coded by Walter A. Narvasa        #
'#       POS2000 - Point of Sales System      #
'#                                            #
'#           area :  frmBackup                #
'#    description :  Backup/Restore Files     #
'#         e-mail :  walter@wancom.8k.com     #
'#            url :  http://wancom.8k.com     #
'#                                            #
'##############################################

Private Sub cmdBackup_Click()
    On Error Resume Next
    Call MacButton("   Backup Database", frmBackup.cmdBackup, 0, 0, 170, 30, frmLogin.Source, 182, 30, 2)
    Dialog.filename = ""
    Dialog.Filter = "Backup files (*.bck) |*.bck|"
    Dialog.ShowSave
    If Dialog.filename <> "" Then
        If FileLen(App.Path & "\DATABASE\POSDATA.MDB") > 1210000 And Mid(Dialog.filename, 1, 1) = "P" Then
            Call MessageBox("frmBackup", "Database size is greater than diskette capacity.", 0)
            frmMessageBox.Show
        Else
            frmLogin.db.Close
            FileCopy (App.Path & "\DATABASE\POSDATA.MDB "), Dialog.filename
            Call MessageBox("frmBackup", "Backup Completed!!", 0)
            frmMessageBox.Show
            Set frmLogin.db = Workspaces(0).OpenDatabase(App.Path & "\DATABASE\POSDATA.MDB")
        End If
    End If
    Call MacButton("   Backup Database", frmBackup.cmdBackup, 0, 0, 170, 30, frmLogin.Source, 147, 0, 2)
End Sub

Private Sub cmdRestore_Click()
    On Error Resume Next
    Call MacButton("   Restore Database", frmBackup.cmdRestore, 0, 0, 170, 30, frmLogin.Source, 182, 30, 2)
    Dialog.ShowOpen
    If Dialog.filename <> "" Then
        Call MessageBox("frmBackup", "Are you sure you want to restore the database", 1)
        frmMessageBox2.Show
    End If
    Call MacButton("   Restore Database", frmBackup.cmdRestore, 0, 0, 170, 30, frmLogin.Source, 147, 0, 2)
End Sub

Public Sub Restore_Database()
    On Error Resume Next
    frmLogin.db.Close
    FileCopy Dialog.filename, "POSDATA.MDB"
    Set frmLogin.db = Workspaces(0).OpenDatabase(App.Path & "\DATABASE\POSDATA.MDB")
    Call MessageBox("frmBackup", "Database has been fully restored!!", 0)
    frmMessageBox.Show
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
    Call CreateMacOSTitleBar(titleBar, " Backup/Restore Files ")
    Call MacButton("     Ok", cmdOk, 0, 0, 73, 50, frmLogin.Source, 0, 0, 1)
    Call ColForm(BoxContainer, 217, 211, 213, 125)
    Call MacButton("   Backup Database", frmBackup.cmdBackup, 0, 0, 170, 30, frmLogin.Source, 147, 0, 2)
    Call MacButton("   Restore Database", frmBackup.cmdRestore, 0, 0, 170, 30, frmLogin.Source, 147, 0, 2)
    KeyPreview = True
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn:
                Call MessageValidation
    End Select
End Sub

Private Sub titleBar_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call DragForm(Me)
End Sub

Function MessageValidation()
    On Error Resume Next
    If CmdType = "frmLogin" Then
        frmLogin.txtUserName.SetFocus
    ElseIf CmdType = "frmProduct" Then
        frmProduct.txtField(1).SetFocus
    ElseIf CmdType = "frmSupplier" Then
        frmSupplier.txtField(0).SetFocus
    ElseIf CmdType = "frmCategory" Then
        frmCategory.txtField(0).SetFocus
    ElseIf CmdType = "frmSelling" Then
        frmSelling.txtField(0).SetFocus
    ElseIf CmdType = "frmPassword" Then
        frmPassword.txtField(0).SetFocus
    ElseIf CmdType = "frmCodeFile" Then
        frmCodeFile.txtField(0).SetFocus
    End If
        Unload Me
End Function

