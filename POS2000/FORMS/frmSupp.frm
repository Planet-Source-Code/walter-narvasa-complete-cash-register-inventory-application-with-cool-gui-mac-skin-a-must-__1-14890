VERSION 5.00
Begin VB.Form frmSupplier 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   0  'None
   Caption         =   "frmSupplier"
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
      TabIndex        =   10
      TabStop         =   0   'False
      Top             =   40
      Width           =   9220
      Begin VB.PictureBox cmdLast 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H00CCCCCC&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   300
         Left            =   2760
         MouseIcon       =   "frmSupp.frx":0000
         MousePointer    =   99  'Custom
         ScaleHeight     =   300
         ScaleWidth      =   390
         TabIndex        =   23
         TabStop         =   0   'False
         Top             =   4635
         Width           =   390
      End
      Begin VB.PictureBox cmdNext 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H00CCCCCC&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   300
         Left            =   2280
         MouseIcon       =   "frmSupp.frx":030A
         MousePointer    =   99  'Custom
         ScaleHeight     =   300
         ScaleWidth      =   390
         TabIndex        =   22
         TabStop         =   0   'False
         Top             =   4635
         Width           =   390
      End
      Begin VB.PictureBox cmdPrev 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H00CCCCCC&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   300
         Left            =   1800
         MouseIcon       =   "frmSupp.frx":0614
         MousePointer    =   99  'Custom
         ScaleHeight     =   300
         ScaleWidth      =   390
         TabIndex        =   21
         TabStop         =   0   'False
         Top             =   4635
         Width           =   390
      End
      Begin VB.PictureBox cmdTop 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H00CCCCCC&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   300
         Left            =   1320
         MouseIcon       =   "frmSupp.frx":091E
         MousePointer    =   99  'Custom
         ScaleHeight     =   300
         ScaleWidth      =   390
         TabIndex        =   20
         TabStop         =   0   'False
         Top             =   4635
         Width           =   390
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
         Left            =   3420
         MultiLine       =   -1  'True
         TabIndex        =   2
         Top             =   2880
         Width           =   3615
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
         Begin VB.PictureBox cmdFind 
            Appearance      =   0  'Flat
            AutoRedraw      =   -1  'True
            BackColor       =   &H00CCCCCC&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   570
            Left            =   6240
            MouseIcon       =   "frmSupp.frx":0C28
            MousePointer    =   99  'Custom
            ScaleHeight     =   570
            ScaleWidth      =   1095
            TabIndex        =   8
            TabStop         =   0   'False
            Top             =   240
            Width           =   1100
         End
         Begin VB.PictureBox cmdDel 
            Appearance      =   0  'Flat
            AutoRedraw      =   -1  'True
            BackColor       =   &H00CCCCCC&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   570
            Left            =   5040
            MouseIcon       =   "frmSupp.frx":0F32
            MousePointer    =   99  'Custom
            ScaleHeight     =   570
            ScaleWidth      =   1095
            TabIndex        =   7
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
            Left            =   3840
            MouseIcon       =   "frmSupp.frx":123C
            MousePointer    =   99  'Custom
            ScaleHeight     =   570
            ScaleWidth      =   1095
            TabIndex        =   6
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
            Left            =   2640
            MouseIcon       =   "frmSupp.frx":1546
            MousePointer    =   99  'Custom
            ScaleHeight     =   570
            ScaleWidth      =   1095
            TabIndex        =   5
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
            Left            =   1440
            MouseIcon       =   "frmSupp.frx":1850
            MousePointer    =   99  'Custom
            ScaleHeight     =   570
            ScaleWidth      =   1095
            TabIndex        =   4
            TabStop         =   0   'False
            Top             =   240
            Width           =   1100
         End
         Begin VB.PictureBox cmdNew 
            Appearance      =   0  'Flat
            AutoRedraw      =   -1  'True
            BackColor       =   &H00CCCCCC&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   570
            Left            =   240
            MouseIcon       =   "frmSupp.frx":1B5A
            MousePointer    =   99  'Custom
            ScaleHeight     =   570
            ScaleWidth      =   1095
            TabIndex        =   3
            TabStop         =   0   'False
            Top             =   240
            Width           =   1100
         End
         Begin VB.PictureBox cmdExit 
            Appearance      =   0  'Flat
            AutoRedraw      =   -1  'True
            BackColor       =   &H00CCCCCC&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   570
            Left            =   7440
            MouseIcon       =   "frmSupp.frx":1E64
            MousePointer    =   99  'Custom
            ScaleHeight     =   570
            ScaleWidth      =   1095
            TabIndex        =   9
            TabStop         =   0   'False
            Top             =   240
            Width           =   1100
         End
      End
      Begin VB.PictureBox titleBar 
         BorderStyle     =   0  'None
         Height          =   375
         Left            =   120
         ScaleHeight     =   375
         ScaleWidth      =   9015
         TabIndex        =   11
         TabStop         =   0   'False
         Top             =   90
         Width           =   9015
         Begin VB.PictureBox Closed 
            Appearance      =   0  'Flat
            AutoRedraw      =   -1  'True
            BackColor       =   &H00CCCCCC&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   280
            Left            =   0
            MouseIcon       =   "frmSupp.frx":216E
            MousePointer    =   99  'Custom
            ScaleHeight     =   285
            ScaleWidth      =   315
            TabIndex        =   14
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
            MouseIcon       =   "frmSupp.frx":2478
            MousePointer    =   99  'Custom
            ScaleHeight     =   285
            ScaleWidth      =   315
            TabIndex        =   13
            TabStop         =   0   'False
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
            Left            =   8430
            MouseIcon       =   "frmSupp.frx":2782
            MousePointer    =   99  'Custom
            ScaleHeight     =   285
            ScaleWidth      =   315
            TabIndex        =   12
            TabStop         =   0   'False
            Top             =   60
            Width           =   315
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
         Index           =   0
         Left            =   3420
         Locked          =   -1  'True
         TabIndex        =   0
         Top             =   2160
         Width           =   1695
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
         Left            =   3420
         MultiLine       =   -1  'True
         TabIndex        =   1
         Top             =   2520
         Width           =   3615
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
         Left            =   360
         TabIndex        =   19
         Top             =   4680
         Width           =   765
      End
      Begin VB.Label Label 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H00808080&
         BackStyle       =   0  'Transparent
         Caption         =   "Description"
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
         Left            =   2160
         TabIndex        =   18
         Top             =   2880
         Width           =   1110
      End
      Begin VB.Label Label 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H00808080&
         BackStyle       =   0  'Transparent
         Caption         =   "Supplier Code"
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
         Left            =   1935
         TabIndex        =   16
         Top             =   2160
         Width           =   1335
      End
      Begin VB.Label Label 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H00808080&
         BackStyle       =   0  'Transparent
         Caption         =   "Name"
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
         Left            =   2715
         TabIndex        =   15
         Top             =   2520
         Width           =   555
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
Attribute VB_Name = "frmSupplier"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'##############################################
'#          Coded by Walter A. Narvasa        #
'#       POS2000 - Point of Sales System      #
'#                                            #
'#           area :  frmSupplier              #
'#    description :  Supplier Masterfile      #
'#         e-mail :  walter@wancom.8k.com     #
'#            url :  http://wancom.8k.com     #
'#                                            #
'##############################################

Dim dummy As DAO.Recordset
Dim datprimary As DAO.Recordset
Dim rec_Isnew As Boolean
Dim p_add, p_edit, p_save, p_undo, p_top, p_prev, p_next, p_last, p_del
Dim p_isadding, p_isediting, p_isnavigate

Private Sub cmdDel_Click()
    On Error Resume Next
    EditMode = True
    Call MessageBox("frmSupplier", "Are you sure you want to delete Supplier Code " + datprimary("SUPCODE") + " ?", 1)
    frmMessageBox2.Show
    Call MacButton("  Delete", frmSupplier.cmdDel, 0, 0, 73, 50, frmLogin.Source, 0, 0, 1)
End Sub

Private Sub cmdDel_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call MacButton("  Delete", frmSupplier.cmdDel, 0, 0, 73, 50, frmLogin.Source, 74, 0, 1)
End Sub

Private Sub cmdDel_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call MacButton("  Delete", frmSupplier.cmdDel, 0, 0, 73, 50, frmLogin.Source, 0, 0, 1)
End Sub

Private Sub cmdEdit_Click()
'    On Error GoTo ErrorEdit
    EditMode = True
    Press_Buttons ("Edit")
    txtField(1).SetFocus
    txtField(0).TabStop = False
'ErrorEdit:
'    Call MessageBox("frmSupplier", Err.Description, 0)
End Sub

Private Sub cmdEdit_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call MacButton("    Edit", frmSupplier.cmdEdit, 0, 0, 73, 50, frmLogin.Source, 74, 0, 1)
End Sub

Private Sub cmdEdit_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call MacButton("    Edit", frmSupplier.cmdEdit, 0, 0, 73, 50, frmLogin.Source, 0, 0, 1)
End Sub

Private Sub cmdExit_Click()
    EditMode = False
    Unload Me
End Sub

Private Sub cmdExit_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call MacButton("    Exit", frmSupplier.cmdExit, 0, 0, 73, 50, frmLogin.Source, 74, 0, 1)
End Sub

Private Sub cmdExit_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call MacButton("    Exit", frmSupplier.cmdExit, 0, 0, 73, 50, frmLogin.Source, 0, 0, 1)
End Sub

Private Sub cmdFind_Click()
    On Error Resume Next
    EditMode = True
    Call FindBox(" Find by Supplier Code ", _
                    "SUPPLIER", "SUPCODE", 0, 1, "frmLogin")
    Call MacButton("   Find", frmSupplier.cmdFind, 0, 0, 73, 50, frmLogin.Source, 0, 0, 1)
End Sub

Public Sub FindSupplierCode()
    On Error Resume Next
    strs = "select * from SUPPLIER where SUPCODE = '" & frmFind.txtWord & "'"
    Set dummy = frmLogin.db.OpenRecordset(strs)
    If dummy.AbsolutePosition <> -1 Then
        datprimary.MoveFirst
        Do While Not datprimary.EOF
            If datprimary("SUPCODE") = frmFind.txtWord Then
                Exit Do
            End If
                datprimary.MoveNext
        Loop
            Display_Fields
            Enable_Buttons
    Else
        Call MessageBox("frmSupplier", "Supplier Code not found", 0)
    End If
End Sub

Private Sub cmdFind_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call MacButton("   Find", frmSupplier.cmdFind, 0, 0, 73, 50, frmLogin.Source, 74, 0, 1)
End Sub

Private Sub cmdFind_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call MacButton("   Find", frmSupplier.cmdFind, 0, 0, 73, 50, frmLogin.Source, 0, 0, 1)
End Sub

Private Sub cmdLast_Click()
    Press_Buttons ("Last")
End Sub

Private Sub cmdLast_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call MacButton("q", frmSupplier.cmdLast, 0, 0, 100, 50, frmLogin.Source, 112, 39, 3)
End Sub

Private Sub cmdLast_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call MacButton("q", frmSupplier.cmdLast, 0, 0, 100, 50, frmLogin.Source, 138, 39, 3)
End Sub

Private Sub cmdNew_Click()
    On Error Resume Next
    EditMode = False
    If datprimary.RecordCount = 0 Then
        txtField(0).Text = "1"
    Else
        datprimary.MoveLast
        txtField(0).Text = Val(datprimary("SUPCODE")) + 1
    End If
    Press_Buttons ("New")
    txtField(0).SetFocus
End Sub

Private Sub cmdNew_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call MacButton("   New", frmSupplier.cmdNew, 0, 0, 73, 50, frmLogin.Source, 74, 0, 1)
End Sub

Private Sub cmdNew_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call MacButton("   New", frmSupplier.cmdNew, 0, 0, 73, 50, frmLogin.Source, 0, 0, 1)
End Sub

Private Sub cmdNext_Click()
    Press_Buttons ("Next")
End Sub

Private Sub cmdNext_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call MacButton("u", frmSupplier.cmdNext, 0, 0, 100, 49, frmLogin.Source, 112, 39, 3)
End Sub

Private Sub cmdNext_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call MacButton("u", frmSupplier.cmdNext, 0, 0, 100, 49, frmLogin.Source, 138, 39, 3)
End Sub

Private Sub cmdOk_Click()
    BoxContainer2.Visible = False
End Sub

Private Sub cmdPrev_Click()
    Press_Buttons ("Prev")
End Sub

Private Sub cmdPrev_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call MacButton("t", frmSupplier.cmdPrev, 0, 0, 100, 50, frmLogin.Source, 112, 39, 3)
End Sub

Private Sub cmdPrev_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call MacButton("t", frmSupplier.cmdPrev, 0, 0, 100, 50, frmLogin.Source, 138, 39, 3)
End Sub

Private Sub cmdSave_Click()
    On Error Resume Next
    If EditMode = True Then
        Call MacButton("    Edit", frmSupplier.cmdEdit, 0, 0, 73, 50, frmLogin.Source, 0, 0, 1)
        Press_Buttons ("Save")
    Else
        Call MacButton("   New", frmSupplier.cmdNew, 0, 0, 73, 50, frmLogin.Source, 0, 0, 1)
        If Get_Supplier_Code Then
            Call MessageBox("frmSupplier", "Supplier Code already exists.", 0)
            frmMessageBox.SetFocus
            txtField(0) = ""
            Press_Buttons ("Undo")
        ElseIf txtField(0) = "" Then
            Call MessageBox("frmSupplier", "Supplier Code cannot be null", 0)
            frmMessageBox.SetFocus
            txtField(0) = ""
            Press_Buttons ("Undo")
        'ENABLE ONLY FOR SPECIFIC LENGTH OF 10 CHARACTERS NUMBERING
        'ElseIf Len(txtField(0)) > 1 And Len(txtField(0)) < 10 Then
        '    Call MessageBox("frmSupplier", "Supplier Code must be 10 in length", 0)
        '    frmMessageBox.SetFocus
        '    txtField(0) = ""
        '    Press_Buttons ("Undo")
        Else
            Press_Buttons ("Save")
        End If
    End If
    Call MacButton("   Save", frmSupplier.cmdSave, 0, 0, 73, 50, frmLogin.Source, 0, 0, 1)
End Sub

Private Sub cmdSave_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call MacButton("   Save", frmSupplier.cmdSave, 0, 0, 73, 50, frmLogin.Source, 74, 0, 1)
End Sub

Private Sub cmdSave_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call MacButton("   Save", frmSupplier.cmdSave, 0, 0, 73, 50, frmLogin.Source, 0, 0, 1)
End Sub

Private Sub cmdTop_Click()
    Press_Buttons ("Top")
End Sub

Private Sub cmdTop_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call MacButton("p", frmSupplier.cmdTop, 0, 0, 100, 50, frmLogin.Source, 112, 39, 3)
End Sub

Private Sub cmdTop_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call MacButton("p", frmSupplier.cmdTop, 0, 0, 100, 50, frmLogin.Source, 138, 39, 3)
End Sub

Private Sub cmdUndo_Click()
    If EditMode = True Then
        Call MacButton("    Edit", frmSupplier.cmdEdit, 0, 0, 73, 50, frmLogin.Source, 0, 0, 1)
    Else
        Call MacButton("   New", frmSupplier.cmdNew, 0, 0, 73, 50, frmLogin.Source, 0, 0, 1)
    End If
    Press_Buttons ("Undo")
    Call MacButton("   Undo", frmSupplier.cmdUndo, 0, 0, 73, 50, frmLogin.Source, 0, 0, 1)
End Sub

Private Sub cmdUndo_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call MacButton("   Undo", frmSupplier.cmdUndo, 0, 0, 73, 50, frmLogin.Source, 74, 0, 1)
End Sub

Private Sub cmdUndo_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call MacButton("   Undo", frmSupplier.cmdUndo, 0, 0, 73, 50, frmLogin.Source, 0, 0, 1)
End Sub

Private Sub Form_Activate()
    On Error Resume Next
    If EditMode = False Then
        strs = "select * from SUPPLIER order by SUPCODE"
        Set datprimary = frmLogin.db.OpenRecordset(strs)
        If datprimary.AbsolutePosition <> -1 Then
            Display_Fields
        End If
            Enable_Fields (True)
            Object_Tab_Trigger (False)
            Enable_Buttons
            Enable_Buttons
            rec_Isnew = False
        If cmdNew.Enabled Then cmdNew.SetFocus
    End If
End Sub

Private Sub Form_Load()
    On Error Resume Next
    Call ColForm(BoxContainer, 217, 211, 213, 125)
    Call ColForm(ButtonContainer, 217, 211, 213, 125)
    Call CreateMacOSTitleBar(titleBar, " Supplier Masterfile ")
    Call BitBlt(frmSupplier.Closed.hDC, 0, 0, 73, 50, frmLogin.Source.hDC, 0, 107, SRCCOPY)
    frmSupplier.Closed.Refresh
    Call BitBlt(frmSupplier.Maximized.hDC, 0, 0, 73, 50, frmLogin.Source.hDC, 0, 72, SRCCOPY)
    frmSupplier.Maximized.Refresh
    Call BitBlt(frmSupplier.Minimized.hDC, 0, 0, 73, 50, frmLogin.Source.hDC, 0, 124, SRCCOPY)
    frmSupplier.Minimized.Refresh
    KeyPreview = True
    Call MacButton("   New", frmSupplier.cmdNew, 0, 0, 73, 50, frmLogin.Source, 0, 0, 1)
    Call MacButton("    Edit", frmSupplier.cmdEdit, 0, 0, 73, 50, frmLogin.Source, 0, 0, 1)
    Call MacButton("   Save", frmSupplier.cmdSave, 0, 0, 73, 50, frmLogin.Source, 0, 0, 1)
    Call MacButton("   Undo", frmSupplier.cmdUndo, 0, 0, 73, 50, frmLogin.Source, 0, 0, 1)
    Call MacButton("  Delete", frmSupplier.cmdDel, 0, 0, 73, 50, frmLogin.Source, 0, 0, 1)
    Call MacButton("   Find", frmSupplier.cmdFind, 0, 0, 73, 50, frmLogin.Source, 0, 0, 1)
    Call MacButton("    Exit", frmSupplier.cmdExit, 0, 0, 73, 50, frmLogin.Source, 0, 0, 1)
    Call MacButton("p", frmSupplier.cmdTop, 0, 0, 100, 50, frmLogin.Source, 138, 39, 3)
    Call MacButton("u", frmSupplier.cmdNext, 0, 0, 100, 50, frmLogin.Source, 138, 39, 3)
    Call MacButton("t", frmSupplier.cmdPrev, 0, 0, 100, 50, frmLogin.Source, 138, 39, 3)
    Call MacButton("q", frmSupplier.cmdLast, 0, 0, 100, 50, frmLogin.Source, 138, 39, 3)
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    On Error Resume Next
    Dim AltDown
    AltDown = (Shift And vbAltMask) > 0
    Select Case KeyCode
        Case vbKeyEscape:
                Me.Hide
        Case vbKeyN:
                If AltDown Then
                    Call MacButton("   New", frmSupplier.cmdNew, 0, 0, 73, 50, frmLogin.Source, 0, 0, 1)
                End If
        Case vbKeyE:
                If AltDown Then
                    Call MacButton("    Edit", frmSupplier.cmdEdit, 0, 0, 73, 50, frmLogin.Source, 0, 0, 1)
                End If
        Case vbKeyS:
                If AltDown Then
                    Call MacButton("   Save", frmSupplier.cmdSave, 0, 0, 73, 50, frmLogin.Source, 0, 0, 1)
                End If
        Case vbKeyU:
                If AltDown Then
                    Call MacButton("   Undo", frmSupplier.cmdUndo, 0, 0, 73, 50, frmLogin.Source, 0, 0, 1)
                End If
        Case vbKeyD:
                If AltDown Then
                    Call MacButton("  Delete", frmSupplier.cmdDel, 0, 0, 73, 50, frmLogin.Source, 0, 0, 1)
                End If
        Case vbKeyF:
                If AltDown Then
                    Call MacButton("   Find", frmSupplier.cmdFind, 0, 0, 73, 50, frmLogin.Source, 0, 0, 1)
                End If
        Case vbKeyX:
                If AltDown Then
                    Call MacButton("    Exit", frmSupplier.cmdExit, 0, 0, 73, 50, frmLogin.Source, 0, 0, 1)
                End If
        Case vbKeyLeft:
                If AltDown Then
                    Call MacButton("t", frmSupplier.cmdPrev, 0, 0, 100, 50, frmLogin.Source, 138, 39, 3)
                End If
        Case vbKeyRight:
                If AltDown Then
                    Call MacButton("u", frmSupplier.cmdNext, 0, 0, 100, 49, frmLogin.Source, 138, 39, 3)
                End If
        Case vbKeyUp:
                If AltDown Then
                    Call MacButton("p", frmSupplier.cmdTop, 0, 0, 100, 50, frmLogin.Source, 138, 39, 3)
                End If
        Case vbKeyDown:
                If AltDown Then
                    Call MacButton("q", frmSupplier.cmdLast, 0, 0, 100, 50, frmLogin.Source, 138, 39, 3)
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
                                      "Note: The following are Supplier Masterfile Key Shortcuts." & vbCrLf & _
                                      Space(1) & vbCrLf & _
                                      "ALT-N=New, ALT-E=Edit, ALT-S=Save, ALT-U=Undo, ALT-D=Delete" & vbCrLf & _
                                      "ALT-F=Find, ALT-X=Exit" & vbCrLf & _
                                      Space(1) & vbCrLf & _
                                      "Left Arrow=Previous Records, Right Arrow=Next Records" & vbCrLf & _
                                      "Top Arrow=Top Record, Down Arrow=Last Record"
        Case vbKeyN:
                If AltDown Then
                    Call MacButton("   New", frmSupplier.cmdNew, 0, 0, 73, 50, frmLogin.Source, 74, 0, 1)
                    cmdNew_Click
                End If
        Case vbKeyE:
                If AltDown Then
                    Call MacButton("    Edit", frmSupplier.cmdEdit, 0, 0, 73, 50, frmLogin.Source, 74, 0, 1)
                    cmdEdit_Click
                End If
        Case vbKeyS:
                If AltDown Then
                    Call MacButton("   Save", frmSupplier.cmdSave, 0, 0, 73, 50, frmLogin.Source, 74, 0, 1)
                    cmdSave_Click
                End If
        Case vbKeyU:
                If AltDown Then
                    Call MacButton("   Undo", frmSupplier.cmdUndo, 0, 0, 73, 50, frmLogin.Source, 74, 0, 1)
                    cmdUndo_Click
                End If
        Case vbKeyD:
                If AltDown Then
                    Call MacButton("  Delete", frmSupplier.cmdDel, 0, 0, 73, 50, frmLogin.Source, 74, 0, 1)
                    cmdDel_Click
                End If
        Case vbKeyF:
                If AltDown Then
                    Call MacButton("   Find", frmSupplier.cmdFind, 0, 0, 73, 50, frmLogin.Source, 74, 0, 1)
                    cmdFind_Click
                End If
        Case vbKeyX:
                If AltDown Then
                    Call MacButton("    Exit", frmSupplier.cmdExit, 0, 0, 73, 50, frmLogin.Source, 74, 0, 1)
                    cmdExit_Click
                End If
        Case vbKeyLeft:
                If AltDown Then
                    Call MacButton("t", frmSupplier.cmdPrev, 0, 0, 100, 50, frmLogin.Source, 112, 39, 3)
                    cmdPrev_Click
                End If
        Case vbKeyRight:
                If AltDown Then
                    Call MacButton("u", frmSupplier.cmdNext, 0, 0, 100, 49, frmLogin.Source, 112, 39, 3)
                    cmdNext_Click
                End If
        Case vbKeyUp:
                If AltDown Then
                    Call MacButton("p", frmSupplier.cmdTop, 0, 0, 100, 50, frmLogin.Source, 112, 39, 3)
                    cmdTop_Click
                End If
        Case vbKeyDown:
                If AltDown Then
                    Call MacButton("q", frmSupplier.cmdLast, 0, 0, 100, 50, frmLogin.Source, 112, 39, 3)
                    cmdLast_Click
                End If
    End Select
End Sub

Private Sub titleBar_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call DragForm(Me)
End Sub

Private Sub Closed_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call BitBlt(frmSupplier.Closed.hDC, 0, 0, 73, 50, frmLogin.Source.hDC, 18, 107, SRCCOPY)
    frmSupplier.Closed.Refresh
End Sub

Private Sub Closed_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call BitBlt(frmSupplier.Closed.hDC, 0, 0, 73, 50, frmLogin.Source.hDC, 0, 107, SRCCOPY)
    frmSupplier.Closed.Refresh
End Sub

Private Sub Maximized_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call BitBlt(frmSupplier.Maximized.hDC, 0, 0, 73, 50, frmLogin.Source.hDC, 18, 72, SRCCOPY)
    frmSupplier.Maximized.Refresh
End Sub

Private Sub Maximized_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call BitBlt(frmSupplier.Maximized.hDC, 0, 0, 73, 50, frmLogin.Source.hDC, 0, 72, SRCCOPY)
    frmSupplier.Maximized.Refresh
End Sub

Private Sub Minimized_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call BitBlt(frmSupplier.Minimized.hDC, 0, 0, 73, 50, frmLogin.Source.hDC, 18, 124, SRCCOPY)
    frmSupplier.Minimized.Refresh
End Sub

Private Sub Minimized_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call BitBlt(frmSupplier.Minimized.hDC, 0, 0, 73, 50, frmLogin.Source.hDC, 0, 124, SRCCOPY)
    frmSupplier.Minimized.Refresh
End Sub

Private Sub Display_Fields()
    On Error Resume Next
    If datprimary.AbsolutePosition <> -1 Then
        txtField(0) = IIf(IsNull(datprimary("SUPCODE")), "", datprimary("SUPCODE"))
        txtField(1) = IIf(IsNull(datprimary("SUPNAME")), "", datprimary("SUPNAME"))
        txtField(2) = IIf(IsNull(datprimary("SUPDES")), "", datprimary("SUPDES"))
    Else
        Clear_Fields
    End If
End Sub

Private Sub Clear_Fields()
    ' ENABLE ONLY TO BLANK FIELD txtField(0)
    'txtField(0) = ""
    txtField(1) = ""
    txtField(2) = ""
End Sub


Private Sub Update_Fields(isNew As Boolean)
    On Error Resume Next
    If isNew Then
        datprimary.AddNew
    Else
        datprimary.Edit
    End If
        datprimary("SUPCODE") = txtField(0)
        datprimary("SUPNAME") = txtField(1)
        datprimary("SUPDES") = txtField(2)
        datprimary.Update
    If isNew Then datprimary.MoveLast
End Sub

Private Sub Enable_Fields(isLock As Boolean)
    On Error Resume Next
    For i = 0 To 2
        txtField(i).Enabled = Not isLock
    Next i
    If p_isadding Then
        txtField(0).Locked = isLock
        txtField(0).TabStop = True
    End If
        Object_Tab_Trigger (Not isLock)
End Sub

Public Sub Press_Buttons(p_type As String)
    On Error Resume Next
    Select Case p_type
            Case "New"
                Clear_Fields
                p_save = True
                p_isadding = True
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
            Case "Top"
                p_save = False
                p_isadding = False
                p_isediting = False
                datprimary.MoveFirst
            Case "Prev"
                p_save = False
                p_isadding = False
                p_isediting = False
                datprimary.MovePrevious
            Case "Next"
                p_save = False
                p_isadding = False
                p_isediting = False
                datprimary.MoveNext
            Case "Last"
                p_save = False
                p_isadding = False
                p_isediting = False
                datprimary.MoveLast
            Case "Delete"
                p_save = False
                p_isadding = False
                p_isediting = False
                p_isdeleting = True
                p_isnavigate = False
                With datprimary
                    .Delete
                    .MoveNext
                    p_isnavigate = True
                    If .RecordCount > 0 And .EOF Then .MoveLast
                End With
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
    cmdNew.Enabled = p_add
    cmdEdit.Enabled = p_edit
    cmdSave.Enabled = p_save
    cmdUndo.Enabled = p_undo
    cmdTop.Enabled = p_top
    cmdPrev.Enabled = p_prev
            cmdNext.Enabled = p_next
    cmdLast.Enabled = p_last
    cmdDel.Enabled = p_del
    If p_del Then
        cmdFind.Enabled = IIf(rec_cnt > 1, True, False)
    Else
        cmdFind.Enabled = False
    End If
        cmdExit.Enabled = Not cmdSave.Enabled
End Sub
                            
Private Sub Object_Tab_Trigger(isTab As Boolean)
    On Error Resume Next
    txtField(1).TabStop = isTab
    txtField(2).TabStop = isTab
End Sub

Function Get_Supplier_Code() As Boolean
    On Error Resume Next
    strs = "select * from SUPPLIER where SUPCODE = '" & txtField(0) & "'"
    Set dummy = frmLogin.db.OpenRecordset(strs)
    If dummy.AbsolutePosition <> -1 Then
        Get_Supplier_Code = True
    Else
        Get_Supplier_Code = False
    End If
End Function
