VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmReceiving 
   BackColor       =   &H00CCCCCC&
   BorderStyle     =   0  'None
   Caption         =   "frmReceiving"
   ClientHeight    =   6375
   ClientLeft      =   3135
   ClientTop       =   1860
   ClientWidth     =   9330
   FillColor       =   &H00808080&
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
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   40
      Width           =   9220
      Begin VB.ListBox lstWords 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   1035
         Left            =   1800
         TabIndex        =   20
         TabStop         =   0   'False
         Top             =   1320
         Visible         =   0   'False
         Width           =   6090
      End
      Begin VB.TextBox txtWord 
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
         Left            =   3705
         TabIndex        =   0
         Text            =   "Type first few letter of the word..."
         Top             =   1080
         Width           =   4170
      End
      Begin VB.TextBox txtQuantitySold 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   300
         Left            =   3720
         MultiLine       =   -1  'True
         TabIndex        =   1
         Top             =   1485
         Width           =   1695
      End
      Begin VB.PictureBox ProductContainer 
         BackColor       =   &H00808080&
         BorderStyle     =   0  'None
         Height          =   3465
         Left            =   240
         ScaleHeight     =   3465
         ScaleWidth      =   8745
         TabIndex        =   9
         TabStop         =   0   'False
         Top             =   2640
         Width           =   8750
         Begin VB.TextBox txtDate 
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   300
            Left            =   4485
            Locked          =   -1  'True
            MultiLine       =   -1  'True
            TabIndex        =   25
            TabStop         =   0   'False
            Top             =   1920
            Width           =   1170
         End
         Begin VB.TextBox txtTotalCost 
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   300
            Left            =   4485
            Locked          =   -1  'True
            MultiLine       =   -1  'True
            TabIndex        =   15
            TabStop         =   0   'False
            Top             =   1560
            Width           =   1695
         End
         Begin VB.TextBox txtAvailableStock 
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   300
            Left            =   4485
            Locked          =   -1  'True
            MultiLine       =   -1  'True
            TabIndex        =   14
            TabStop         =   0   'False
            Top             =   1200
            Width           =   1695
         End
         Begin VB.TextBox txtUnitCost 
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   300
            Left            =   4485
            Locked          =   -1  'True
            MultiLine       =   -1  'True
            TabIndex        =   13
            TabStop         =   0   'False
            Top             =   840
            Width           =   1695
         End
         Begin VB.TextBox txtProductCode 
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   300
            Left            =   4485
            Locked          =   -1  'True
            MultiLine       =   -1  'True
            TabIndex        =   12
            TabStop         =   0   'False
            Top             =   480
            Width           =   1695
         End
         Begin VB.PictureBox cmdTag 
            Appearance      =   0  'Flat
            AutoRedraw      =   -1  'True
            BackColor       =   &H00CCCCCC&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   570
            Left            =   3120
            MouseIcon       =   "frmRecvd.frx":0000
            MousePointer    =   99  'Custom
            ScaleHeight     =   570
            ScaleWidth      =   1095
            TabIndex        =   11
            TabStop         =   0   'False
            Top             =   2640
            Width           =   1100
         End
         Begin VB.PictureBox cmdClose 
            Appearance      =   0  'Flat
            AutoRedraw      =   -1  'True
            BackColor       =   &H00CCCCCC&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   570
            Left            =   4335
            MouseIcon       =   "frmRecvd.frx":030A
            MousePointer    =   99  'Custom
            ScaleHeight     =   570
            ScaleWidth      =   1095
            TabIndex        =   10
            TabStop         =   0   'False
            Top             =   2640
            Width           =   1100
         End
         Begin VB.Label Lbls 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Last Date Received"
            BeginProperty Font 
               Name            =   "Chicago"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   225
            Index           =   1
            Left            =   2460
            TabIndex        =   24
            Top             =   1920
            Width           =   1875
         End
         Begin VB.Label Lbls 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Total Cost"
            BeginProperty Font 
               Name            =   "Chicago"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   225
            Index           =   10
            Left            =   3360
            TabIndex        =   19
            Top             =   1605
            Width           =   975
         End
         Begin VB.Label Lbls 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Available Stock"
            BeginProperty Font 
               Name            =   "Chicago"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   225
            Index           =   9
            Left            =   2835
            TabIndex        =   18
            Top             =   1245
            Width           =   1500
         End
         Begin VB.Label Lbls 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Unit Cost"
            BeginProperty Font 
               Name            =   "Chicago"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   225
            Index           =   8
            Left            =   3450
            TabIndex        =   17
            Top             =   885
            Width           =   885
         End
         Begin VB.Label Lbls 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Product Code"
            BeginProperty Font 
               Name            =   "Chicago"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   225
            Index           =   7
            Left            =   3030
            TabIndex        =   16
            Top             =   525
            Width           =   1305
         End
      End
      Begin VB.PictureBox titleBar 
         BorderStyle     =   0  'None
         Height          =   375
         Left            =   120
         ScaleHeight     =   375
         ScaleWidth      =   9015
         TabIndex        =   4
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
            MouseIcon       =   "frmRecvd.frx":0614
            MousePointer    =   99  'Custom
            ScaleHeight     =   285
            ScaleWidth      =   315
            TabIndex        =   7
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
            MouseIcon       =   "frmRecvd.frx":091E
            MousePointer    =   99  'Custom
            ScaleHeight     =   285
            ScaleWidth      =   315
            TabIndex        =   6
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
            MouseIcon       =   "frmRecvd.frx":0C28
            MousePointer    =   99  'Custom
            ScaleHeight     =   285
            ScaleWidth      =   315
            TabIndex        =   5
            TabStop         =   0   'False
            Top             =   60
            Width           =   315
         End
      End
      Begin MSComCtl2.DTPicker DTPick 
         DataField       =   "DateSold"
         DataSource      =   "datPrimaryRS"
         Height          =   300
         Left            =   3720
         TabIndex        =   2
         Top             =   1845
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   529
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
         CurrentDate     =   36527
      End
      Begin VB.Label Lbls 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Date Received"
         BeginProperty Font 
            Name            =   "Chicago"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   225
         Index           =   0
         Left            =   2190
         TabIndex        =   23
         Top             =   1845
         Width           =   1395
      End
      Begin VB.Label Lbls 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Item Description"
         BeginProperty Font 
            Name            =   "Chicago"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   225
         Index           =   6
         Left            =   1935
         TabIndex        =   22
         Top             =   1080
         Width           =   1650
      End
      Begin VB.Label Lbls 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Quantity Received"
         BeginProperty Font 
            Name            =   "Chicago"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   225
         Index           =   11
         Left            =   1800
         TabIndex        =   21
         Top             =   1485
         Width           =   1785
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
         Left            =   240
         TabIndex        =   8
         Top             =   2280
         Width           =   765
      End
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   2
      Height          =   6360
      Left            =   0
      Top             =   0
      Width           =   9300
   End
End
Attribute VB_Name = "frmReceiving"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'##############################################
'#          Coded by Walter A. Narvasa        #
'#       POS2000 - Point of Sales System      #
'#                                            #
'#           area :  frmReceiving             #
'#    description :  Receiving Transaction    #
'#         e-mail :  walter@wancom.8k.com     #
'#            url :  http://wancom.8k.com     #
'#                                            #
'##############################################

Dim p_add, p_edit, p_save, p_undo, p_top, p_prev, p_next, p_last, p_del
Dim p_isadding, p_isediting, p_isnavigate
Dim dummy As DAO.Recordset
Dim datprimary2 As DAO.Recordset
Dim invoice As DAO.Recordset
        
Private Sub cmdClose_Click()
    EditMode = False
    Screen.MousePointer = vbDefault
    Unload Me
End Sub

Private Sub cmdClose_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call MacButton("    Exit", frmReceiving.cmdClose, 0, 0, 73, 50, frmLogin.Source, 74, 0, 1)
End Sub

Private Sub cmdClose_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call MacButton("    Exit", frmReceiving.cmdClose, 0, 0, 73, 50, frmLogin.Source, 0, 0, 1)
End Sub

Private Sub cmdTag_Click()
    On Error Resume Next
    Call Receive_Stocks(frmReceiving.txtProductCode, frmReceiving.txtQuantitySold)
    txtAvailableStock = Convert_Numeric(CDbl(frmReceiving.txtAvailableStock) + CDbl(frmReceiving.txtQuantitySold), False)
    txtTotalCost = Convert_Numeric(CDbl(frmReceiving.txtUnitCost) * CDbl(frmReceiving.txtAvailableStock), True)
    txtDate = Format(DTPick.Value, "MM" & "/" & "DD" & "/" & "YYYY ")
    Call MacButton("    Tag", frmReceiving.cmdTag, 0, 0, 73, 50, frmLogin.Source, 0, 0, 1)
End Sub

Private Sub cmdTag_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call MacButton("    Tag", frmReceiving.cmdTag, 0, 0, 73, 50, frmLogin.Source, 74, 0, 1)
End Sub

Private Sub cmdTag_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call MacButton("    Tag", frmReceiving.cmdTag, 0, 0, 73, 50, frmLogin.Source, 0, 0, 1)
End Sub

Private Sub Form_Load()
    On Error Resume Next
    Call ColForm(BoxContainer, 217, 211, 213, 125)
    Call ColForm(ProductContainer, 217, 211, 213, 125)
    Call CreateMacOSTitleBar(titleBar, " Receiving Transaction ")
    Call BitBlt(frmReceiving.Closed.hDC, 0, 0, 73, 50, frmLogin.Source.hDC, 0, 107, SRCCOPY)
    frmMain.Closed.Refresh
    Call BitBlt(frmReceiving.Maximized.hDC, 0, 0, 73, 50, frmLogin.Source.hDC, 0, 72, SRCCOPY)
    frmMain.Maximized.Refresh
    Call BitBlt(frmReceiving.Minimized.hDC, 0, 0, 73, 50, frmLogin.Source.hDC, 0, 124, SRCCOPY)
    frmMain.Minimized.Refresh
    Call MacButton("    Tag", frmReceiving.cmdTag, 0, 0, 73, 50, frmLogin.Source, 0, 0, 1)
    Call MacButton("    Exit", frmReceiving.cmdClose, 0, 0, 73, 50, frmLogin.Source, 0, 0, 1)
    KeyPreview = True
    frmReceiving.txtWord.SelLength = Len(frmReceiving.txtWord.Text)
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    On Error Resume Next
    Dim AltDown
    AltDown = (Shift And vbAltMask) > 0
    Select Case KeyCode
        Case vbKeyEscape:
                Me.Hide
        Case vbKeyF1:
                frmHelp.Show
                frmHelp.Help_Values = Space(1) & vbCrLf & _
                                      "Note: The following are Receiving Transaction Key Shortcuts." & vbCrLf & _
                                      Space(1) & vbCrLf & _
                                      "ALT-T=Tag, ALT-X=Exit"
        Case vbKeyT:
                If AltDown Then
                    Call MacButton("    Tag", frmReceiving.cmdTag, 0, 0, 73, 50, frmLogin.Source, 0, 0, 1)
                End If
        Case vbKeyX:
                If AltDown Then
                    Call MacButton("    Exit", frmReceiving.cmdClose, 0, 0, 73, 50, frmLogin.Source, 0, 0, 1)
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
        Case vbKeyT:
                If AltDown Then
                    Call MacButton("    Tag", frmReceiving.cmdTag, 0, 0, 73, 50, frmLogin.Source, 74, 0, 1)
                    cmdTag_Click
                End If
        Case vbKeyX:
                If AltDown Then
                    Call MacButton("    Exit", frmReceiving.cmdClose, 0, 0, 73, 50, frmLogin.Source, 74, 0, 1)
                    cmdClose_Click
                End If
    End Select
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Screen.MousePointer = vbDefault
End Sub

Private Sub titleBar_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call DragForm(Me)
End Sub

Private Sub lstWords_Click()
    On Error Resume Next
    Call QueryReceiving(lstWords.List(lstWords.ListIndex))
End Sub

Private Sub txtWord_Change()
    On Error Resume Next
    If txtWord = "" Then
        lstWords.Visible = True
    End If
    Call QueryReceiving(txtWord.Text)
End Sub
