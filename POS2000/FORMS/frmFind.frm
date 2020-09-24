VERSION 5.00
Begin VB.Form frmFind 
   BorderStyle     =   0  'None
   Caption         =   "frmFind"
   ClientHeight    =   5130
   ClientLeft      =   4995
   ClientTop       =   1905
   ClientWidth     =   3975
   LinkTopic       =   "Form1"
   ScaleHeight     =   5130
   ScaleWidth      =   3975
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox BoxContainer 
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   0  'None
      FillColor       =   &H00C0C0C0&
      Height          =   5055
      Left            =   40
      ScaleHeight     =   5055
      ScaleWidth      =   3900
      TabIndex        =   3
      Top             =   40
      Width           =   3900
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
         Left            =   240
         TabIndex        =   0
         Text            =   "Type first few letter of the word..."
         Top             =   600
         Width           =   3455
      End
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
         Height          =   2595
         Left            =   240
         TabIndex        =   1
         Top             =   960
         Width           =   3455
      End
      Begin VB.TextBox txtExp 
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
         Height          =   495
         Left            =   240
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         TabIndex        =   2
         Top             =   3600
         Width           =   3455
      End
      Begin VB.PictureBox cmdOk 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   570
         Left            =   1440
         MouseIcon       =   "frmFind.frx":0000
         MousePointer    =   99  'Custom
         ScaleHeight     =   570
         ScaleWidth      =   1095
         TabIndex        =   5
         Top             =   4200
         Width           =   1100
      End
      Begin VB.PictureBox titleBar 
         BorderStyle     =   0  'None
         Height          =   375
         Left            =   120
         ScaleHeight     =   375
         ScaleWidth      =   3645
         TabIndex        =   4
         Top             =   90
         Width           =   3645
      End
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   2
      Height          =   5115
      Left            =   0
      Top             =   0
      Width           =   3960
   End
End
Attribute VB_Name = "frmFind"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'##############################################
'#          Coded by Walter A. Narvasa        #
'#       POS2000 - Point of Sales System      #
'#                                            #
'#           area :  frmFind                  #
'#    description :  Find Box                 #
'#         e-mail :  walter@wancom.8k.com     #
'#            url :  http://wancom.8k.com     #
'#                                            #
'##############################################

Private Sub cmdOk_Click()
    Call FindValidation
End Sub

Private Sub cmdOk_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call MacButton("     Ok", cmdOk, 0, 0, 73, 50, frmLogin.Source, 74, 0, 1)
End Sub

Private Sub cmdOk_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call MacButton("     Ok", cmdOk, 0, 0, 73, 50, frmLogin.Source, 0, 0, 1)
End Sub

Private Sub Form_Activate()
    txtWord.SelLength = Len(txtWord.Text)
End Sub

Private Sub Form_Load()
    Call CreateMacOSTitleBar(titleBar, FindType)
    Call MacButton("     Ok", cmdOk, 0, 0, 73, 50, frmLogin.Source, 0, 0, 1)
    Call ColForm(BoxContainer, 217, 211, 213, 125)
    KeyPreview = True
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyEscape:
                Me.Hide
        Case vbKeyReturn:
                cmdOk_Click
    End Select
End Sub

Private Sub titleBar_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call DragForm(Me)
End Sub

Private Sub lstWords_Click()
    Call QueryData(lstWords.List(lstWords.ListIndex), TableType, _
                    FieldType, ListType, DescType, FormType)
End Sub

Private Sub txtWord_Change()
   Call QueryData(txtWord.Text, TableType, FieldType, ListType, _
                    DescType, FormType)
End Sub

Function FindValidation()
    If TableType = "PROD_STOCKS" Then
        Call frmProduct.FindProductDes
    ElseIf TableType = "SUPPLIER" Then
        Call frmSupplier.FindSupplierCode
    ElseIf TableType = "CATEGORY" Then
        Call frmCategory.FindCategoryCode
    ElseIf TableType = "INVOICE" Then
        Call frmSelling.FindInvoiceNo
    ElseIf TableType = "USER_PASSWORD" Then
        Call frmPassword.FindUserName
    ElseIf TableType = "CODE_FILE" Then
        Call frmCodeFile.FindFileCode
    End If
    Unload Me
End Function

