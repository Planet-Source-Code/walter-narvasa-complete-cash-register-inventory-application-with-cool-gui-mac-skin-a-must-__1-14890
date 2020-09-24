VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmProduct 
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
      TabIndex        =   23
      TabStop         =   0   'False
      Top             =   40
      Width           =   9220
      Begin VB.ComboBox txtCombo 
         Enabled         =   0   'False
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
         Index           =   1
         Left            =   1980
         TabIndex        =   9
         Text            =   "Combo1"
         Top             =   4080
         Width           =   2055
      End
      Begin VB.TextBox txtField 
         Alignment       =   1  'Right Justify
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
         Left            =   1980
         MultiLine       =   -1  'True
         TabIndex        =   7
         Top             =   3360
         Width           =   1695
      End
      Begin VB.ComboBox txtCombo 
         Enabled         =   0   'False
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
         Index           =   0
         Left            =   1980
         TabIndex        =   3
         Text            =   "Combo1"
         Top             =   1920
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
         Height          =   300
         Index           =   7
         Left            =   1980
         MultiLine       =   -1  'True
         TabIndex        =   2
         Top             =   1560
         Width           =   3615
      End
      Begin VB.PictureBox PictureContainer 
         BackColor       =   &H00808080&
         BorderStyle     =   0  'None
         Height          =   4335
         Left            =   5760
         ScaleHeight     =   4335
         ScaleWidth      =   3255
         TabIndex        =   36
         TabStop         =   0   'False
         Top             =   600
         Width           =   3255
         Begin MSComDlg.CommonDialog cdlPicture 
            Left            =   240
            Top             =   3360
            _ExtentX        =   847
            _ExtentY        =   847
            _Version        =   393216
         End
         Begin VB.PictureBox Photo 
            BackColor       =   &H00000000&
            BorderStyle     =   0  'None
            FillColor       =   &H00808080&
            ForeColor       =   &H00808080&
            Height          =   2895
            Left            =   240
            ScaleHeight     =   2895
            ScaleWidth      =   2835
            TabIndex        =   37
            TabStop         =   0   'False
            Top             =   240
            Width           =   2830
         End
         Begin VB.PictureBox cmdPicture 
            Appearance      =   0  'Flat
            AutoRedraw      =   -1  'True
            BackColor       =   &H00CCCCCC&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   450
            Left            =   480
            MouseIcon       =   "frmProd.frx":0000
            MousePointer    =   99  'Custom
            ScaleHeight     =   450
            ScaleWidth      =   2355
            TabIndex        =   22
            TabStop         =   0   'False
            Top             =   3720
            Width           =   2355
         End
         Begin VB.PictureBox cmdTop 
            Appearance      =   0  'Flat
            AutoRedraw      =   -1  'True
            BackColor       =   &H00CCCCCC&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   300
            Left            =   720
            MouseIcon       =   "frmProd.frx":030A
            MousePointer    =   99  'Custom
            ScaleHeight     =   300
            ScaleWidth      =   390
            TabIndex        =   18
            TabStop         =   0   'False
            Top             =   3240
            Width           =   390
         End
         Begin VB.PictureBox cmdPrev 
            Appearance      =   0  'Flat
            AutoRedraw      =   -1  'True
            BackColor       =   &H00CCCCCC&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   300
            Left            =   1200
            MouseIcon       =   "frmProd.frx":0614
            MousePointer    =   99  'Custom
            ScaleHeight     =   300
            ScaleWidth      =   390
            TabIndex        =   19
            TabStop         =   0   'False
            Top             =   3240
            Width           =   390
         End
         Begin VB.PictureBox cmdNext 
            Appearance      =   0  'Flat
            AutoRedraw      =   -1  'True
            BackColor       =   &H00CCCCCC&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   300
            Left            =   1680
            MouseIcon       =   "frmProd.frx":091E
            MousePointer    =   99  'Custom
            ScaleHeight     =   300
            ScaleWidth      =   390
            TabIndex        =   20
            TabStop         =   0   'False
            Top             =   3240
            Width           =   390
         End
         Begin VB.PictureBox cmdLast 
            Appearance      =   0  'Flat
            AutoRedraw      =   -1  'True
            BackColor       =   &H00CCCCCC&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   300
            Left            =   2160
            MouseIcon       =   "frmProd.frx":0C28
            MousePointer    =   99  'Custom
            ScaleHeight     =   300
            ScaleWidth      =   390
            TabIndex        =   21
            TabStop         =   0   'False
            Top             =   3240
            Width           =   390
         End
      End
      Begin VB.PictureBox ButtonContainer 
         BackColor       =   &H00808080&
         BorderStyle     =   0  'None
         Height          =   975
         Left            =   240
         ScaleHeight     =   975
         ScaleWidth      =   8775
         TabIndex        =   35
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
            MouseIcon       =   "frmProd.frx":0F32
            MousePointer    =   99  'Custom
            ScaleHeight     =   570
            ScaleWidth      =   1095
            TabIndex        =   16
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
            MouseIcon       =   "frmProd.frx":123C
            MousePointer    =   99  'Custom
            ScaleHeight     =   570
            ScaleWidth      =   1095
            TabIndex        =   15
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
            MouseIcon       =   "frmProd.frx":1546
            MousePointer    =   99  'Custom
            ScaleHeight     =   570
            ScaleWidth      =   1095
            TabIndex        =   14
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
            MouseIcon       =   "frmProd.frx":1850
            MousePointer    =   99  'Custom
            ScaleHeight     =   570
            ScaleWidth      =   1095
            TabIndex        =   13
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
            MouseIcon       =   "frmProd.frx":1B5A
            MousePointer    =   99  'Custom
            ScaleHeight     =   570
            ScaleWidth      =   1095
            TabIndex        =   12
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
            MouseIcon       =   "frmProd.frx":1E64
            MousePointer    =   99  'Custom
            ScaleHeight     =   570
            ScaleWidth      =   1095
            TabIndex        =   11
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
            MouseIcon       =   "frmProd.frx":216E
            MousePointer    =   99  'Custom
            ScaleHeight     =   570
            ScaleWidth      =   1095
            TabIndex        =   17
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
         TabIndex        =   24
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
            MouseIcon       =   "frmProd.frx":2478
            MousePointer    =   99  'Custom
            ScaleHeight     =   285
            ScaleWidth      =   315
            TabIndex        =   27
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
            MouseIcon       =   "frmProd.frx":2782
            MousePointer    =   99  'Custom
            ScaleHeight     =   285
            ScaleWidth      =   315
            TabIndex        =   26
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
            MouseIcon       =   "frmProd.frx":2A8C
            MousePointer    =   99  'Custom
            ScaleHeight     =   285
            ScaleWidth      =   315
            TabIndex        =   25
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
         Index           =   1
         Left            =   1980
         Locked          =   -1  'True
         TabIndex        =   0
         Top             =   840
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
         Index           =   2
         Left            =   1980
         MultiLine       =   -1  'True
         TabIndex        =   1
         Top             =   1200
         Width           =   3615
      End
      Begin VB.TextBox txtField 
         Alignment       =   1  'Right Justify
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
         Index           =   3
         Left            =   1980
         TabIndex        =   4
         Top             =   2280
         Width           =   1695
      End
      Begin VB.TextBox txtField 
         Alignment       =   1  'Right Justify
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
         Index           =   4
         Left            =   1980
         TabIndex        =   5
         Top             =   2640
         Width           =   1695
      End
      Begin VB.TextBox txtField 
         Alignment       =   1  'Right Justify
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
         Index           =   6
         Left            =   1980
         Locked          =   -1  'True
         TabIndex        =   8
         Top             =   3720
         Width           =   2535
      End
      Begin VB.TextBox txtField 
         Alignment       =   1  'Right Justify
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
         Index           =   5
         Left            =   1980
         TabIndex        =   6
         Top             =   3000
         Width           =   1695
      End
      Begin MSComCtl2.DTPicker DatePick 
         Height          =   315
         Left            =   1980
         TabIndex        =   10
         Top             =   4440
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
         Left            =   4320
         TabIndex        =   42
         Top             =   4680
         Width           =   765
      End
      Begin VB.Label Label 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H00808080&
         BackStyle       =   0  'Transparent
         Caption         =   "Category"
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
         Left            =   930
         TabIndex        =   41
         Top             =   1920
         Width           =   900
      End
      Begin VB.Label Label 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H00808080&
         BackStyle       =   0  'Transparent
         Caption         =   "Reorder Level"
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
         Left            =   465
         TabIndex        =   40
         Top             =   3360
         Width           =   1365
      End
      Begin VB.Label Label 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H00808080&
         BackStyle       =   0  'Transparent
         Caption         =   "Supplier"
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
         Left            =   1035
         TabIndex        =   39
         Top             =   4125
         Width           =   795
      End
      Begin VB.Label Label 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H00808080&
         BackStyle       =   0  'Transparent
         Caption         =   "Generic Name"
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
         Left            =   480
         TabIndex        =   38
         Top             =   1560
         Width           =   1350
      End
      Begin VB.Label Label 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H00808080&
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
         Height          =   225
         Index           =   3
         Left            =   525
         TabIndex        =   34
         Top             =   840
         Width           =   1305
      End
      Begin VB.Label Label 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H00808080&
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
         Height          =   225
         Index           =   4
         Left            =   945
         TabIndex        =   33
         Top             =   2280
         Width           =   885
      End
      Begin VB.Label Label 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H00808080&
         BackStyle       =   0  'Transparent
         Caption         =   "Quantity"
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
         Left            =   990
         TabIndex        =   32
         Top             =   3000
         Width           =   840
      End
      Begin VB.Label Label 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H00808080&
         BackStyle       =   0  'Transparent
         Caption         =   "Date Purchased"
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
         Left            =   300
         TabIndex        =   31
         Top             =   4440
         Width           =   1530
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
         Index           =   7
         Left            =   720
         TabIndex        =   30
         Top             =   1200
         Width           =   1110
      End
      Begin VB.Label Label 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H00808080&
         BackStyle       =   0  'Transparent
         Caption         =   "Selling Cost"
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
         Left            =   690
         TabIndex        =   29
         Top             =   2640
         Width           =   1140
      End
      Begin VB.Label Label 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H00808080&
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
         Height          =   225
         Index           =   10
         Left            =   855
         TabIndex        =   28
         Top             =   3765
         Width           =   975
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
Attribute VB_Name = "frmProduct"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'##############################################
'#          Coded by Walter A. Narvasa        #
'#       POS2000 - Point of Sales System      #
'#                                            #
'#           area :  frmProduct               #
'#    description :  Product Masterfile       #
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
    If Not Search_Exist(txtField(1), "PRODCODE", "INVOICE_DETAIL") Then
        Call MessageBox("frmProduct", "Are you sure you want to delete Product Code " + datprimary("PRODCODE") + " ?", 1)
        frmMessageBox2.Show
    Else
        Call MessageBox("frmProduct", "Cannot delete record, Record exist in Sales", 0)
        frmMessageBox2.Show
    End If
    Call MacButton("  Delete", frmProduct.cmdDel, 0, 0, 73, 50, frmLogin.Source, 0, 0, 1)
End Sub

Private Sub cmdDel_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call MacButton("  Delete", frmProduct.cmdDel, 0, 0, 73, 50, frmLogin.Source, 74, 0, 1)
End Sub

Private Sub cmdDel_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call MacButton("  Delete", frmProduct.cmdDel, 0, 0, 73, 50, frmLogin.Source, 0, 0, 1)
End Sub

Private Sub cmdEdit_Click()
'    On Error GoTo ErrorEdit
    EditMode = True
    Press_Buttons ("Edit")
    txtField(2).SetFocus
    txtField(1).TabStop = False
'ErrorEdit:
'    Call MessageBox("frmProduct", Err.Description, 0)
End Sub

Private Sub cmdEdit_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call MacButton("    Edit", frmProduct.cmdEdit, 0, 0, 73, 50, frmLogin.Source, 74, 0, 1)
End Sub

Private Sub cmdEdit_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call MacButton("    Edit", frmProduct.cmdEdit, 0, 0, 73, 50, frmLogin.Source, 0, 0, 1)
End Sub

Private Sub cmdExit_Click()
    EditMode = False
    Unload Me
End Sub

Private Sub cmdExit_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call MacButton("    Exit", frmProduct.cmdExit, 0, 0, 73, 50, frmLogin.Source, 74, 0, 1)
End Sub

Private Sub cmdExit_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call MacButton("    Exit", frmProduct.cmdExit, 0, 0, 73, 50, frmLogin.Source, 0, 0, 1)
End Sub

Private Sub cmdFind_Click()
    On Error Resume Next
    EditMode = True
    Call FindBox(" Find by Product Description ", _
                    "PROD_STOCKS", "PRODDES", 0, 1, "frmLogin")
    Call MacButton("   Find", frmProduct.cmdFind, 0, 0, 73, 50, frmLogin.Source, 0, 0, 1)
End Sub

Public Sub FindProductDes()
    On Error Resume Next
    strs = "select * from PROD_STOCKS where PRODDES = '" & frmFind.txtWord & "'"
    Set dummy = frmLogin.db.OpenRecordset(strs)
    If dummy.AbsolutePosition <> -1 Then
        datprimary.MoveFirst
        Do While Not datprimary.EOF
            If datprimary("PRODDES") = frmFind.txtWord Then
                Exit Do
            End If
                datprimary.MoveNext
        Loop
            Display_Fields
            Enable_Buttons
    Else
        Call MessageBox("frmProduct", "Product Description not found", 0)
    End If
End Sub

Private Sub cmdFind_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call MacButton("   Find", frmProduct.cmdFind, 0, 0, 73, 50, frmLogin.Source, 74, 0, 1)
End Sub

Private Sub cmdFind_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call MacButton("   Find", frmProduct.cmdFind, 0, 0, 73, 50, frmLogin.Source, 0, 0, 1)
End Sub

Private Sub cmdLast_Click()
    Press_Buttons ("Last")
End Sub

Private Sub cmdLast_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call MacButton("q", frmProduct.cmdLast, 0, 0, 100, 50, frmLogin.Source, 112, 39, 3)
End Sub

Private Sub cmdLast_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call MacButton("q", frmProduct.cmdLast, 0, 0, 100, 50, frmLogin.Source, 138, 39, 3)
End Sub

Private Sub cmdNew_Click()
    On Error Resume Next
    EditMode = False
    If datprimary.RecordCount = 0 Then
        txtField(1).Text = "1"
    Else
        datprimary.MoveLast
        txtField(1).Text = Val(datprimary("PRODCODE")) + 1
    End If
    Press_Buttons ("New")
    DatePick.Value = Get_Last_Date("PDATE", "PROD_STOCKS")
    txtField(1).SetFocus
End Sub

Private Sub cmdNew_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call MacButton("   New", frmProduct.cmdNew, 0, 0, 73, 50, frmLogin.Source, 74, 0, 1)
End Sub

Private Sub cmdNew_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call MacButton("   New", frmProduct.cmdNew, 0, 0, 73, 50, frmLogin.Source, 0, 0, 1)
End Sub

Private Sub cmdNext_Click()
    Press_Buttons ("Next")
End Sub

Private Sub cmdNext_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call MacButton("u", frmProduct.cmdNext, 0, 0, 100, 49, frmLogin.Source, 112, 39, 3)
End Sub

Private Sub cmdNext_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call MacButton("u", frmProduct.cmdNext, 0, 0, 100, 49, frmLogin.Source, 138, 39, 3)
End Sub

Private Sub cmdOk_Click()
    BoxContainer2.Visible = False
End Sub

Private Sub cmdPicture_Click()
    On Error GoTo CancelLoad
        cdlPicture.ShowOpen
    On Error GoTo BadLoad
        Photo = LoadPicture(cdlPicture.filename)
    On Error GoTo 0
    On Error GoTo 0
    Exit Sub
CancelLoad:
    If Err.Number <> cdlCancel Then
        Call MessageBox("frmProduct", Err.Description, 0)
        frmMessageBox.Show
    Else
        Exit Sub
    End If
BadLoad:
        Call MessageBox("frmProduct", Err.Description, 0)
        frmMessageBox.Show
    Exit Sub
End Sub

Private Sub cmdPicture_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call MacButton("     Change Picture", frmProduct.cmdPicture, 0, 0, 170, 30, frmLogin.Source, 182, 30, 2)
End Sub

Private Sub cmdPicture_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call MacButton("     Change Picture", frmProduct.cmdPicture, 0, 0, 170, 30, frmLogin.Source, 147, 0, 2)
End Sub

Private Sub cmdPrev_Click()
    Press_Buttons ("Prev")
End Sub

Private Sub cmdPrev_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call MacButton("t", frmProduct.cmdPrev, 0, 0, 100, 50, frmLogin.Source, 112, 39, 3)
End Sub

Private Sub cmdPrev_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call MacButton("t", frmProduct.cmdPrev, 0, 0, 100, 50, frmLogin.Source, 138, 39, 3)
End Sub

Private Sub cmdSave_Click()
    On Error Resume Next
    If EditMode = True Then
        Call MacButton("    Edit", frmProduct.cmdEdit, 0, 0, 73, 50, frmLogin.Source, 0, 0, 1)
        Press_Buttons ("Save")
    Else
        Call MacButton("   New", frmProduct.cmdNew, 0, 0, 73, 50, frmLogin.Source, 0, 0, 1)
        If Get_Product_Code Then
            Call MessageBox("frmProduct", "Product Code already exists.", 0)
            frmMessageBox.SetFocus
            txtField(1) = ""
            Press_Buttons ("Undo")
        ElseIf txtField(1) = "" Then
            Call MessageBox("frmProduct", "Product Code cannot be null", 0)
            frmMessageBox.SetFocus
            txtField(1) = ""
            Press_Buttons ("Undo")
        'ENABLE ONLY FOR SPECIFIC LENGTH OF 10 CHARACTERS NUMBERING
        'ElseIf Len(txtField(1)) > 1 And Len(txtField(1)) < 10 Then
        '    Call MessageBox("frmProduct", "Product number must be 10 in length", 0)
        '    frmMessageBox.SetFocus
        '    txtField(1) = ""
        '    Press_Buttons ("Undo")
        Else
            Press_Buttons ("Save")
        End If
    End If
    Call MacButton("   Save", frmProduct.cmdSave, 0, 0, 73, 50, frmLogin.Source, 0, 0, 1)
End Sub

Private Sub cmdSave_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call MacButton("   Save", frmProduct.cmdSave, 0, 0, 73, 50, frmLogin.Source, 74, 0, 1)
End Sub

Private Sub cmdSave_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call MacButton("   Save", frmProduct.cmdSave, 0, 0, 73, 50, frmLogin.Source, 0, 0, 1)
End Sub

Private Sub cmdTop_Click()
    Press_Buttons ("Top")
End Sub

Private Sub cmdTop_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call MacButton("p", frmProduct.cmdTop, 0, 0, 100, 50, frmLogin.Source, 112, 39, 3)
End Sub

Private Sub cmdTop_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call MacButton("p", frmProduct.cmdTop, 0, 0, 100, 50, frmLogin.Source, 138, 39, 3)
End Sub

Private Sub cmdUndo_Click()
    If EditMode = True Then
        Call MacButton("    Edit", frmProduct.cmdEdit, 0, 0, 73, 50, frmLogin.Source, 0, 0, 1)
    Else
        Call MacButton("   New", frmProduct.cmdNew, 0, 0, 73, 50, frmLogin.Source, 0, 0, 1)
    End If
    Press_Buttons ("Undo")
    Call MacButton("   Undo", frmProduct.cmdUndo, 0, 0, 73, 50, frmLogin.Source, 0, 0, 1)
End Sub

Private Sub cmdUndo_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call MacButton("   Undo", frmProduct.cmdUndo, 0, 0, 73, 50, frmLogin.Source, 74, 0, 1)
End Sub

Private Sub cmdUndo_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call MacButton("   Undo", frmProduct.cmdUndo, 0, 0, 73, 50, frmLogin.Source, 0, 0, 1)
End Sub

Private Sub Form_Activate()
    On Error Resume Next
    If EditMode = False Then
        strs = "select * from PROD_STOCKS order by PRODCODE"
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
    Call ColForm(PictureContainer, 217, 211, 213, 125)
    Call CreateMacOSTitleBar(titleBar, " Product Masterfile ")
    Call BitBlt(frmProduct.Closed.hDC, 0, 0, 73, 50, frmLogin.Source.hDC, 0, 107, SRCCOPY)
    frmProduct.Closed.Refresh
    Call BitBlt(frmProduct.Maximized.hDC, 0, 0, 73, 50, frmLogin.Source.hDC, 0, 72, SRCCOPY)
    frmProduct.Maximized.Refresh
    Call BitBlt(frmProduct.Minimized.hDC, 0, 0, 73, 50, frmLogin.Source.hDC, 0, 124, SRCCOPY)
    frmProduct.Minimized.Refresh
    KeyPreview = True
    Call MacButton("   New", frmProduct.cmdNew, 0, 0, 73, 50, frmLogin.Source, 0, 0, 1)
    Call MacButton("    Edit", frmProduct.cmdEdit, 0, 0, 73, 50, frmLogin.Source, 0, 0, 1)
    Call MacButton("   Save", frmProduct.cmdSave, 0, 0, 73, 50, frmLogin.Source, 0, 0, 1)
    Call MacButton("   Undo", frmProduct.cmdUndo, 0, 0, 73, 50, frmLogin.Source, 0, 0, 1)
    Call MacButton("  Delete", frmProduct.cmdDel, 0, 0, 73, 50, frmLogin.Source, 0, 0, 1)
    Call MacButton("   Find", frmProduct.cmdFind, 0, 0, 73, 50, frmLogin.Source, 0, 0, 1)
    Call MacButton("    Exit", frmProduct.cmdExit, 0, 0, 73, 50, frmLogin.Source, 0, 0, 1)
    Call MacButton("     Change Picture", frmProduct.cmdPicture, 0, 0, 170, 30, frmLogin.Source, 147, 0, 2)
    Call MacButton("p", frmProduct.cmdTop, 0, 0, 100, 50, frmLogin.Source, 138, 39, 3)
    Call MacButton("u", frmProduct.cmdNext, 0, 0, 100, 50, frmLogin.Source, 138, 39, 3)
    Call MacButton("t", frmProduct.cmdPrev, 0, 0, 100, 50, frmLogin.Source, 138, 39, 3)
    Call MacButton("q", frmProduct.cmdLast, 0, 0, 100, 50, frmLogin.Source, 138, 39, 3)
    txtCombo(0).Clear
    strs = "select CATCODE from CATEGORY order by CATCODE"
    Set dummy = frmLogin.db.OpenRecordset(strs)
    If dummy.AbsolutePosition <> -1 Then
      Do While Not dummy.EOF
        If Not IsNull(dummy(0)) Then
            txtCombo(0).AddItem (dummy(0))
        End If
        dummy.MoveNext
       Loop
    End If
    txtCombo(1).Clear
    strs = "select SUPCODE from SUPPLIER order by SUPCODE"
    Set dummy = frmLogin.db.OpenRecordset(strs)
    If dummy.AbsolutePosition <> -1 Then
        Do While Not dummy.EOF
            If Not IsNull(dummy(0)) Then
                txtCombo(1).AddItem (dummy(0))
            End If
            dummy.MoveNext
       Loop
    End If
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
                    Call MacButton("   New", frmProduct.cmdNew, 0, 0, 73, 50, frmLogin.Source, 0, 0, 1)
                End If
        Case vbKeyE:
                If AltDown Then
                    Call MacButton("    Edit", frmProduct.cmdEdit, 0, 0, 73, 50, frmLogin.Source, 0, 0, 1)
                End If
        Case vbKeyS:
                If AltDown Then
                    Call MacButton("   Save", frmProduct.cmdSave, 0, 0, 73, 50, frmLogin.Source, 0, 0, 1)
                End If
        Case vbKeyU:
                If AltDown Then
                    Call MacButton("   Undo", frmProduct.cmdUndo, 0, 0, 73, 50, frmLogin.Source, 0, 0, 1)
                End If
        Case vbKeyD:
                If AltDown Then
                    Call MacButton("  Delete", frmProduct.cmdDel, 0, 0, 73, 50, frmLogin.Source, 0, 0, 1)
                End If
        Case vbKeyF:
                If AltDown Then
                    Call MacButton("   Find", frmProduct.cmdFind, 0, 0, 73, 50, frmLogin.Source, 0, 0, 1)
                End If
        Case vbKeyX:
                If AltDown Then
                    Call MacButton("    Exit", frmProduct.cmdExit, 0, 0, 73, 50, frmLogin.Source, 0, 0, 1)
                End If
        Case vbKeyP:
                If AltDown Then
                    Call MacButton("     Change Picture", frmProduct.cmdPicture, 0, 0, 170, 30, frmLogin.Source, 147, 0, 2)
                End If
        Case vbKeyLeft:
                If AltDown Then
                    Call MacButton("t", frmProduct.cmdPrev, 0, 0, 100, 50, frmLogin.Source, 138, 39, 3)
                End If
        Case vbKeyRight:
                If AltDown Then
                    Call MacButton("u", frmProduct.cmdNext, 0, 0, 100, 49, frmLogin.Source, 138, 39, 3)
                End If
        Case vbKeyUp:
                If AltDown Then
                    Call MacButton("p", frmProduct.cmdTop, 0, 0, 100, 50, frmLogin.Source, 138, 39, 3)
                End If
        Case vbKeyDown:
                If AltDown Then
                    Call MacButton("q", frmProduct.cmdLast, 0, 0, 100, 50, frmLogin.Source, 138, 39, 3)
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
                                      "Note: The following are Product Masterfile Key Shortcuts." & vbCrLf & _
                                      Space(1) & vbCrLf & _
                                      "ALT-N=New, ALT-E=Edit, ALT-S=Save, ALT-U=Undo, ALT-D=Delete" & vbCrLf & _
                                      "ALT-F=Find, ALT-X=Exit, ALT-P=Change Picture" & vbCrLf & _
                                      Space(1) & vbCrLf & _
                                      "Left Arrow=Previous Records, Right Arrow=Next Records" & vbCrLf & _
                                      "Top Arrow=Top Record, Down Arrow=Last Record"
        Case vbKeyN:
                If AltDown Then
                    Call MacButton("   New", frmProduct.cmdNew, 0, 0, 73, 50, frmLogin.Source, 74, 0, 1)
                    cmdNew_Click
                End If
        Case vbKeyE:
                If AltDown Then
                    Call MacButton("    Edit", frmProduct.cmdEdit, 0, 0, 73, 50, frmLogin.Source, 74, 0, 1)
                    cmdEdit_Click
                End If
        Case vbKeyS:
                If AltDown Then
                    Call MacButton("   Save", frmProduct.cmdSave, 0, 0, 73, 50, frmLogin.Source, 74, 0, 1)
                    cmdSave_Click
                End If
        Case vbKeyU:
                If AltDown Then
                    Call MacButton("   Undo", frmProduct.cmdUndo, 0, 0, 73, 50, frmLogin.Source, 74, 0, 1)
                    cmdUndo_Click
                End If
        Case vbKeyD:
                If AltDown Then
                    Call MacButton("  Delete", frmProduct.cmdDel, 0, 0, 73, 50, frmLogin.Source, 74, 0, 1)
                    cmdDel_Click
                End If
        Case vbKeyF:
                If AltDown Then
                    Call MacButton("   Find", frmProduct.cmdFind, 0, 0, 73, 50, frmLogin.Source, 74, 0, 1)
                    cmdFind_Click
                End If
        Case vbKeyX:
                If AltDown Then
                    Call MacButton("    Exit", frmProduct.cmdExit, 0, 0, 73, 50, frmLogin.Source, 74, 0, 1)
                    cmdExit_Click
                End If
        Case vbKeyP:
                If AltDown Then
                    Call MacButton("     Change Picture", frmProduct.cmdPicture, 0, 0, 170, 30, frmLogin.Source, 182, 30, 2)
                    cmdPicture_Click
                End If
        Case vbKeyLeft:
                If AltDown Then
                    Call MacButton("t", frmProduct.cmdPrev, 0, 0, 100, 50, frmLogin.Source, 112, 39, 3)
                    cmdPrev_Click
                End If
        Case vbKeyRight:
                If AltDown Then
                    Call MacButton("u", frmProduct.cmdNext, 0, 0, 100, 49, frmLogin.Source, 112, 39, 3)
                    cmdNext_Click
                End If
        Case vbKeyUp:
                If AltDown Then
                    Call MacButton("p", frmProduct.cmdTop, 0, 0, 100, 50, frmLogin.Source, 112, 39, 3)
                    cmdTop_Click
                End If
        Case vbKeyDown:
                If AltDown Then
                    Call MacButton("q", frmProduct.cmdLast, 0, 0, 100, 50, frmLogin.Source, 112, 39, 3)
                    cmdLast_Click
                End If
    End Select
End Sub

Private Sub titleBar_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call DragForm(Me)
End Sub

Private Sub Closed_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call BitBlt(frmProduct.Closed.hDC, 0, 0, 73, 50, frmLogin.Source.hDC, 18, 107, SRCCOPY)
    frmProduct.Closed.Refresh
End Sub

Private Sub Closed_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call BitBlt(frmProduct.Closed.hDC, 0, 0, 73, 50, frmLogin.Source.hDC, 0, 107, SRCCOPY)
    frmProduct.Closed.Refresh
End Sub

Private Sub Maximized_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call BitBlt(frmProduct.Maximized.hDC, 0, 0, 73, 50, frmLogin.Source.hDC, 18, 72, SRCCOPY)
    frmProduct.Maximized.Refresh
End Sub

Private Sub Maximized_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call BitBlt(frmProduct.Maximized.hDC, 0, 0, 73, 50, frmLogin.Source.hDC, 0, 72, SRCCOPY)
    frmProduct.Maximized.Refresh
End Sub

Private Sub Minimized_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call BitBlt(frmProduct.Minimized.hDC, 0, 0, 73, 50, frmLogin.Source.hDC, 18, 124, SRCCOPY)
    frmProduct.Minimized.Refresh
End Sub

Private Sub Minimized_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call BitBlt(frmProduct.Minimized.hDC, 0, 0, 73, 50, frmLogin.Source.hDC, 0, 124, SRCCOPY)
    frmProduct.Minimized.Refresh
End Sub

Private Sub Display_Fields()
    On Error Resume Next
    If datprimary.AbsolutePosition <> -1 Then
        txtField(1) = IIf(IsNull(datprimary("PRODCODE")), "", datprimary("PRODCODE"))
        txtField(7) = IIf(IsNull(datprimary("GENERICNAME")), "", datprimary("GENERICNAME"))
        txtField(2) = IIf(IsNull(datprimary("PRODDES")), "", datprimary("PRODDES"))
        txtCombo(0) = IIf(IsNull(datprimary("CATCODE")), "", datprimary("CATCODE"))
        txtField(3) = IIf(IsNull(datprimary("UNIT_COST")), "0", datprimary("UNIT_COST"))
        txtField(4) = IIf(IsNull(datprimary("SELLING_PRICE")), "0", datprimary("SELLING_PRICE"))
        txtField(5) = IIf(IsNull(datprimary("QUAN")), "0", datprimary("QUAN"))
        txtField(0) = IIf(IsNull(datprimary("REORDER")), "0", datprimary("REORDER"))
        txtCombo(1) = IIf(IsNull(datprimary("SUPCODE")), "", datprimary("SUPCODE"))
        DatePick.Value = IIf(IsDate(datprimary("PDATE")), datprimary("PDATE"), Date)
        Call ResetPicture(Photo)
        Call DisplayPicture(datprimary("PICTURE"))
    Else
        Clear_Fields
    End If
        txtField(3) = Convert_Numeric(txtField(3), True)
        txtField(4) = Convert_Numeric(txtField(4), True)
        txtField(5) = Convert_Numeric(txtField(5), False)
        txtField(0) = Convert_Numeric(txtField(0), False)
        txtField(6) = Convert_Numeric(CDbl(txtField(3)) * CDbl(txtField(5)), True)
End Sub

Private Sub Clear_Fields()
    ' ENABLE ONLY TO BLANK FIELD txtField(1)
    'txtField(1) = ""
    txtField(7) = ""
    txtField(2) = ""
    txtField(3) = "0"
    txtField(4) = "0"
    txtField(5) = "0"
    txtField(0) = "0"
    txtField(6) = "0"
End Sub


Private Sub Update_Fields(isNew As Boolean)
    On Error Resume Next
    Dim SaveToPicture
    If isNew Then
        datprimary.AddNew
    Else
        datprimary.Edit
    End If
        datprimary("PRODCODE") = txtField(1)
        datprimary("GENERICNAME") = txtField(7)
        datprimary("PRODDES") = txtField(2)
        datprimary("CATCODE") = txtCombo(0)
        datprimary("UNIT_COST") = IIf(txtField(3) = "", 0, txtField(3))
        datprimary("SELLING_PRICE") = IIf(txtField(4) = "", 0, txtField(4))
        datprimary("QUAN") = IIf(txtField(5) = "", 0, txtField(5))
        datprimary("REORDER") = IIf(txtField(0) = "", 0, txtField(0))
        datprimary("SUPCODE") = txtCombo(1)
        datprimary("PDATE") = DatePick.Value
        SaveToPicture = CopyFileToField(cdlPicture.filename, datprimary("PICTURE"))
        datprimary.Update
    If isNew Then datprimary.MoveLast
End Sub

Private Sub Enable_Fields(isLock As Boolean)
    On Error Resume Next
    For i = 0 To 7
        txtField(i).Enabled = Not isLock
    Next i
        txtCombo(0).Enabled = Not isLock
        txtCombo(1).Enabled = Not isLock
        DatePick.Enabled = Not isLock
    If p_isadding Then
        txtField(1).Locked = isLock
        txtField(1).TabStop = True
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
                            
Private Sub txtField_KeyPress(Index As Integer, KeyAscii As Integer)
    On Error Resume Next
    If Index = 0 Or Index = 3 Or Index = 4 Or Index = 5 Then
        Select Case KeyAscii
                Case Asc("0") To Asc("9")
                Case Asc(".")
                    KeyAscii = IIf(Index = 5 Or Index = 1, 0, KeyAscii)
                Case KeyAscii = vbKeyBack
                Case Else
                    KeyAscii = 0
        End Select
    End If
End Sub

Private Sub txtField_LostFocus(Index As Integer)
    On Error Resume Next
    If Index = 3 Then txtField(3) = Convert_Numeric(txtField(3), True)
    If Index = 4 Then txtField(4) = Convert_Numeric(txtField(4), True)
    If Index = 0 Or Index = 5 Or Index = 3 Then
        txtField(5) = Convert_Numeric(txtField(5), False)
        txtField(0) = Convert_Numeric(txtField(0), False)
        txtField(6) = Convert_Numeric(CStr(CDbl(txtField(3)) * CDbl(txtField(5))), True)
    End If
End Sub

Private Sub Object_Tab_Trigger(isTab As Boolean)
    On Error Resume Next
    txtField(2).TabStop = isTab
    txtField(7).TabStop = isTab
    txtCombo(0).TabStop = isTab
    txtField(3).TabStop = isTab
    txtField(4).TabStop = isTab
    txtField(5).TabStop = isTab
    txtField(0).TabStop = isTab
    txtCombo(1).TabStop = isTab
    DatePick.TabStop = isTab
End Sub

Function Get_Product_Code() As Boolean
    On Error Resume Next
    strs = "select * from PROD_STOCKS where PRODCODE = '" & txtField(1) & "'"
    Set dummy = frmLogin.db.OpenRecordset(strs)
    If dummy.AbsolutePosition <> -1 Then
        Get_Product_Code = True
    Else
        Get_Product_Code = False
    End If
End Function

Function DisplayPicture(xField As Field)
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
        Photo = LoadPicture(filename)
    End If
End Function

Function ResetPicture(xPicture As PictureBox)
    On Error Resume Next
    filename = ""
    xPicture.Picture = LoadPicture("")
    xPicture.Refresh
End Function
