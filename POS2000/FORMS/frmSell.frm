VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Begin VB.Form frmSelling 
   BackColor       =   &H00CCCCCC&
   BorderStyle     =   0  'None
   Caption         =   "frmSelling"
   ClientHeight    =   9000
   ClientLeft      =   3135
   ClientTop       =   1860
   ClientWidth     =   12000
   FillColor       =   &H00808080&
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
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
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   40
      Width           =   11890
      Begin VB.PictureBox DetailBoxContainer 
         BackColor       =   &H00808080&
         BorderStyle     =   0  'None
         Height          =   2175
         Left            =   480
         ScaleHeight     =   2175
         ScaleWidth      =   11010
         TabIndex        =   41
         TabStop         =   0   'False
         Top             =   6360
         Visible         =   0   'False
         Width           =   11010
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
            Left            =   600
            TabIndex        =   43
            TabStop         =   0   'False
            Top             =   960
            Visible         =   0   'False
            Width           =   5970
         End
         Begin VB.PictureBox cmdReturn 
            Appearance      =   0  'Flat
            AutoRedraw      =   -1  'True
            BackColor       =   &H00CCCCCC&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   570
            Left            =   3000
            MouseIcon       =   "frmSell.frx":0000
            MousePointer    =   99  'Custom
            ScaleHeight     =   570
            ScaleWidth      =   1095
            TabIndex        =   56
            TabStop         =   0   'False
            Top             =   1440
            Width           =   1100
         End
         Begin VB.PictureBox cmdTag 
            Appearance      =   0  'Flat
            AutoRedraw      =   -1  'True
            BackColor       =   &H00CCCCCC&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   570
            Left            =   1800
            MouseIcon       =   "frmSell.frx":030A
            MousePointer    =   99  'Custom
            ScaleHeight     =   570
            ScaleWidth      =   1095
            TabIndex        =   55
            TabStop         =   0   'False
            Top             =   1440
            Width           =   1100
         End
         Begin VB.PictureBox cmdStop 
            Appearance      =   0  'Flat
            AutoRedraw      =   -1  'True
            BackColor       =   &H00CCCCCC&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   570
            Left            =   4200
            MouseIcon       =   "frmSell.frx":0614
            MousePointer    =   99  'Custom
            ScaleHeight     =   570
            ScaleWidth      =   1095
            TabIndex        =   54
            TabStop         =   0   'False
            Top             =   1440
            Width           =   1100
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
            Left            =   2400
            MultiLine       =   -1  'True
            TabIndex        =   7
            Top             =   960
            Width           =   1695
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
            Left            =   8595
            Locked          =   -1  'True
            MultiLine       =   -1  'True
            TabIndex        =   52
            TabStop         =   0   'False
            Top             =   1680
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
            Left            =   8595
            Locked          =   -1  'True
            MultiLine       =   -1  'True
            TabIndex        =   51
            TabStop         =   0   'False
            Top             =   1320
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
            Left            =   8595
            Locked          =   -1  'True
            MultiLine       =   -1  'True
            TabIndex        =   50
            TabStop         =   0   'False
            Top             =   960
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
            Left            =   8595
            Locked          =   -1  'True
            MultiLine       =   -1  'True
            TabIndex        =   45
            TabStop         =   0   'False
            Top             =   600
            Width           =   1695
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
            Left            =   2400
            TabIndex        =   6
            Text            =   "Type first few letter of the word..."
            Top             =   600
            Width           =   4170
         End
         Begin VB.PictureBox titleBar2 
            BorderStyle     =   0  'None
            Height          =   375
            Left            =   120
            ScaleHeight     =   375
            ScaleWidth      =   10815
            TabIndex        =   42
            TabStop         =   0   'False
            Top             =   90
            Width           =   10815
         End
         Begin VB.Label Lbls 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Quantity Sold"
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
            Left            =   960
            TabIndex        =   53
            Top             =   1005
            Width           =   1305
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
            Left            =   7485
            TabIndex        =   49
            Top             =   1725
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
            Left            =   6960
            TabIndex        =   48
            Top             =   1365
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
            Left            =   7575
            TabIndex        =   47
            Top             =   1005
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
            Left            =   7155
            TabIndex        =   46
            Top             =   645
            Width           =   1305
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
            Left            =   615
            TabIndex        =   44
            Top             =   600
            Width           =   1650
         End
      End
      Begin VB.PictureBox DetailsContainer 
         BackColor       =   &H00808080&
         BorderStyle     =   0  'None
         Height          =   4185
         Left            =   240
         ScaleHeight     =   4185
         ScaleWidth      =   11445
         TabIndex        =   25
         TabStop         =   0   'False
         Top             =   3360
         Width           =   11450
         Begin VB.PictureBox cmdTotal 
            Appearance      =   0  'Flat
            AutoRedraw      =   -1  'True
            BackColor       =   &H00CCCCCC&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   450
            Left            =   240
            MouseIcon       =   "frmSell.frx":091E
            MousePointer    =   99  'Custom
            ScaleHeight     =   450
            ScaleWidth      =   2355
            TabIndex        =   39
            TabStop         =   0   'False
            Top             =   3560
            Width           =   2355
         End
         Begin VB.TextBox txtChange 
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
            Height          =   285
            Left            =   9560
            TabIndex        =   5
            Top             =   3720
            Width           =   1695
         End
         Begin VB.TextBox txtPaid 
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
            Height          =   285
            Left            =   9560
            TabIndex        =   4
            Top             =   3360
            Width           =   1695
         End
         Begin VB.TextBox recTxt 
            Height          =   285
            Left            =   6720
            Locked          =   -1  'True
            TabIndex        =   31
            TabStop         =   0   'False
            Top             =   3480
            Visible         =   0   'False
            Width           =   150
         End
         Begin VB.TextBox txtTotal 
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
            Height          =   285
            Left            =   9560
            TabIndex        =   3
            Top             =   3000
            Width           =   1695
         End
         Begin VB.PictureBox cmdTop 
            Appearance      =   0  'Flat
            AutoRedraw      =   -1  'True
            BackColor       =   &H00CCCCCC&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   300
            Left            =   2760
            MouseIcon       =   "frmSell.frx":0C28
            MousePointer    =   99  'Custom
            ScaleHeight     =   300
            ScaleWidth      =   390
            TabIndex        =   30
            TabStop         =   0   'False
            Top             =   3075
            Width           =   390
         End
         Begin VB.PictureBox cmdPrev 
            Appearance      =   0  'Flat
            AutoRedraw      =   -1  'True
            BackColor       =   &H00CCCCCC&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   300
            Left            =   3240
            MouseIcon       =   "frmSell.frx":0F32
            MousePointer    =   99  'Custom
            ScaleHeight     =   300
            ScaleWidth      =   390
            TabIndex        =   29
            TabStop         =   0   'False
            Top             =   3075
            Width           =   390
         End
         Begin VB.PictureBox cmdNext 
            Appearance      =   0  'Flat
            AutoRedraw      =   -1  'True
            BackColor       =   &H00CCCCCC&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   300
            Left            =   3720
            MouseIcon       =   "frmSell.frx":123C
            MousePointer    =   99  'Custom
            ScaleHeight     =   300
            ScaleWidth      =   390
            TabIndex        =   28
            TabStop         =   0   'False
            Top             =   3075
            Width           =   390
         End
         Begin VB.PictureBox cmdLast 
            Appearance      =   0  'Flat
            AutoRedraw      =   -1  'True
            BackColor       =   &H00CCCCCC&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   300
            Left            =   4200
            MouseIcon       =   "frmSell.frx":1546
            MousePointer    =   99  'Custom
            ScaleHeight     =   300
            ScaleWidth      =   390
            TabIndex        =   27
            TabStop         =   0   'False
            Top             =   3075
            Width           =   390
         End
         Begin VB.PictureBox cmdDetail 
            Appearance      =   0  'Flat
            AutoRedraw      =   -1  'True
            BackColor       =   &H00CCCCCC&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   450
            Left            =   240
            MouseIcon       =   "frmSell.frx":1850
            MousePointer    =   99  'Custom
            ScaleHeight     =   450
            ScaleWidth      =   2355
            TabIndex        =   26
            TabStop         =   0   'False
            Top             =   3000
            Width           =   2355
         End
         Begin VB.TextBox txtField 
            DataField       =   "Invoice_No"
            DataSource      =   "datPrimaryRS"
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
            Left            =   1500
            Locked          =   -1  'True
            MaxLength       =   5
            TabIndex        =   0
            Top             =   217
            Width           =   1695
         End
         Begin VB.TextBox txtField 
            DataField       =   "ORNumber"
            DataSource      =   "datPrimaryRS"
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
            Left            =   3880
            Locked          =   -1  'True
            TabIndex        =   1
            Top             =   217
            Width           =   1455
         End
         Begin MSDBGrid.DBGrid grdDataGrid 
            Bindings        =   "frmSell.frx":1B5A
            Height          =   2295
            Left            =   240
            OleObjectBlob   =   "frmSell.frx":1B77
            TabIndex        =   32
            TabStop         =   0   'False
            Top             =   600
            Width           =   11010
         End
         Begin MSComCtl2.DTPicker DTPick 
            DataField       =   "DateSold"
            DataSource      =   "datPrimaryRS"
            Height          =   300
            Left            =   6500
            TabIndex        =   2
            Top             =   210
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
            Format          =   24707075
            CurrentDate     =   36527
         End
         Begin VB.Label Lbls 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Change"
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
            Index           =   5
            Left            =   8670
            TabIndex        =   38
            Top             =   3720
            Width           =   720
         End
         Begin VB.Label Lbls 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Amount Paid"
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
            Index           =   4
            Left            =   8160
            TabIndex        =   37
            Top             =   3360
            Width           =   1230
         End
         Begin VB.Label Lbls 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Total Amount Due"
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
            Index           =   3
            Left            =   7680
            TabIndex        =   36
            Top             =   3030
            Width           =   1710
         End
         Begin VB.Label Lbls 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Date Sold"
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
            Index           =   2
            Left            =   5460
            TabIndex        =   35
            Top             =   240
            Width           =   915
         End
         Begin VB.Label Lbls 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Invoice No."
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
            Left            =   240
            TabIndex        =   34
            Top             =   240
            Width           =   1110
         End
         Begin VB.Label Lbls 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "OR #"
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
            Left            =   3315
            TabIndex        =   33
            Top             =   240
            Width           =   450
         End
      End
      Begin VB.PictureBox CounterContainer 
         BackColor       =   &H00808080&
         BorderStyle     =   0  'None
         Height          =   2630
         Left            =   240
         ScaleHeight     =   2625
         ScaleWidth      =   11445
         TabIndex        =   23
         TabStop         =   0   'False
         Top             =   600
         Width           =   11450
         Begin VB.Data datSecondaryRS 
            Caption         =   "datSecondaryRS"
            Connect         =   "Access"
            DatabaseName    =   "C:\POS2000\DATABASE\POSDATA.mdb"
            DefaultCursorType=   0  'DefaultCursor
            DefaultType     =   2  'UseODBC
            Exclusive       =   0   'False
            Height          =   345
            Left            =   4560
            Options         =   0
            ReadOnly        =   0   'False
            RecordsetType   =   1  'Dynaset
            RecordSource    =   ""
            Top             =   960
            Visible         =   0   'False
            Width           =   2340
         End
         Begin VB.Data datPrimaryRS 
            Caption         =   "datPrimaryRS"
            Connect         =   "Access"
            DatabaseName    =   "\ARTFAB\Artdb.mdb"
            DefaultCursorType=   0  'DefaultCursor
            DefaultType     =   2  'UseODBC
            Exclusive       =   0   'False
            Height          =   345
            Left            =   2040
            Options         =   0
            ReadOnly        =   0   'False
            RecordsetType   =   1  'Dynaset
            RecordSource    =   "Invoice"
            Top             =   960
            Visible         =   0   'False
            Width           =   2340
         End
         Begin VB.Label TerminalDescription 
            BackColor       =   &H00000000&
            BeginProperty Font 
               Name            =   "Chicago"
               Size            =   15.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H0000FF00&
            Height          =   405
            Left            =   240
            TabIndex        =   40
            Top             =   240
            Width           =   11010
         End
         Begin VB.Label TerminalCounter 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00000000&
            Caption         =   "0.00"
            BeginProperty Font 
               Name            =   "Chicago"
               Size            =   72
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H0000FF00&
            Height          =   1695
            Left            =   240
            TabIndex        =   24
            Top             =   720
            Width           =   11010
         End
      End
      Begin VB.PictureBox ButtonContainer 
         BackColor       =   &H00808080&
         BorderStyle     =   0  'None
         Height          =   975
         Left            =   240
         ScaleHeight     =   975
         ScaleWidth      =   11490
         TabIndex        =   13
         TabStop         =   0   'False
         Top             =   7680
         Width           =   11490
         Begin VB.PictureBox cmdHelp 
            Appearance      =   0  'Flat
            AutoRedraw      =   -1  'True
            BackColor       =   &H00CCCCCC&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   570
            Left            =   8760
            MouseIcon       =   "frmSell.frx":254E
            MousePointer    =   99  'Custom
            ScaleHeight     =   570
            ScaleWidth      =   1095
            TabIndex        =   22
            TabStop         =   0   'False
            Top             =   240
            Width           =   1100
         End
         Begin VB.PictureBox cmdPrint 
            Appearance      =   0  'Flat
            AutoRedraw      =   -1  'True
            BackColor       =   &H00CCCCCC&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   570
            Left            =   6360
            MouseIcon       =   "frmSell.frx":2858
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
            Left            =   1560
            MouseIcon       =   "frmSell.frx":2B62
            MousePointer    =   99  'Custom
            ScaleHeight     =   570
            ScaleWidth      =   1095
            TabIndex        =   20
            TabStop         =   0   'False
            Top             =   240
            Width           =   1100
         End
         Begin VB.PictureBox cmdAdd 
            Appearance      =   0  'Flat
            AutoRedraw      =   -1  'True
            BackColor       =   &H00CCCCCC&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   570
            Left            =   360
            MouseIcon       =   "frmSell.frx":2E6C
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
            Left            =   3960
            MouseIcon       =   "frmSell.frx":3176
            MousePointer    =   99  'Custom
            ScaleHeight     =   570
            ScaleWidth      =   1095
            TabIndex        =   18
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
            Left            =   2760
            MouseIcon       =   "frmSell.frx":3480
            MousePointer    =   99  'Custom
            ScaleHeight     =   570
            ScaleWidth      =   1095
            TabIndex        =   17
            TabStop         =   0   'False
            Top             =   240
            Width           =   1100
         End
         Begin VB.PictureBox cmdFind 
            Appearance      =   0  'Flat
            AutoRedraw      =   -1  'True
            BackColor       =   &H00CCCCCC&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   570
            Left            =   7560
            MouseIcon       =   "frmSell.frx":378A
            MousePointer    =   99  'Custom
            ScaleHeight     =   570
            ScaleWidth      =   1095
            TabIndex        =   16
            TabStop         =   0   'False
            Top             =   240
            Width           =   1100
         End
         Begin VB.PictureBox cmdDelete 
            Appearance      =   0  'Flat
            AutoRedraw      =   -1  'True
            BackColor       =   &H00CCCCCC&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   570
            Left            =   5160
            MouseIcon       =   "frmSell.frx":3A94
            MousePointer    =   99  'Custom
            ScaleHeight     =   570
            ScaleWidth      =   1095
            TabIndex        =   15
            TabStop         =   0   'False
            Top             =   240
            Width           =   1100
         End
         Begin VB.PictureBox cmdClose 
            Appearance      =   0  'Flat
            AutoRedraw      =   -1  'True
            BackColor       =   &H00CCCCCC&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   570
            Left            =   9960
            MouseIcon       =   "frmSell.frx":3D9E
            MousePointer    =   99  'Custom
            ScaleHeight     =   570
            ScaleWidth      =   1095
            TabIndex        =   14
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
         ScaleWidth      =   11655
         TabIndex        =   9
         TabStop         =   0   'False
         Top             =   90
         Width           =   11655
         Begin VB.PictureBox Maximized 
            Appearance      =   0  'Flat
            AutoRedraw      =   -1  'True
            BackColor       =   &H00CCCCCC&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   280
            Left            =   11065
            MouseIcon       =   "frmSell.frx":40A8
            MousePointer    =   99  'Custom
            ScaleHeight     =   285
            ScaleWidth      =   315
            TabIndex        =   12
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
            Left            =   11335
            MouseIcon       =   "frmSell.frx":43B2
            MousePointer    =   99  'Custom
            ScaleHeight     =   285
            ScaleWidth      =   315
            TabIndex        =   11
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
            MouseIcon       =   "frmSell.frx":46BC
            MousePointer    =   99  'Custom
            ScaleHeight     =   285
            ScaleWidth      =   315
            TabIndex        =   10
            TabStop         =   0   'False
            Top             =   60
            Width           =   315
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
Attribute VB_Name = "frmSelling"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'##############################################
'#          Coded by Walter A. Narvasa        #
'#       POS2000 - Point of Sales System      #
'#                                            #
'#           area :  frmSelling               #
'#    description :  Selling Transaction      #
'#         e-mail :  walter@wancom.8k.com     #
'#            url :  http://wancom.8k.com     #
'#                                            #
'##############################################

Dim p_add, p_edit, p_save, p_undo, p_top, p_prev, p_next, p_last, p_del
Dim p_isadding, p_isediting, p_isnavigate
Dim dummy As DAO.Recordset
Dim datprimary2 As DAO.Recordset
Dim invoice As DAO.Recordset
        
Private Sub cmdAdd_Click()
'    On Error GoTo ErrorAdd
    Dim xCount
    EditMode = False
    If datPrimaryRS.Recordset.RecordCount = 0 Then
        xCount = "1"
    Else
        datPrimaryRS.Recordset.MoveLast
        xCount = Val(datPrimaryRS.Recordset.Fields("INVOICE_NO")) + 1
    End If
    Press_Button ("ADD")
    datPrimaryRS.Recordset.AddNew
    txtField(0).Text = xCount
    DTPick.Value = Get_Last_Date("DATESOLD", "INVOICE")
    txtField(1) = Get_OrNo
    txtField(0).SetFocus
    txtTotal = ""
'ErrorAdd:
'    Call MessageBox("frmSelling", Err.Description, 0)
End Sub

Private Sub cmdAdd_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call MacButton("   New", frmSelling.cmdAdd, 0, 0, 73, 50, frmLogin.Source, 74, 0, 1)
End Sub

Private Sub cmdAdd_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call MacButton("   New", frmSelling.cmdAdd, 0, 0, 73, 50, frmLogin.Source, 0, 0, 1)
End Sub

Private Sub cmdDelete_Click()
    'On Error GoTo ErrorDelete
    EditMode = True
    Call MacButton("  Delete", frmSelling.cmdDelete, 0, 0, 73, 50, frmLogin.Source, 0, 0, 1)
    Call MessageBox("frmSelling", "Are you sure you want to delete Invoice #" + txtField(0) + " ?", 1)
    frmMessageBox2.Show
'ErrorDelete:
'    Call MessageBox("frmSelling", Err.Description, 0)
End Sub

Private Sub cmdDelete_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call MacButton("  Delete", frmSelling.cmdDelete, 0, 0, 73, 50, frmLogin.Source, 74, 0, 1)
End Sub

Private Sub cmdDelete_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call MacButton("  Delete", frmSelling.cmdDelete, 0, 0, 73, 50, frmLogin.Source, 0, 0, 1)
End Sub

Private Sub cmdClose_Click()
    EditMode = False
    Screen.MousePointer = vbDefault
    Unload Me
End Sub

Private Sub cmdClose_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call MacButton("    Exit", frmSelling.cmdClose, 0, 0, 73, 50, frmLogin.Source, 74, 0, 1)
End Sub

Private Sub cmdClose_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call MacButton("    Exit", frmSelling.cmdClose, 0, 0, 73, 50, frmLogin.Source, 0, 0, 1)
End Sub

Private Sub cmdDetail_Click()
'    On Error GoTo ErrorDetail
    DetailBoxContainer.Visible = True
    txtField(0).TabStop = False
    txtField(1).TabStop = False
    DTPick.TabStop = False
    txtTotal.TabStop = False
    txtPaid.TabStop = False
    txtChange.TabStop = False
    txtWord.TabStop = True
    txtQuantitySold.TabStop = True
    txtWord.SetFocus
    txtWord = ""
'ErrorDetail:
    Call MacButton("   Tag/Return Items", frmSelling.cmdDetail, 0, 0, 170, 30, frmLogin.Source, 147, 0, 2)
End Sub

Private Sub cmdDetail_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call MacButton("   Tag/Return Items", frmSelling.cmdDetail, 0, 0, 170, 30, frmLogin.Source, 182, 30, 2)
End Sub

Private Sub cmdDetail_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call MacButton("   Tag/Return Items", frmSelling.cmdDetail, 0, 0, 170, 30, frmLogin.Source, 147, 0, 2)
End Sub

Private Sub cmdEdit_Click()
'    On Error GoTo ErrorEdit
    EditMode = True
    Press_Button ("EDIT")
    datPrimaryRS.Recordset.Edit
'ErrorEdit:
'    Call MessageBox("frmSelling", Err.Description, 0)
End Sub

Private Sub cmdEdit_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call MacButton("    Edit", frmSelling.cmdEdit, 0, 0, 73, 50, frmLogin.Source, 74, 0, 1)
End Sub

Private Sub cmdEdit_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call MacButton("    Edit", frmSelling.cmdEdit, 0, 0, 73, 50, frmLogin.Source, 0, 0, 1)
End Sub

Private Sub cmdFind_Click()
'    On Error GoTo ErrorFind
    EditMode = True
    Call FindBox(" Find by Invoice Number ", _
                    "INVOICE", "INVOICE_NO", 0, 1, "frmLogin")
    Call MacButton("   Find", frmSelling.cmdFind, 0, 0, 73, 50, frmLogin.Source, 0, 0, 1)
'ErrorFind:
'    Call MessageBox("frmSelling", Err.Description, 0)
End Sub

Public Sub FindInvoiceNo()
'    On Error GoTo ErrorFindInvoiceNo
    datPrimaryRS.Recordset.FindFirst "INVOICE_NO >= '" & frmFind.txtWord & "'"
'ErrorFindInvoiceNo:
'    Call MessageBox("frmSelling", Err.Description, 0)
End Sub


Private Sub cmdFind_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call MacButton("   Find", frmSelling.cmdFind, 0, 0, 73, 50, frmLogin.Source, 74, 0, 1)
End Sub

Private Sub cmdFind_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call MacButton("   Find", frmSelling.cmdFind, 0, 0, 73, 50, frmLogin.Source, 0, 0, 1)
End Sub

Private Sub cmdLast_Click()
    On Error Resume Next
    Press_Button ("LAST")
End Sub

Private Sub cmdLast_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call MacButton("q", frmSelling.cmdLast, 0, 0, 100, 50, frmLogin.Source, 112, 39, 3)
End Sub

Private Sub cmdLast_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call MacButton("q", frmSelling.cmdLast, 0, 0, 100, 50, frmLogin.Source, 138, 39, 3)
End Sub

Private Sub cmdNext_Click()
    On Error Resume Next
    Press_Button ("NEXT")
End Sub

Private Sub cmdNext_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call MacButton("u", frmSelling.cmdNext, 0, 0, 100, 49, frmLogin.Source, 112, 39, 3)
End Sub

Private Sub cmdNext_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call MacButton("u", frmSelling.cmdNext, 0, 0, 100, 49, frmLogin.Source, 138, 39, 3)
End Sub

Private Sub cmdPrev_Click()
    On Error Resume Next
    Press_Button ("PREV")
End Sub

Private Sub cmdPrev_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call MacButton("t", frmSelling.cmdPrev, 0, 0, 100, 50, frmLogin.Source, 112, 39, 3)
End Sub

Private Sub cmdPrev_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call MacButton("t", frmSelling.cmdPrev, 0, 0, 100, 50, frmLogin.Source, 138, 39, 3)
End Sub

Private Sub cmdPrint_Click()
'    On Error GoTo ErrorPrint
    If txtPaid = "" Then
        Call MessageBox("frmSelling", "Cannot print if not yet paid..", 0)
    Else
        Call frmSelling.PrintInvoice
    End If
    Call MacButton("   Print", frmSelling.cmdPrint, 0, 0, 73, 50, frmLogin.Source, 0, 0, 1)
'ErrorPrint:
End Sub

Private Sub cmdPrint_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call MacButton("   Print", frmSelling.cmdPrint, 0, 0, 73, 50, frmLogin.Source, 74, 0, 1)
End Sub

Private Sub cmdPrint_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call MacButton("   Print", frmSelling.cmdPrint, 0, 0, 73, 50, frmLogin.Source, 0, 0, 1)
End Sub

Private Sub cmdHelp_Click()
'    On Error GoTo ErrorHelp
    Call MacButton("   Help", frmSelling.cmdHelp, 0, 0, 73, 50, frmLogin.Source, 0, 0, 1)
'ErrorHelp:
'    Call MessageBox("frmSelling", Err.Description, 0)
End Sub

Private Sub cmdHelp_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call MacButton("   Help", frmSelling.cmdHelp, 0, 0, 73, 50, frmLogin.Source, 74, 0, 1)
End Sub

Private Sub cmdHelp_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call MacButton("   Help", frmSelling.cmdHelp, 0, 0, 73, 50, frmLogin.Source, 0, 0, 1)
End Sub

Private Sub cmdSave_Click()
'    On Error GoTo ErrorSave
    If EditMode = True Then
        Call MacButton("    Edit", frmSelling.cmdEdit, 0, 0, 73, 50, frmLogin.Source, 0, 0, 1)
        Press_Button ("SAVE")
    Else
        Call MacButton("   New", frmSelling.cmdAdd, 0, 0, 73, 50, frmLogin.Source, 0, 0, 1)
        If Search_Exist(txtField(0), "INVOICE_NO", "INVOICE") Then
            Call MessageBox("frmSelling", "Invoice # already exist.", 0)
            frmMessageBox.SetFocus
            Press_Button ("UNDO")
        ElseIf txtField(0) = "" Then
            Call MessageBox("frmSelling", "Invoice # cannot be null", 0)
            frmMessageBox.SetFocus
            Press_Button ("UNDO")
            'ENABLE ONLY FOR SPECIFIC LENGTH OF 5 CHARACTERS NUMBERING
            'ElseIf Len(txtField(0)) > 1 And Len(txtField(0)) < 5 Then
            '    Call MessageBox("frmSelling", "Product number must be 5 in length", 0)
            '    frmMessageBox.SetFocus
            '    Press_Button ("UNDO")
        Else
            Press_Button ("SAVE")
        End If
    End If
    Call MacButton("   Save", frmSelling.cmdSave, 0, 0, 73, 50, frmLogin.Source, 0, 0, 1)
'ErrorSave:
'        Call MessageBox("frmSelling", Err.Description, 0)
End Sub

Private Sub cmdSave_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call MacButton("   Save", frmSelling.cmdSave, 0, 0, 73, 50, frmLogin.Source, 74, 0, 1)
End Sub

Private Sub cmdSave_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call MacButton("   Save", frmSelling.cmdSave, 0, 0, 73, 50, frmLogin.Source, 0, 0, 1)
End Sub

Private Sub cmdStop_Click()
    txtField(0).TabStop = True
    txtField(1).TabStop = True
    DTPick.TabStop = True
    txtTotal.TabStop = True
    txtPaid.TabStop = True
    txtChange.TabStop = True
    txtWord.TabStop = False
    txtQuantitySold.TabStop = False
    DetailBoxContainer.Visible = False
    Call MacButton("   Close", frmSelling.cmdStop, 0, 0, 73, 50, frmLogin.Source, 0, 0, 1)
End Sub

Private Sub cmdStop_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call MacButton("   Close", frmSelling.cmdStop, 0, 0, 73, 50, frmLogin.Source, 74, 0, 1)
End Sub

Private Sub cmdStop_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call MacButton("   Close", frmSelling.cmdStop, 0, 0, 73, 50, frmLogin.Source, 0, 0, 1)
End Sub

Private Sub cmdTag_Click()
    On Error Resume Next
    If Valid_Fields Then
        invoice.AddNew
        invoice("INVOICE_NOD") = txtField(0)
        invoice("QTYSOLD") = txtQuantitySold
        invoice("PRODCODE") = txtProductCode
        invoice.Update
        Call Update_Stocks(frmSelling.txtProductCode, False, frmSelling.txtQuantitySold)
        txtAvailableStock = Convert_Numeric(CDbl(frmSelling.txtAvailableStock) - CDbl(frmSelling.txtQuantitySold), False)
        txtTotalCost = Convert_Numeric(CDbl(frmSelling.txtUnitCost) * CDbl(frmSelling.txtAvailableStock), True)
        Set invoice = frmLogin.db.OpenRecordset("select * from INVOICE_DETAIL where INVOICE_NOD = '" & txtField(0) & "'")
        datSecondaryRS.Refresh
        txtTotal = dspTotal
    End If
    Call MacButton("    Tag", frmSelling.cmdTag, 0, 0, 73, 50, frmLogin.Source, 0, 0, 1)
End Sub

Private Sub cmdTag_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call MacButton("    Tag", frmSelling.cmdTag, 0, 0, 73, 50, frmLogin.Source, 74, 0, 1)
End Sub

Private Sub cmdTag_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call MacButton("    Tag", frmSelling.cmdTag, 0, 0, 73, 50, frmLogin.Source, 0, 0, 1)
End Sub

Function Valid_Fields() As Boolean
    On Error Resume Next
    Dim r_val As Boolean
    r_val = True
    If txtField(0) = "" Then
        Call MessageBox("frmSelling", "Invoice Number required.", 0)
        frmMessageBox.SetFocus
        r_val = False
    End If
    If txtProductCode = "" Then
        r_val = False
        Call MessageBox("frmSelling", "Product Number required.", 0)
        frmMessageBox.SetFocus
    End If
    If txtQuantitySold = "" Then
        r_val = False
    End If
    If CDbl(txtAvailableStock) < CDbl(txtQuantitySold) Then
        Call MessageBox("frmSelling", "Quantity Sold must be less than available product.", 0)
        frmMessageBox.SetFocus
        txtQuantitySold.SetFocus
        r_val = False
    End If
    Valid_Fields = r_val
End Function

Private Sub cmdReturn_Click()
    On Error Resume Next
    Call Update_Stocks(txtProductCode, True, txtQuantitySold)
    Delete_Detail
    txtAvailableStock = Convert_Numeric(CDbl(txtAvailableStock) + CDbl(txtQuantitySold), False)
    txtTotalCost = Convert_Numeric(CDbl(txtUnitCost) * CDbl(txtAvailableStock), True)
    Set invoice = frmLogin.db.OpenRecordset("select * from INVOICE_DETAIL where INVOICE_NOD = '" & txtField(0) & "'")
    datSecondaryRS.Refresh
    txtTotal = dspTotal
    Call MacButton("  Return", frmSelling.cmdReturn, 0, 0, 73, 50, frmLogin.Source, 0, 0, 1)
End Sub

Private Sub cmdReturn_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call MacButton("  Return", frmSelling.cmdReturn, 0, 0, 73, 50, frmLogin.Source, 74, 0, 1)
End Sub

Private Sub cmdReturn_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call MacButton("  Return", frmSelling.cmdReturn, 0, 0, 73, 50, frmLogin.Source, 0, 0, 1)
End Sub

Private Sub Delete_Detail()
    On Error Resume Next
    strs = "select * from INVOICE_DETAIL where PRODCODE = '" & txtProductCode & "'"
    Set dummy = frmLogin.db.OpenRecordset(strs)
    If dummy.AbsolutePosition <> -1 Then
        With dummy
            .Edit
            .Fields("QTYSOLD") = .Fields("QTYSOLD") - txtQuantitySold
            .Update
            If .Fields("QTYSOLD") < 1 Then
                .Delete
                .MoveFirst
            End If
        End With
    End If
End Sub

Private Sub cmdTop_Click()
    On Error Resume Next
    Press_Button ("TOP")
End Sub

Private Sub cmdTop_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call MacButton("p", frmSelling.cmdTop, 0, 0, 100, 50, frmLogin.Source, 112, 39, 3)
End Sub

Private Sub cmdTop_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call MacButton("p", frmSelling.cmdTop, 0, 0, 100, 50, frmLogin.Source, 138, 39, 3)
End Sub

Private Sub cmdTotal_Click()
    txtPaid.SetFocus
    Call MacButton("     Total-Payment", frmSelling.cmdTotal, 0, 0, 170, 30, frmLogin.Source, 147, 0, 2)
End Sub

Private Sub cmdTotal_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call MacButton("     Total-Payment", frmSelling.cmdTotal, 0, 0, 170, 30, frmLogin.Source, 182, 30, 2)
End Sub

Private Sub cmdTotal_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call MacButton("     Total-Payment", frmSelling.cmdTotal, 0, 0, 170, 30, frmLogin.Source, 147, 0, 2)
End Sub

Private Sub cmdUndo_Click()
'    On Error GoTo ErrorUndo
    If EditMode = True Then
        Call MacButton("    Edit", frmSelling.cmdEdit, 0, 0, 73, 50, frmLogin.Source, 0, 0, 1)
    Else
        Call MacButton("   New", frmSelling.cmdAdd, 0, 0, 73, 50, frmLogin.Source, 0, 0, 1)
    End If
    Press_Button ("UNDO")
    Call MacButton("   Undo", frmSelling.cmdUndo, 0, 0, 73, 50, frmLogin.Source, 0, 0, 1)
'ErrorUndo:
End Sub

Private Sub cmdUndo_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call MacButton("   Undo", frmSelling.cmdUndo, 0, 0, 73, 50, frmLogin.Source, 74, 0, 1)
End Sub

Private Sub cmdUndo_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call MacButton("   Undo", frmSelling.cmdUndo, 0, 0, 73, 50, frmLogin.Source, 0, 0, 1)
End Sub

Private Sub datPrimaryRS_Error(DataErr As Integer, Response As Integer)
    Call MessageBox("frmSelling", "Data error event hit err:" & Error$(DataErr), 0)
    Response = 0
End Sub

Private Sub datPrimaryRS_Reposition()
    On Error Resume Next
    Screen.MousePointer = vbDefault
    recTxt.Text = "Record: " & (datPrimaryRS.Recordset.AbsolutePosition + 1)
    datSecondaryRS.RecordSource = " select DESCRIPTION, QTY_SOLD, AMOUNT from QRYINVOICE where INVOICE_NO = '" & datPrimaryRS.Recordset![INVOICE_NO] & "'"
    datSecondaryRS.Refresh
End Sub

Private Sub Form_Activate()
    On Error Resume Next
    Enable_Buttons
    Enable_Buttons
    datPrimaryRS.DatabaseName = App.Path & "\DATABASE\POSDATA.MDB"
    datSecondaryRS.DatabaseName = App.Path & "\DATABASE\POSDATA.MDB"
    txtTotal = dspTotal
End Sub

Private Sub Form_Load()
    On Error Resume Next
    Call ColForm(BoxContainer, 217, 211, 213, 125)
    Call ColForm(CounterContainer, 217, 211, 213, 125)
    Call ColForm(DetailsContainer, 217, 211, 213, 125)
    Call ColForm(DetailBoxContainer, 217, 211, 213, 125)
    Call ColForm(ButtonContainer, 217, 211, 213, 125)
    Call CreateMacOSTitleBar(titleBar, " Selling Transaction ")
    Call CreateMacOSTitleBar(titleBar2, " Tag/Return Items Selection ")
    Call BitBlt(frmSelling.Closed.hDC, 0, 0, 73, 50, frmLogin.Source.hDC, 0, 107, SRCCOPY)
    frmMain.Closed.Refresh
    Call BitBlt(frmSelling.Maximized.hDC, 0, 0, 73, 50, frmLogin.Source.hDC, 0, 72, SRCCOPY)
    frmMain.Maximized.Refresh
    Call BitBlt(frmSelling.Minimized.hDC, 0, 0, 73, 50, frmLogin.Source.hDC, 0, 124, SRCCOPY)
    frmMain.Minimized.Refresh
    Call MacButton("   New", frmSelling.cmdAdd, 0, 0, 73, 50, frmLogin.Source, 0, 0, 1)
    Call MacButton("    Edit", frmSelling.cmdEdit, 0, 0, 73, 50, frmLogin.Source, 0, 0, 1)
    Call MacButton("   Save", frmSelling.cmdSave, 0, 0, 73, 50, frmLogin.Source, 0, 0, 1)
    Call MacButton("   Undo", frmSelling.cmdUndo, 0, 0, 73, 50, frmLogin.Source, 0, 0, 1)
    Call MacButton("  Delete", frmSelling.cmdDelete, 0, 0, 73, 50, frmLogin.Source, 0, 0, 1)
    Call MacButton("   Print", frmSelling.cmdPrint, 0, 0, 73, 50, frmLogin.Source, 0, 0, 1)
    Call MacButton("   Help", frmSelling.cmdHelp, 0, 0, 73, 50, frmLogin.Source, 0, 0, 1)
    Call MacButton("   Find", frmSelling.cmdFind, 0, 0, 73, 50, frmLogin.Source, 0, 0, 1)
    Call MacButton("    Exit", frmSelling.cmdClose, 0, 0, 73, 50, frmLogin.Source, 0, 0, 1)
    Call MacButton("    Tag", frmSelling.cmdTag, 0, 0, 73, 50, frmLogin.Source, 0, 0, 1)
    Call MacButton("  Return", frmSelling.cmdReturn, 0, 0, 73, 50, frmLogin.Source, 0, 0, 1)
    Call MacButton("   Close", frmSelling.cmdStop, 0, 0, 73, 50, frmLogin.Source, 0, 0, 1)
    Call MacButton("   Tag/Return Items", frmSelling.cmdDetail, 0, 0, 170, 30, frmLogin.Source, 147, 0, 2)
    Call MacButton("     Total-Payment", frmSelling.cmdTotal, 0, 0, 170, 30, frmLogin.Source, 147, 0, 2)
    Call MacButton("p", frmSelling.cmdTop, 0, 0, 100, 50, frmLogin.Source, 138, 39, 3)
    Call MacButton("u", frmSelling.cmdNext, 0, 0, 100, 50, frmLogin.Source, 138, 39, 3)
    Call MacButton("t", frmSelling.cmdPrev, 0, 0, 100, 50, frmLogin.Source, 138, 39, 3)
    Call MacButton("q", frmSelling.cmdLast, 0, 0, 100, 50, frmLogin.Source, 138, 39, 3)
    KeyPreview = True
    datPrimaryRS.DatabaseName = App.Path & "\DATABASE\POSDATA.MDB"
    datSecondaryRS.DatabaseName = App.Path & "\DATABASE\POSDATA.MDB"
    Set datprimary2 = frmLogin.db.OpenRecordset("select * from SETUP order by COMPANY_NAME")
    Set invoice = frmLogin.db.OpenRecordset("select * from INVOICE_DETAIL where INVOICE_NOD = '" & txtField(0) & "'")
    p_isadding = False
    p_isadding = False
    p_isediting = False
    Enable_Fields (True)
'    frmLogin.Show
'    frmLogin.Hide
    frmSelling.txtWord.SelLength = Len(frmSelling.txtWord.Text)
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
                    Call MacButton("   New", frmSelling.cmdAdd, 0, 0, 73, 50, frmLogin.Source, 0, 0, 1)
                End If
        Case vbKeyE:
                If AltDown Then
                    Call MacButton("    Edit", frmSelling.cmdEdit, 0, 0, 73, 50, frmLogin.Source, 0, 0, 1)
                End If
        Case vbKeyS:
                If AltDown Then
                    Call MacButton("   Save", frmSelling.cmdSave, 0, 0, 73, 50, frmLogin.Source, 0, 0, 1)
                End If
        Case vbKeyU:
                If AltDown Then
                    Call MacButton("   Undo", frmSelling.cmdUndo, 0, 0, 73, 50, frmLogin.Source, 0, 0, 1)
                End If
        Case vbKeyD:
                If AltDown Then
                    Call MacButton("  Delete", frmSelling.cmdDelete, 0, 0, 73, 50, frmLogin.Source, 0, 0, 1)
                End If
        Case vbKeyP:
                If AltDown Then
                    Call MacButton("   Print", frmSelling.cmdPrint, 0, 0, 73, 50, frmLogin.Source, 0, 0, 1)
                End If
        Case vbKeyH:
                If AltDown Then
                    Call MacButton("   Help", frmSelling.cmdHelp, 0, 0, 73, 50, frmLogin.Source, 0, 0, 1)
                End If
        Case vbKeyF:
                If AltDown Then
                    Call MacButton("   Find", frmSelling.cmdFind, 0, 0, 73, 50, frmLogin.Source, 0, 0, 1)
                End If
        Case vbKeyX:
                If AltDown Then
                    Call MacButton("    Exit", frmSelling.cmdClose, 0, 0, 73, 50, frmLogin.Source, 0, 0, 1)
                End If
        Case vbKeyI:
                If AltDown Then
                    Call MacButton("   Tag/Return Items", frmSelling.cmdDetail, 0, 0, 170, 30, frmLogin.Source, 147, 0, 2)
                End If
        Case vbKeyY:
                If AltDown Then
                    Call MacButton("     Total-Payment", frmSelling.cmdTotal, 0, 0, 170, 30, frmLogin.Source, 147, 0, 2)
                End If
        Case vbKeyT:
                If AltDown Then
                    Call MacButton("    Tag", frmSelling.cmdTag, 0, 0, 73, 50, frmLogin.Source, 0, 0, 1)
                End If
        Case vbKeyR:
                If AltDown Then
                    Call MacButton("  Return", frmSelling.cmdReturn, 0, 0, 73, 50, frmLogin.Source, 0, 0, 1)
                End If
        Case vbKeyC:
                If AltDown Then
                    Call MacButton("   Close", frmSelling.cmdStop, 0, 0, 73, 50, frmLogin.Source, 0, 0, 1)
                End If
        Case vbKeyLeft:
                If AltDown Then
                    Call MacButton("t", frmSelling.cmdPrev, 0, 0, 100, 50, frmLogin.Source, 138, 39, 3)
                End If
        Case vbKeyRight:
                If AltDown Then
                    Call MacButton("u", frmSelling.cmdNext, 0, 0, 100, 49, frmLogin.Source, 138, 39, 3)
                End If
        Case vbKeyUp:
                If AltDown Then
                    Call MacButton("p", frmSelling.cmdTop, 0, 0, 100, 50, frmLogin.Source, 138, 39, 3)
                End If
        Case vbKeyDown:
                If AltDown Then
                    Call MacButton("q", frmSelling.cmdLast, 0, 0, 100, 50, frmLogin.Source, 138, 39, 3)
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
                                      "Note: The following are Key Shortcuts." & vbCrLf & _
                                      Space(1) & vbCrLf & _
                                      "ALT-N=New, ALT-E=Edit, ALT-S=Save, ALT-U=Undo, ALT-D=Delete" & vbCrLf & _
                                      "ALT-F=Find, ALT-X=Exit, ALT-P=Print" & vbCrLf & _
                                      "ALT-I=Tag/Return Items, ALT-Y=Total-Payment, " & vbCrLf & _
                                      Space(1) & vbCrLf & _
                                      "Left Arrow=Previous Records, Right Arrow=Next Records" & vbCrLf & _
                                      "Top Arrow=Top Record, Down Arrow=Last Record" & vbCrLf & _
                                      Space(1) & vbCrLf & _
                                      "Tag/Return Item Selection Key Shortcuts." & vbCrLf & _
                                      "ALT-T=Tag, ALT-R=Return, ALT-C=Close"
        Case vbKeyN:
                If AltDown Then
                    Call MacButton("   New", frmSelling.cmdAdd, 0, 0, 73, 50, frmLogin.Source, 74, 0, 1)
                    cmdAdd_Click
                End If
        Case vbKeyE:
                If AltDown Then
                    Call MacButton("    Edit", frmSelling.cmdEdit, 0, 0, 73, 50, frmLogin.Source, 74, 0, 1)
                    cmdEdit_Click
                End If
        Case vbKeyS:
                If AltDown Then
                    Call MacButton("   Save", frmSelling.cmdSave, 0, 0, 73, 50, frmLogin.Source, 74, 0, 1)
                    cmdSave_Click
                End If
        Case vbKeyU:
                If AltDown Then
                    Call MacButton("   Undo", frmSelling.cmdUndo, 0, 0, 73, 50, frmLogin.Source, 74, 0, 1)
                    cmdUndo_Click
                End If
        Case vbKeyD:
                If AltDown Then
                    Call MacButton("  Delete", frmSelling.cmdDelete, 0, 0, 73, 50, frmLogin.Source, 74, 0, 1)
                    cmdDelete_Click
                End If
        Case vbKeyP:
                If AltDown Then
                    Call MacButton("   Print", frmSelling.cmdPrint, 0, 0, 73, 50, frmLogin.Source, 74, 0, 1)
                    cmdPrint_Click
                End If
        Case vbKeyH:
                If AltDown Then
                    Call MacButton("   Help", frmSelling.cmdHelp, 0, 0, 73, 50, frmLogin.Source, 74, 0, 1)
                    cmdHelp_Click
                End If
        Case vbKeyF:
                If AltDown Then
                    Call MacButton("   Find", frmSelling.cmdFind, 0, 0, 73, 50, frmLogin.Source, 74, 0, 1)
                    cmdFind_Click
                End If
        Case vbKeyX:
                If AltDown Then
                    Call MacButton("    Exit", frmSelling.cmdClose, 0, 0, 73, 50, frmLogin.Source, 74, 0, 1)
                    cmdClose_Click
                End If
        Case vbKeyI:
                If AltDown Then
                    Call MacButton("   Tag/Return Items", frmSelling.cmdDetail, 0, 0, 170, 30, frmLogin.Source, 182, 30, 2)
                    cmdDetail_Click
                End If
        Case vbKeyY:
                If AltDown Then
                    Call MacButton("     Total-Payment", frmSelling.cmdTotal, 0, 0, 170, 30, frmLogin.Source, 182, 30, 2)
                    cmdTotal_Click
                End If
        Case vbKeyT:
                If AltDown Then
                    Call MacButton("    Tag", frmSelling.cmdTag, 0, 0, 73, 50, frmLogin.Source, 74, 0, 1)
                    cmdTag_Click
                End If
        Case vbKeyR:
                If AltDown Then
                    Call MacButton("  Return", frmSelling.cmdReturn, 0, 0, 73, 50, frmLogin.Source, 74, 0, 1)
                    cmdReturn_Click
                End If
        Case vbKeyC:
                If AltDown Then
                    Call MacButton("   Close", frmSelling.cmdStop, 0, 0, 73, 50, frmLogin.Source, 74, 0, 1)
                    cmdStop_Click
                End If
        Case vbKeyLeft:
                If AltDown Then
                    Call MacButton("t", frmSelling.cmdPrev, 0, 0, 100, 50, frmLogin.Source, 112, 39, 3)
                    cmdPrev_Click
                End If
        Case vbKeyRight:
                If AltDown Then
                    Call MacButton("u", frmSelling.cmdNext, 0, 0, 100, 49, frmLogin.Source, 112, 39, 3)
                    cmdNext_Click
                End If
        Case vbKeyUp:
                If AltDown Then
                    Call MacButton("p", frmSelling.cmdTop, 0, 0, 100, 50, frmLogin.Source, 112, 39, 3)
                    cmdTop_Click
                End If
        Case vbKeyDown:
                If AltDown Then
                    Call MacButton("q", frmSelling.cmdLast, 0, 0, 100, 50, frmLogin.Source, 112, 39, 3)
                    cmdLast_Click
                End If
    End Select
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Screen.MousePointer = vbDefault
End Sub

Private Sub Enable_Fields(p As Boolean)
    txtField(0).Locked = p
End Sub

Public Sub Press_Button(p_type As String)
    On Error Resume Next
    Select Case p_type
            Case "ADD"
                p_save = True
                p_isadding = True
            Case "EDIT"
                p_isediting = True
                p_save = True
                p_isadding = False
            Case "SAVE"
                Update_ORNumber
                datPrimaryRS.UpdateRecord
                datPrimaryRS.Recordset.Bookmark = datPrimaryRS.Recordset.LastModified
                p_save = False
                p_isadding = False
                p_isediting = False
            Case "UNDO"
                p_save = False
                p_isadding = False
                p_isediting = False
                datPrimaryRS.Recordset.CancelUpdate
            Case "TOP"
                p_save = False
                p_isadding = False
                p_isediting = False
                datPrimaryRS.Recordset.MoveFirst
            Case "PREV"
                p_save = False
                p_isadding = False
                p_isediting = False
                datPrimaryRS.Recordset.MovePrevious
            Case "NEXT"
                p_save = False
                p_isadding = False
                p_isediting = False
                datPrimaryRS.Recordset.MoveNext
            Case "LAST"
                p_save = False
                p_isadding = False
                p_isediting = False
                datPrimaryRS.Recordset.MoveLast
            Case "DELETE"
                p_save = False
                p_isadding = False
                p_isediting = False
                Delete_Details
                With datPrimaryRS.Recordset
                    .Delete
                    .MoveNext
                    If .RecordCount > 0 And .EOF Then .MoveLast
                End With
    End Select
    Enable_Fields (Not p_save)
    Enable_Buttons
    txtTotal = dspTotal
    TerminalDescription = " Total Amount Due"
End Sub

Public Sub Enable_Buttons()
    On Error Resume Next
    Dim cur_rec, fst_rec, lst_rec, rec_cnt As Integer
    Dim mark_rec As Variant
    rec_cnt = datPrimaryRS.Recordset.RecordCount
    If rec_cnt > 0 Then
        If Not datPrimaryRS.Recordset.BOF Or Not datPrimaryRS.Recordset.EOF Then
            cur_rec = datPrimaryRS.Recordset.AbsolutePosition + 1
            mark_rec = datPrimaryRS.Recordset.Bookmark
        End If
        datPrimaryRS.Recordset.MoveFirst
        fst_rec = datPrimaryRS.Recordset.AbsolutePosition + 1
        datPrimaryRS.Recordset.MoveLast
        lst_rec = datPrimaryRS.Recordset.AbsolutePosition + 1
        If Not datPrimaryRS.Recordset.BOF Or Not datPrimaryRS.Recordset.EOF Then
            datPrimaryRS.Recordset.Bookmark = mark_rec
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
    cmdAdd.Enabled = p_add
    cmdEdit.Enabled = p_edit
    cmdSave.Enabled = p_save
    cmdUndo.Enabled = p_undo
    cmdTop.Enabled = p_top
    cmdPrev.Enabled = p_prev
    cmdNext.Enabled = p_next
    cmdLast.Enabled = p_last
    cmdDelete.Enabled = p_del
    If p_del Then
        cmdFind.Enabled = IIf(rec_cnt > 1, True, False)
    Else
        cmdFind.Enabled = False
    End If
        cmdDetail.Enabled = p_add
End Sub

Private Sub titleBar_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call DragForm(Me)
End Sub


Private Sub txtChange_Change()
    TerminalCounter = txtChange
End Sub

Private Sub txtChange_GotFocus()
    TerminalDescription = " Change"
    TerminalCounter = txtChange
End Sub

Private Sub Update_ORNumber()
    On Error Resume Next
    strs = "select * from CODE_FILE where CODE_NAME = 'ORNO'"
    Set dummy = frmLogin.db.OpenRecordset(strs)
    If dummy.AbsolutePosition <> -1 Then
        dummy.Edit
        dummy("CODE_VALUE") = CStr(CInt(dummy("CODE_VALUE")) + 1)
        dummy.Update
    End If
End Sub

Function Get_OrNo() As String
    On Error Resume Next
    strs = "select * from CODE_FILE where CODE_NAME = 'ORNO'"
    Set dummy = frmLogin.db.OpenRecordset(strs)
    If dummy.AbsolutePosition <> -1 Then
       Get_OrNo = Pad_Str(dummy("CODE_VALUE"), "0", 5, False)
    Else
       Call MessageBox("frmSelling", "Warning Or number not found in Code File!!", 0)
       frmMessageBox.SetFocus
    End If
End Function

Private Sub Delete_Details()
    On Error Resume Next
    strs = "select * from INVOICE_DETAIL where INVOICE_NOD = '" & txtField(0) & "'"
    Set dummy = frmLogin.db.OpenRecordset(strs)
    Do While Not dummy.EOF
        If dummy.AbsolutePosition <> -1 Then
            Call Update_Stocks(dummy("PRODCODE"), True, dummy("QTYSOLD"))
            dummy.Delete
        Else
             Exit Do
        End If
    Loop
End Sub

Function dspTotal() As String
    On Error Resume Next
    strs = "select sum(AMOUNT) from QRYINVOICE where INVOICE_NO = '" & IIf(txtField(0) = "", "0", txtField(0)) & "'"
    Set dummy = frmLogin.db.OpenRecordset(strs)
    dspTotal = Convert_Numeric(IIf(IsNull(dummy(0)), "0", dummy(0)), True)
    txtTotal.Text = dspTotal
    TerminalDescription = " Total Amount Due"
End Function

Function PrintInvoice()
    On Error Resume Next
    Today = Now
    datSecondaryRS.Recordset.MoveFirst
        Printer.FontName = "Courier New"
        Printer.FontSize = 10
        Printer.Print IIf(IsNull(datprimary2("COMPANY_NAME")), "", Left(datprimary2("COMPANY_NAME"), 28))
        Printer.FontSize = 8
        Printer.Print IIf(IsNull(datprimary2("COMPANY_ADDRESS")), "", datprimary2("COMPANY_ADDRESS"))
        Printer.Print IIf(IsNull(datprimary2("COMPANY_TELEPHONE")), "", datprimary2("COMPANY_TELEPHONE"))
        Printer.Print IIf(IsNull(datprimary2("SELLING_HEADER")), "", datprimary2("SELLING_HEADER"))
        Printer.Print Space(1)
        Printer.Print "              THIS IS AN OFFICIAL          "
        Printer.Print "                    RECEIPT                "
        Printer.Print Space(1)
        Printer.Print "Inv. #" + txtField(0)
        Printer.Print "OR #" + txtField(1)
        Printer.Print "==========================================="
        Printer.Print "DESCRIPTION                 QTY      AMOUNT"
        Printer.Print "==========================================="
    Do While Not datSecondaryRS.Recordset.EOF
        Printer.Print Pad_Str(grdDataGrid.Columns(0), " ", 27, True) + _
                      Space(1) + Space(5 - Len(Convert_Numeric(grdDataGrid.Columns(1), False))) + _
                      Convert_Numeric(grdDataGrid.Columns(1), False) + _
                      Space(1) + Space(9 - Len(Convert_Numeric(grdDataGrid.Columns(2), True))) + _
                      Convert_Numeric(grdDataGrid.Columns(2), True)
        datSecondaryRS.Recordset.MoveNext
    Loop
        Printer.Print "                                 ----------"
        Printer.Print Pad_Str("TOTAL", " ", 27, True) + Space(7) + Space(9 - Len(IIf(Convert_Numeric(txtTotal, True) = 0, "", Convert_Numeric(txtTotal, True)))) + _
                        IIf(Convert_Numeric(txtTotal, True) = 0, "", Convert_Numeric(txtTotal, True))
        Printer.Print Pad_Str("CASH", " ", 27, True) + Space(7) + Space(9 - Len(IIf(Convert_Numeric(txtPaid, True) = 0, "", Convert_Numeric(txtPaid, True)))) + _
                        IIf(Convert_Numeric(txtPaid, True) = 0, "", Convert_Numeric(txtPaid, True))
        Printer.Print Pad_Str("CHANGE", " ", 27, True) + Space(7) + Space(9 - Len(IIf(Convert_Numeric(txtChange, True) = 0, "", Convert_Numeric(txtChange, True)))) + _
                        IIf(Convert_Numeric(txtChange, True) = 0, "", Convert_Numeric(txtChange, True))
        Printer.Print "                                 =========="
        Printer.Print Format(Today, "dd " & "mmmm " & "yyyy ") & Space(1) & Format(Today, "hh:mm:ss ampm ")
        Printer.Print frmLogin.txtUserName
        Printer.Print IIf(IsNull(datprimary2("SELLING_FOOTER")), "", datprimary2("SELLING_FOOTER"))
        Printer.EndDoc
End Function

Private Sub txtPaid_Change()
    txtChange = Convert_Numeric(txtPaid, True) - Convert_Numeric(txtTotal, True)
    TerminalCounter = txtPaid
End Sub

Private Sub txtPaid_GotFocus()
    TerminalDescription = " Amount Paid"
    TerminalCounter = txtPaid
End Sub

Private Sub txtTotal_Change()
    TerminalCounter = txtTotal
End Sub

Private Sub txtTotal_GotFocus()
    TerminalDescription = " Total Amount Due"
    TerminalCounter = txtTotal
End Sub

Private Sub lstWords_Click()
    On Error Resume Next
    Call QueryDetail(lstWords.List(lstWords.ListIndex))
End Sub

Private Sub txtWord_Change()
    On Error Resume Next
    If txtWord = "" Then
        lstWords.Visible = True
    End If
    Call QueryDetail(txtWord.Text)
End Sub


