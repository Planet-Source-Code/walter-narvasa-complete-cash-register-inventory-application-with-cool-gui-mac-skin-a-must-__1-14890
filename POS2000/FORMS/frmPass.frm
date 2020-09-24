VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmPassword 
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
      TabIndex        =   19
      TabStop         =   0   'False
      Top             =   40
      Width           =   9220
      Begin VB.PictureBox CheckContainer 
         BackColor       =   &H00808080&
         BorderStyle     =   0  'None
         Height          =   2655
         Left            =   240
         ScaleHeight     =   2655
         ScaleWidth      =   8775
         TabIndex        =   20
         TabStop         =   0   'False
         Top             =   2280
         Width           =   8775
         Begin VB.PictureBox cmdTop 
            Appearance      =   0  'Flat
            AutoRedraw      =   -1  'True
            BackColor       =   &H00CCCCCC&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   300
            Left            =   1440
            MouseIcon       =   "frmPass.frx":0000
            MousePointer    =   99  'Custom
            ScaleHeight     =   300
            ScaleWidth      =   390
            TabIndex        =   55
            TabStop         =   0   'False
            Top             =   2160
            Width           =   390
         End
         Begin VB.PictureBox cmdPrev 
            Appearance      =   0  'Flat
            AutoRedraw      =   -1  'True
            BackColor       =   &H00CCCCCC&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   300
            Left            =   1920
            MouseIcon       =   "frmPass.frx":030A
            MousePointer    =   99  'Custom
            ScaleHeight     =   300
            ScaleWidth      =   390
            TabIndex        =   54
            TabStop         =   0   'False
            Top             =   2160
            Width           =   390
         End
         Begin VB.PictureBox cmdNext 
            Appearance      =   0  'Flat
            AutoRedraw      =   -1  'True
            BackColor       =   &H00CCCCCC&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   300
            Left            =   2400
            MouseIcon       =   "frmPass.frx":0614
            MousePointer    =   99  'Custom
            ScaleHeight     =   300
            ScaleWidth      =   390
            TabIndex        =   53
            TabStop         =   0   'False
            Top             =   2160
            Width           =   390
         End
         Begin VB.PictureBox cmdLast 
            Appearance      =   0  'Flat
            AutoRedraw      =   -1  'True
            BackColor       =   &H00CCCCCC&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   300
            Left            =   2880
            MouseIcon       =   "frmPass.frx":091E
            MousePointer    =   99  'Custom
            ScaleHeight     =   300
            ScaleWidth      =   390
            TabIndex        =   52
            TabStop         =   0   'False
            Top             =   2160
            Width           =   390
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
            Index           =   14
            Left            =   4920
            TabIndex        =   18
            Top             =   1920
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
            Index           =   13
            Left            =   4920
            TabIndex        =   17
            Top             =   1680
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
            Index           =   12
            Left            =   4920
            TabIndex        =   16
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
            Index           =   11
            Left            =   4920
            TabIndex        =   15
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
            Index           =   10
            Left            =   4920
            TabIndex        =   14
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
            Index           =   9
            Left            =   4920
            TabIndex        =   13
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
            Index           =   8
            Left            =   4920
            TabIndex        =   12
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
            Index           =   7
            Left            =   4920
            TabIndex        =   11
            Top             =   240
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
            Index           =   6
            Left            =   480
            TabIndex        =   10
            Top             =   1680
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
            Index           =   5
            Left            =   480
            TabIndex        =   9
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
            Left            =   480
            TabIndex        =   8
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
            Left            =   480
            TabIndex        =   7
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
            Left            =   480
            TabIndex        =   6
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
            Left            =   480
            TabIndex        =   5
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
            Left            =   480
            TabIndex        =   4
            Top             =   240
            Width           =   200
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
            TabIndex        =   56
            Top             =   2160
            Width           =   765
         End
         Begin VB.Label Label 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackColor       =   &H00808080&
            BackStyle       =   0  'Transparent
            Caption         =   "Software Setup"
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
            Index           =   19
            Left            =   5175
            TabIndex        =   51
            Top             =   1920
            Width           =   1530
         End
         Begin VB.Label Label 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackColor       =   &H00808080&
            BackStyle       =   0  'Transparent
            Caption         =   "Code File Setup"
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
            Index           =   18
            Left            =   5175
            TabIndex        =   50
            Top             =   1680
            Width           =   1500
         End
         Begin VB.Label Label 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackColor       =   &H00808080&
            BackStyle       =   0  'Transparent
            Caption         =   "Password Security"
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
            Index           =   17
            Left            =   5175
            TabIndex        =   49
            Top             =   1440
            Width           =   1830
         End
         Begin VB.Label Label 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackColor       =   &H00808080&
            BackStyle       =   0  'Transparent
            Caption         =   "Backup/Restore Files"
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
            Index           =   16
            Left            =   5175
            TabIndex        =   48
            Top             =   1200
            Width           =   2085
         End
         Begin VB.Label Label 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackColor       =   &H00808080&
            BackStyle       =   0  'Transparent
            Caption         =   "Supplier Listing Report"
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
            Left            =   5175
            TabIndex        =   47
            Top             =   960
            Width           =   2235
         End
         Begin VB.Label Label 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackColor       =   &H00808080&
            BackStyle       =   0  'Transparent
            Caption         =   "Product Listing Report"
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
            Left            =   5175
            TabIndex        =   46
            Top             =   720
            Width           =   2205
         End
         Begin VB.Label Label 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackColor       =   &H00808080&
            BackStyle       =   0  'Transparent
            Caption         =   "Stocks Per Supplier Report"
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
            Left            =   5175
            TabIndex        =   45
            Top             =   480
            Width           =   2610
         End
         Begin VB.Label Label 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackColor       =   &H00808080&
            BackStyle       =   0  'Transparent
            Caption         =   "Receiving History Report"
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
            Left            =   5160
            TabIndex        =   44
            Top             =   240
            Width           =   2430
         End
         Begin VB.Label Label 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackColor       =   &H00808080&
            BackStyle       =   0  'Transparent
            Caption         =   "Selling History Report"
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
            Left            =   720
            TabIndex        =   43
            Top             =   1680
            Width           =   2130
         End
         Begin VB.Label Label 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackColor       =   &H00808080&
            BackStyle       =   0  'Transparent
            Caption         =   "Stocks Re-order Report"
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
            Left            =   720
            TabIndex        =   42
            Top             =   1440
            Width           =   2310
         End
         Begin VB.Label Label 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackColor       =   &H00808080&
            BackStyle       =   0  'Transparent
            Caption         =   "Receiving Transaction"
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
            Left            =   720
            TabIndex        =   41
            Top             =   1200
            Width           =   2145
         End
         Begin VB.Label Label 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackColor       =   &H00808080&
            BackStyle       =   0  'Transparent
            Caption         =   "Selling Transaction"
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
            Left            =   720
            TabIndex        =   40
            Top             =   960
            Width           =   1845
         End
         Begin VB.Label Label 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackColor       =   &H00808080&
            BackStyle       =   0  'Transparent
            Caption         =   "Category Masterfile"
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
            Left            =   720
            TabIndex        =   39
            Top             =   720
            Width           =   1995
         End
         Begin VB.Label Label 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackColor       =   &H00808080&
            BackStyle       =   0  'Transparent
            Caption         =   "Product Masterfile"
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
            Left            =   720
            TabIndex        =   38
            Top             =   480
            Width           =   1860
         End
         Begin VB.Label Label 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackColor       =   &H00808080&
            BackStyle       =   0  'Transparent
            Caption         =   "Supplier Masterfile"
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
            Left            =   720
            TabIndex        =   37
            Top             =   240
            Width           =   1890
         End
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
         Left            =   4140
         TabIndex        =   3
         Text            =   "Combo1"
         Top             =   1800
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
         Left            =   4140
         MultiLine       =   -1  'True
         TabIndex        =   1
         Top             =   1080
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
         Index           =   0
         Left            =   4140
         Locked          =   -1  'True
         TabIndex        =   0
         Top             =   720
         Width           =   1695
      End
      Begin VB.PictureBox titleBar 
         BorderStyle     =   0  'None
         Height          =   375
         Left            =   120
         ScaleHeight     =   375
         ScaleWidth      =   9015
         TabIndex        =   29
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
            MouseIcon       =   "frmPass.frx":0C28
            MousePointer    =   99  'Custom
            ScaleHeight     =   285
            ScaleWidth      =   315
            TabIndex        =   32
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
            MouseIcon       =   "frmPass.frx":0F32
            MousePointer    =   99  'Custom
            ScaleHeight     =   285
            ScaleWidth      =   315
            TabIndex        =   31
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
            MouseIcon       =   "frmPass.frx":123C
            MousePointer    =   99  'Custom
            ScaleHeight     =   285
            ScaleWidth      =   315
            TabIndex        =   30
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
         TabIndex        =   21
         TabStop         =   0   'False
         Top             =   5040
         Width           =   8775
         Begin VB.PictureBox cmdExit 
            Appearance      =   0  'Flat
            AutoRedraw      =   -1  'True
            BackColor       =   &H00CCCCCC&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   570
            Left            =   7440
            MouseIcon       =   "frmPass.frx":1546
            MousePointer    =   99  'Custom
            ScaleHeight     =   570
            ScaleWidth      =   1095
            TabIndex        =   28
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
            MouseIcon       =   "frmPass.frx":1850
            MousePointer    =   99  'Custom
            ScaleHeight     =   570
            ScaleWidth      =   1095
            TabIndex        =   27
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
            MouseIcon       =   "frmPass.frx":1B5A
            MousePointer    =   99  'Custom
            ScaleHeight     =   570
            ScaleWidth      =   1095
            TabIndex        =   26
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
            MouseIcon       =   "frmPass.frx":1E64
            MousePointer    =   99  'Custom
            ScaleHeight     =   570
            ScaleWidth      =   1095
            TabIndex        =   25
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
            MouseIcon       =   "frmPass.frx":216E
            MousePointer    =   99  'Custom
            ScaleHeight     =   570
            ScaleWidth      =   1095
            TabIndex        =   24
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
            MouseIcon       =   "frmPass.frx":2478
            MousePointer    =   99  'Custom
            ScaleHeight     =   570
            ScaleWidth      =   1095
            TabIndex        =   23
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
            Left            =   6240
            MouseIcon       =   "frmPass.frx":2782
            MousePointer    =   99  'Custom
            ScaleHeight     =   570
            ScaleWidth      =   1095
            TabIndex        =   22
            TabStop         =   0   'False
            Top             =   240
            Width           =   1100
         End
      End
      Begin MSComCtl2.DTPicker DatePick 
         Height          =   315
         Left            =   4140
         TabIndex        =   2
         Top             =   1440
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
         Caption         =   "Password"
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
         Left            =   3030
         TabIndex        =   36
         Top             =   1080
         Width           =   960
      End
      Begin VB.Label Label 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H00808080&
         BackStyle       =   0  'Transparent
         Caption         =   "User Name"
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
         Left            =   2940
         TabIndex        =   35
         Top             =   720
         Width           =   1050
      End
      Begin VB.Label Label 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H00808080&
         BackStyle       =   0  'Transparent
         Caption         =   "Birth Date"
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
         Left            =   3000
         TabIndex        =   34
         Top             =   1440
         Width           =   990
      End
      Begin VB.Label Label 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H00808080&
         BackStyle       =   0  'Transparent
         Caption         =   "Type"
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
         Left            =   3540
         TabIndex        =   33
         Top             =   1800
         Width           =   450
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
Attribute VB_Name = "frmPassword"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'##############################################
'#          Coded by Walter A. Narvasa        #
'#       POS2000 - Point of Sales System      #
'#                                            #
'#           area :  frmPassword              #
'#    description :  Password Security        #
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
    Call MessageBox("frmPassword", "Are you sure you want to delete User Name " + datprimary("USER_NAME") + " ?", 1)
    frmMessageBox2.Show
    Call MacButton("  Delete", frmPassword.cmdDel, 0, 0, 73, 50, frmLogin.Source, 0, 0, 1)
End Sub

Private Sub cmdDel_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call MacButton("  Delete", frmPassword.cmdDel, 0, 0, 73, 50, frmLogin.Source, 74, 0, 1)
End Sub

Private Sub cmdDel_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call MacButton("  Delete", frmPassword.cmdDel, 0, 0, 73, 50, frmLogin.Source, 0, 0, 1)
End Sub

Private Sub cmdEdit_Click()
'    On Error GoTo ErrorEdit
    EditMode = True
    Press_Buttons ("Edit")
    txtField(1).SetFocus
    txtField(0).TabStop = False
'ErrorEdit:
'    Call MessageBox("frmPassword", Err.Description, 0)
End Sub

Private Sub cmdEdit_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call MacButton("    Edit", frmPassword.cmdEdit, 0, 0, 73, 50, frmLogin.Source, 74, 0, 1)
End Sub

Private Sub cmdEdit_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call MacButton("    Edit", frmPassword.cmdEdit, 0, 0, 73, 50, frmLogin.Source, 0, 0, 1)
End Sub

Private Sub cmdExit_Click()
    EditMode = False
    Unload Me
End Sub

Private Sub cmdExit_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call MacButton("    Exit", frmPassword.cmdExit, 0, 0, 73, 50, frmLogin.Source, 74, 0, 1)
End Sub

Private Sub cmdExit_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call MacButton("    Exit", frmPassword.cmdExit, 0, 0, 73, 50, frmLogin.Source, 0, 0, 1)
End Sub

Private Sub cmdFind_Click()
    On Error Resume Next
    EditMode = True
    Call FindBox(" Find by User Name ", _
                    "USER_PASSWORD", "USER_NAME", 0, 1, "frmLogin")
    Call MacButton("   Find", frmPassword.cmdFind, 0, 0, 73, 50, frmLogin.Source, 0, 0, 1)
End Sub

Public Sub FindUserName()
    On Error Resume Next
    strs = "select * from USER_PASSWORD where USER_NAME = '" & frmFind.txtWord & "'"
    Set dummy = frmLogin.db.OpenRecordset(strs)
    If dummy.AbsolutePosition <> -1 Then
        datprimary.MoveFirst
        Do While Not datprimary.EOF
            If datprimary("USER_NAME") = frmFind.txtWord Then
                Exit Do
            End If
                datprimary.MoveNext
        Loop
            Display_Fields
            Enable_Buttons
    Else
        Call MessageBox("frmPassword", "User Name not found", 0)
    End If
End Sub

Private Sub cmdFind_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call MacButton("   Find", frmPassword.cmdFind, 0, 0, 73, 50, frmLogin.Source, 74, 0, 1)
End Sub

Private Sub cmdFind_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call MacButton("   Find", frmPassword.cmdFind, 0, 0, 73, 50, frmLogin.Source, 0, 0, 1)
End Sub

Private Sub cmdLast_Click()
    Press_Buttons ("Last")
End Sub

Private Sub cmdLast_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call MacButton("q", frmPassword.cmdLast, 0, 0, 100, 50, frmLogin.Source, 112, 39, 3)
End Sub

Private Sub cmdLast_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call MacButton("q", frmPassword.cmdLast, 0, 0, 100, 50, frmLogin.Source, 138, 39, 3)
End Sub

Private Sub cmdNew_Click()
    On Error Resume Next
    EditMode = False
    If datprimary.RecordCount = 0 Then
        txtField(0).Text = "1"
    Else
        datprimary.MoveLast
        txtField(0).Text = Val(datprimary("USER_NAME")) + 1
    End If
    Press_Buttons ("New")
    DatePick.Value = Get_Last_Date("USER_BIRTHDATE", "USER_PASSWORD")
    txtField(0).SetFocus
End Sub

Private Sub cmdNew_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call MacButton("   New", frmPassword.cmdNew, 0, 0, 73, 50, frmLogin.Source, 74, 0, 1)
End Sub

Private Sub cmdNew_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call MacButton("   New", frmPassword.cmdNew, 0, 0, 73, 50, frmLogin.Source, 0, 0, 1)
End Sub

Private Sub cmdNext_Click()
    Press_Buttons ("Next")
End Sub

Private Sub cmdNext_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call MacButton("u", frmPassword.cmdNext, 0, 0, 100, 49, frmLogin.Source, 112, 39, 3)
End Sub

Private Sub cmdNext_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call MacButton("u", frmPassword.cmdNext, 0, 0, 100, 49, frmLogin.Source, 138, 39, 3)
End Sub

Private Sub cmdOk_Click()
    BoxContainer2.Visible = False
End Sub


Private Sub cmdPrev_Click()
    Press_Buttons ("Prev")
End Sub

Private Sub cmdPrev_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call MacButton("t", frmPassword.cmdPrev, 0, 0, 100, 50, frmLogin.Source, 112, 39, 3)
End Sub

Private Sub cmdPrev_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call MacButton("t", frmPassword.cmdPrev, 0, 0, 100, 50, frmLogin.Source, 138, 39, 3)
End Sub

Private Sub cmdSave_Click()
    On Error Resume Next
    If EditMode = True Then
        Call MacButton("    Edit", frmPassword.cmdEdit, 0, 0, 73, 50, frmLogin.Source, 0, 0, 1)
        Press_Buttons ("Save")
    Else
        Call MacButton("   New", frmPassword.cmdNew, 0, 0, 73, 50, frmLogin.Source, 0, 0, 1)
        If Get_User_Name Then
            Call MessageBox("frmPassword", "User Name already exists.", 0)
            frmMessageBox.SetFocus
            txtField(0) = ""
            Press_Buttons ("Undo")
        ElseIf txtField(0) = "" Then
            Call MessageBox("frmPassword", "User Name cannot be null", 0)
            frmMessageBox.SetFocus
            txtField(0) = ""
            Press_Buttons ("Undo")
        'ENABLE ONLY FOR SPECIFIC LENGTH OF 10 CHARACTERS NUMBERING
        'ElseIf Len(txtField(1)) > 1 And Len(txtField(1)) < 20 Then
        '    Call MessageBox("frmPassword", "User Name must be 20 in length", 0)
        '    frmMessageBox.SetFocus
        '    txtField(0) = ""
        '    Press_Buttons ("Undo")
        Else
            Press_Buttons ("Save")
        End If
    End If
    Call MacButton("   Save", frmPassword.cmdSave, 0, 0, 73, 50, frmLogin.Source, 0, 0, 1)
End Sub

Private Sub cmdSave_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call MacButton("   Save", frmPassword.cmdSave, 0, 0, 73, 50, frmLogin.Source, 74, 0, 1)
End Sub

Private Sub cmdSave_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call MacButton("   Save", frmPassword.cmdSave, 0, 0, 73, 50, frmLogin.Source, 0, 0, 1)
End Sub

Private Sub cmdTop_Click()
    Press_Buttons ("Top")
End Sub

Private Sub cmdTop_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call MacButton("p", frmPassword.cmdTop, 0, 0, 100, 50, frmLogin.Source, 112, 39, 3)
End Sub

Private Sub cmdTop_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call MacButton("p", frmPassword.cmdTop, 0, 0, 100, 50, frmLogin.Source, 138, 39, 3)
End Sub

Private Sub cmdUndo_Click()
    If EditMode = True Then
        Call MacButton("    Edit", frmPassword.cmdEdit, 0, 0, 73, 50, frmLogin.Source, 0, 0, 1)
    Else
        Call MacButton("   New", frmPassword.cmdNew, 0, 0, 73, 50, frmLogin.Source, 0, 0, 1)
    End If
    Press_Buttons ("Undo")
    Call MacButton("   Undo", frmPassword.cmdUndo, 0, 0, 73, 50, frmLogin.Source, 0, 0, 1)
End Sub

Private Sub cmdUndo_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call MacButton("   Undo", frmPassword.cmdUndo, 0, 0, 73, 50, frmLogin.Source, 74, 0, 1)
End Sub

Private Sub cmdUndo_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call MacButton("   Undo", frmPassword.cmdUndo, 0, 0, 73, 50, frmLogin.Source, 0, 0, 1)
End Sub

Private Sub Form_Activate()
    On Error Resume Next
    If EditMode = False Then
        strs = "select * from USER_PASSWORD order by USER_NAME"
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
    Call ColForm(CheckContainer, 217, 211, 213, 125)
    Call CreateMacOSTitleBar(titleBar, " Password Security ")
    Call BitBlt(frmPassword.Closed.hDC, 0, 0, 73, 50, frmLogin.Source.hDC, 0, 107, SRCCOPY)
    frmPassword.Closed.Refresh
    Call BitBlt(frmPassword.Maximized.hDC, 0, 0, 73, 50, frmLogin.Source.hDC, 0, 72, SRCCOPY)
    frmPassword.Maximized.Refresh
    Call BitBlt(frmPassword.Minimized.hDC, 0, 0, 73, 50, frmLogin.Source.hDC, 0, 124, SRCCOPY)
    frmPassword.Minimized.Refresh
    KeyPreview = True
    Call MacButton("   New", frmPassword.cmdNew, 0, 0, 73, 50, frmLogin.Source, 0, 0, 1)
    Call MacButton("    Edit", frmPassword.cmdEdit, 0, 0, 73, 50, frmLogin.Source, 0, 0, 1)
    Call MacButton("   Save", frmPassword.cmdSave, 0, 0, 73, 50, frmLogin.Source, 0, 0, 1)
    Call MacButton("   Undo", frmPassword.cmdUndo, 0, 0, 73, 50, frmLogin.Source, 0, 0, 1)
    Call MacButton("  Delete", frmPassword.cmdDel, 0, 0, 73, 50, frmLogin.Source, 0, 0, 1)
    Call MacButton("   Find", frmPassword.cmdFind, 0, 0, 73, 50, frmLogin.Source, 0, 0, 1)
    Call MacButton("    Exit", frmPassword.cmdExit, 0, 0, 73, 50, frmLogin.Source, 0, 0, 1)
    Call MacButton("p", frmPassword.cmdTop, 0, 0, 100, 50, frmLogin.Source, 138, 39, 3)
    Call MacButton("u", frmPassword.cmdNext, 0, 0, 100, 50, frmLogin.Source, 138, 39, 3)
    Call MacButton("t", frmPassword.cmdPrev, 0, 0, 100, 50, frmLogin.Source, 138, 39, 3)
    Call MacButton("q", frmPassword.cmdLast, 0, 0, 100, 50, frmLogin.Source, 138, 39, 3)
    txtCombo(0).Clear
    txtCombo(0).AddItem "Administrator"
    txtCombo(0).AddItem "End-User"
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
                    Call MacButton("   New", frmPassword.cmdNew, 0, 0, 73, 50, frmLogin.Source, 0, 0, 1)
                End If
        Case vbKeyE:
                If AltDown Then
                    Call MacButton("    Edit", frmPassword.cmdEdit, 0, 0, 73, 50, frmLogin.Source, 0, 0, 1)
                End If
        Case vbKeyS:
                If AltDown Then
                    Call MacButton("   Save", frmPassword.cmdSave, 0, 0, 73, 50, frmLogin.Source, 0, 0, 1)
                End If
        Case vbKeyU:
                If AltDown Then
                    Call MacButton("   Undo", frmPassword.cmdUndo, 0, 0, 73, 50, frmLogin.Source, 0, 0, 1)
                End If
        Case vbKeyD:
                If AltDown Then
                    Call MacButton("  Delete", frmPassword.cmdDel, 0, 0, 73, 50, frmLogin.Source, 0, 0, 1)
                End If
        Case vbKeyF:
                If AltDown Then
                    Call MacButton("   Find", frmPassword.cmdFind, 0, 0, 73, 50, frmLogin.Source, 0, 0, 1)
                End If
        Case vbKeyX:
                If AltDown Then
                    Call MacButton("    Exit", frmPassword.cmdExit, 0, 0, 73, 50, frmLogin.Source, 0, 0, 1)
                End If
        Case vbKeyLeft:
                If AltDown Then
                    Call MacButton("t", frmPassword.cmdPrev, 0, 0, 100, 50, frmLogin.Source, 138, 39, 3)
                End If
        Case vbKeyRight:
                If AltDown Then
                    Call MacButton("u", frmPassword.cmdNext, 0, 0, 100, 49, frmLogin.Source, 138, 39, 3)
                End If
        Case vbKeyUp:
                If AltDown Then
                    Call MacButton("p", frmPassword.cmdTop, 0, 0, 100, 50, frmLogin.Source, 138, 39, 3)
                End If
        Case vbKeyDown:
                If AltDown Then
                    Call MacButton("q", frmPassword.cmdLast, 0, 0, 100, 50, frmLogin.Source, 138, 39, 3)
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
                                      "Note: The following are Password Security Key Shortcuts." & vbCrLf & _
                                      Space(1) & vbCrLf & _
                                      "ALT-N=New, ALT-E=Edit, ALT-S=Save, ALT-U=Undo, ALT-D=Delete" & vbCrLf & _
                                      "ALT-F=Find, ALT-X=Exit" & vbCrLf & _
                                      Space(1) & vbCrLf & _
                                      "Left Arrow=Previous Records, Right Arrow=Next Records" & vbCrLf & _
                                      "Top Arrow=Top Record, Down Arrow=Last Record"
        Case vbKeyN:
                If AltDown Then
                    Call MacButton("   New", frmPassword.cmdNew, 0, 0, 73, 50, frmLogin.Source, 74, 0, 1)
                    cmdNew_Click
                End If
        Case vbKeyE:
                If AltDown Then
                    Call MacButton("    Edit", frmPassword.cmdEdit, 0, 0, 73, 50, frmLogin.Source, 74, 0, 1)
                    cmdEdit_Click
                End If
        Case vbKeyS:
                If AltDown Then
                    Call MacButton("   Save", frmPassword.cmdSave, 0, 0, 73, 50, frmLogin.Source, 74, 0, 1)
                    cmdSave_Click
                End If
        Case vbKeyU:
                If AltDown Then
                    Call MacButton("   Undo", frmPassword.cmdUndo, 0, 0, 73, 50, frmLogin.Source, 74, 0, 1)
                    cmdUndo_Click
                End If
        Case vbKeyD:
                If AltDown Then
                    Call MacButton("  Delete", frmPassword.cmdDel, 0, 0, 73, 50, frmLogin.Source, 74, 0, 1)
                    cmdDel_Click
                End If
        Case vbKeyF:
                If AltDown Then
                    Call MacButton("   Find", frmPassword.cmdFind, 0, 0, 73, 50, frmLogin.Source, 74, 0, 1)
                    cmdFind_Click
                End If
        Case vbKeyX:
                If AltDown Then
                    Call MacButton("    Exit", frmPassword.cmdExit, 0, 0, 73, 50, frmLogin.Source, 74, 0, 1)
                    cmdExit_Click
                End If
        Case vbKeyLeft:
                If AltDown Then
                    Call MacButton("t", frmPassword.cmdPrev, 0, 0, 100, 50, frmLogin.Source, 112, 39, 3)
                    cmdPrev_Click
                End If
        Case vbKeyRight:
                If AltDown Then
                    Call MacButton("u", frmPassword.cmdNext, 0, 0, 100, 49, frmLogin.Source, 112, 39, 3)
                    cmdNext_Click
                End If
        Case vbKeyUp:
                If AltDown Then
                    Call MacButton("p", frmPassword.cmdTop, 0, 0, 100, 50, frmLogin.Source, 112, 39, 3)
                    cmdTop_Click
                End If
        Case vbKeyDown:
                If AltDown Then
                    Call MacButton("q", frmPassword.cmdLast, 0, 0, 100, 50, frmLogin.Source, 112, 39, 3)
                    cmdLast_Click
                End If
    End Select
End Sub

Private Sub titleBar_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call DragForm(Me)
End Sub

Private Sub Closed_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call BitBlt(frmPassword.Closed.hDC, 0, 0, 73, 50, frmLogin.Source.hDC, 18, 107, SRCCOPY)
    frmPassword.Closed.Refresh
End Sub

Private Sub Closed_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call BitBlt(frmPassword.Closed.hDC, 0, 0, 73, 50, frmLogin.Source.hDC, 0, 107, SRCCOPY)
    frmPassword.Closed.Refresh
End Sub

Private Sub Maximized_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call BitBlt(frmPassword.Maximized.hDC, 0, 0, 73, 50, frmLogin.Source.hDC, 18, 72, SRCCOPY)
    frmPassword.Maximized.Refresh
End Sub

Private Sub Maximized_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call BitBlt(frmPassword.Maximized.hDC, 0, 0, 73, 50, frmLogin.Source.hDC, 0, 72, SRCCOPY)
    frmPassword.Maximized.Refresh
End Sub

Private Sub Minimized_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call BitBlt(frmPassword.Minimized.hDC, 0, 0, 73, 50, frmLogin.Source.hDC, 18, 124, SRCCOPY)
    frmPassword.Minimized.Refresh
End Sub

Private Sub Minimized_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call BitBlt(frmPassword.Minimized.hDC, 0, 0, 73, 50, frmLogin.Source.hDC, 0, 124, SRCCOPY)
    frmPassword.Minimized.Refresh
End Sub

Private Sub Display_Fields()
    On Error Resume Next
    If datprimary.AbsolutePosition <> -1 Then
        txtField(0) = IIf(IsNull(datprimary("USER_NAME")), "", datprimary("USER_NAME"))
        txtField(1) = IIf(IsNull(datprimary("USER_PASSWORD")), "", UnCode_Pass(datprimary("USER_PASSWORD")))
        DatePick.Value = IIf(IsDate(datprimary("USER_BIRTHDATE")), datprimary("USER_BIRTHDATE"), Date)
        txtCombo(0) = IIf(IsNull(datprimary("USER_TYPE")), "", IIf(datprimary("USER_TYPE") = "A", "Administrator", "End-User"))
        txtCheck(0).Value = IIf(IsNull(datprimary("USER_ALLOW_SM")), 0, IIf(datprimary("USER_ALLOW_SM") = 0, 0, 1))
        txtCheck(1).Value = IIf(IsNull(datprimary("USER_ALLOW_PM")), 0, IIf(datprimary("USER_ALLOW_PM") = 0, 0, 1))
        txtCheck(2).Value = IIf(IsNull(datprimary("USER_ALLOW_CM")), 0, IIf(datprimary("USER_ALLOW_CM") = 0, 0, 1))
        txtCheck(3).Value = IIf(IsNull(datprimary("USER_ALLOW_ST")), 0, IIf(datprimary("USER_ALLOW_ST") = 0, 0, 1))
        txtCheck(4).Value = IIf(IsNull(datprimary("USER_ALLOW_RT")), 0, IIf(datprimary("USER_ALLOW_RT") = 0, 0, 1))
        txtCheck(5).Value = IIf(IsNull(datprimary("USER_ALLOW_SRR")), 0, IIf(datprimary("USER_ALLOW_SRR") = 0, 0, 1))
        txtCheck(6).Value = IIf(IsNull(datprimary("USER_ALLOW_SHR")), 0, IIf(datprimary("USER_ALLOW_SHR") = 0, 0, 1))
        txtCheck(7).Value = IIf(IsNull(datprimary("USER_ALLOW_RHR")), 0, IIf(datprimary("USER_ALLOW_RHR") = 0, 0, 1))
        txtCheck(8).Value = IIf(IsNull(datprimary("USER_ALLOW_SPSR")), 0, IIf(datprimary("USER_ALLOW_SPSR") = 0, 0, 1))
        txtCheck(9).Value = IIf(IsNull(datprimary("USER_ALLOW_PLR")), 0, IIf(datprimary("USER_ALLOW_PLR") = 0, 0, 1))
        txtCheck(10).Value = IIf(IsNull(datprimary("USER_ALLOW_SLR")), 0, IIf(datprimary("USER_ALLOW_SLR") = 0, 0, 1))
        txtCheck(11).Value = IIf(IsNull(datprimary("USER_ALLOW_BRF")), 0, IIf(datprimary("USER_ALLOW_BRF") = 0, 0, 1))
        txtCheck(12).Value = IIf(IsNull(datprimary("USER_ALLOW_PS")), 0, IIf(datprimary("USER_ALLOW_PS") = 0, 0, 1))
        txtCheck(13).Value = IIf(IsNull(datprimary("USER_ALLOW_CFS")), 0, IIf(datprimary("USER_ALLOW_CFS") = 0, 0, 1))
        txtCheck(14).Value = IIf(IsNull(datprimary("USER_ALLOW_SS")), 0, IIf(datprimary("USER_ALLOW_SS") = 0, 0, 1))
    Else
        Clear_Fields
    End If
End Sub

Private Sub Clear_Fields()
    ' ENABLE ONLY TO BLANK FIELD txtField(1)
    'txtField(0) = ""
    txtField(1) = ""
    For i = 0 To 14
        txtCheck(i) = 0
    Next i
End Sub

Private Sub Update_Fields(isNew As Boolean)
    On Error Resume Next
    If isNew Then
        datprimary.AddNew
    Else
        datprimary.Edit
    End If
        datprimary("USER_NAME") = txtField(0)
        datprimary("USER_PASSWORD") = Decode_Pass(txtField(1))
        datprimary("USER_BIRTHDATE") = DatePick.Value
        datprimary("USER_TYPE") = IIf(txtCombo(0) = "Administrator", "A", "E")
        datprimary("USER_ALLOW_SM") = IIf(txtCheck(0).Value = 0, 0, -1)
        datprimary("USER_ALLOW_PM") = IIf(txtCheck(1).Value = 0, 0, -1)
        datprimary("USER_ALLOW_CM") = IIf(txtCheck(2).Value = 0, 0, -1)
        datprimary("USER_ALLOW_ST") = IIf(txtCheck(3).Value = 0, 0, -1)
        datprimary("USER_ALLOW_RT") = IIf(txtCheck(4).Value = 0, 0, -1)
        datprimary("USER_ALLOW_SRR") = IIf(txtCheck(5).Value = 0, 0, -1)
        datprimary("USER_ALLOW_SHR") = IIf(txtCheck(6).Value = 0, 0, -1)
        datprimary("USER_ALLOW_RHR") = IIf(txtCheck(7).Value = 0, 0, -1)
        datprimary("USER_ALLOW_SPSR") = IIf(txtCheck(8).Value = 0, 0, -1)
        datprimary("USER_ALLOW_PLR") = IIf(txtCheck(9).Value = 0, 0, -1)
        datprimary("USER_ALLOW_SLR") = IIf(txtCheck(10).Value = 0, 0, -1)
        datprimary("USER_ALLOW_BRF") = IIf(txtCheck(11).Value = 0, 0, -1)
        datprimary("USER_ALLOW_PS") = IIf(txtCheck(12).Value = 0, 0, -1)
        datprimary("USER_ALLOW_CFS") = IIf(txtCheck(13).Value = 0, 0, -1)
        datprimary("USER_ALLOW_SS") = IIf(txtCheck(14).Value = 0, 0, -1)
        datprimary.Update
    If isNew Then datprimary.MoveLast
End Sub

Private Sub Enable_Fields(isLock As Boolean)
    On Error Resume Next
    For i = 0 To 1
        txtField(i).Enabled = Not isLock
    Next i
    For i = 0 To 14
        txtCheck(i).Enabled = Not isLock
    Next i
        txtCombo(0).Enabled = Not isLock
        DatePick.Enabled = Not isLock
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
    DatePick.TabStop = isTab
    txtCombo(0).TabStop = isTab
    For i = 0 To 14
        txtCheck(i).TabStop = isTab
    Next i
End Sub

Function Get_User_Name() As Boolean
    On Error Resume Next
    strs = "select * from USER_PASSWORD where USER_NAME= '" & txtField(0) & "'"
    Set dummy = frmLogin.db.OpenRecordset(strs)
    If dummy.AbsolutePosition <> -1 Then
        Get_User_Name = True
    Else
        Get_User_Name = False
    End If
End Function
