VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#4.6#0"; "CRYSTL32.OCX"
Begin VB.Form frmPrint 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   0  'None
   Caption         =   "frmPrint"
   ClientHeight    =   2985
   ClientLeft      =   0
   ClientTop       =   -30
   ClientWidth     =   5565
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2985
   ScaleWidth      =   5565
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox BoxContainer 
      BackColor       =   &H00808080&
      BorderStyle     =   0  'None
      FillColor       =   &H00C0C0C0&
      Height          =   2890
      Left            =   40
      ScaleHeight     =   2895
      ScaleWidth      =   5460
      TabIndex        =   0
      Top             =   40
      Width           =   5460
      Begin VB.PictureBox cmdClose 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   570
         Left            =   2880
         MouseIcon       =   "frmPrint.frx":0000
         MousePointer    =   99  'Custom
         ScaleHeight     =   570
         ScaleWidth      =   1095
         TabIndex        =   8
         Top             =   2040
         Width           =   1100
      End
      Begin VB.OptionButton RadioItem 
         BackColor       =   &H00C0C0C0&
         Caption         =   "From"
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
         Left            =   840
         TabIndex        =   6
         Top             =   1125
         Width           =   200
      End
      Begin VB.OptionButton RadioItem 
         BackColor       =   &H00C0C0C0&
         Caption         =   "All Records"
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
         Left            =   840
         TabIndex        =   5
         Top             =   720
         Width           =   200
      End
      Begin VB.OptionButton RadioItem 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Current Date"
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
         Left            =   840
         TabIndex        =   4
         Top             =   1560
         Width           =   200
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
      Begin VB.PictureBox cmdPrint 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   570
         Left            =   1560
         MouseIcon       =   "frmPrint.frx":030A
         MousePointer    =   99  'Custom
         ScaleHeight     =   570
         ScaleWidth      =   1095
         TabIndex        =   1
         Top             =   2040
         Width           =   1100
      End
      Begin Crystal.CrystalReport PrintIt 
         Left            =   720
         Top             =   2040
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   262150
         WindowLeft      =   230
         WindowTop       =   45
         WindowWidth     =   552
         WindowHeight    =   494
         WindowBorderStyle=   0
         WindowControlBox=   0   'False
         WindowMaxButton =   0   'False
         WindowMinButton =   0   'False
      End
      Begin MSComCtl2.DTPicker FromDate 
         Height          =   315
         Left            =   1800
         TabIndex        =   3
         Top             =   1080
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   556
         _Version        =   393216
         Enabled         =   0   'False
         CalendarTitleBackColor=   8421504
         CustomFormat    =   "MM/dd/yyyy"
         Format          =   24772611
         CurrentDate     =   36516
      End
      Begin MSComCtl2.DTPicker ToDate 
         Height          =   315
         Left            =   3480
         TabIndex        =   7
         Top             =   1080
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   556
         _Version        =   393216
         Enabled         =   0   'False
         CalendarTitleBackColor=   8421504
         CustomFormat    =   "MM/dd/yyyy"
         Format          =   24772611
         CurrentDate     =   36516
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "To"
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
         Left            =   3120
         TabIndex        =   12
         Top             =   1125
         Width           =   210
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Current Date"
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
         Left            =   1200
         TabIndex        =   11
         Top             =   1560
         Width           =   1260
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "From"
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
         Left            =   1200
         TabIndex        =   10
         Top             =   1125
         Width           =   495
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "All Records"
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
         Left            =   1200
         TabIndex        =   9
         Top             =   720
         Width           =   1080
      End
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   2
      Height          =   2950
      Left            =   0
      Top             =   0
      Width           =   5520
   End
End
Attribute VB_Name = "frmPrint"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim choice As Integer

Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub cmdClose_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call MacButton("  Close", frmPrint.cmdClose, 0, 0, 73, 50, frmLogin.Source, 74, 0, 1)
End Sub

Private Sub cmdClose_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call MacButton("  Close", frmPrint.cmdClose, 0, 0, 73, 50, frmLogin.Source, 0, 0, 1)
End Sub

Private Sub cmdPrint_Click()
    Select Case frmMain.What_Rpt
            Case "Selling History Report"
                    Select Case choice
                        Case 0
                            PrintIt.SelectionFormula = ""
                        Case 1
                            PrintIt.SelectionFormula = "{QrySellingRpt.DATENUM} >= " & CDbl(Format(FromDate.Value, "YYYYMMDD")) & " And {QrySellingRpt.dATENUM} <= " & CDbl(Format(ToDate.Value, "YYYYMMDD"))
                        Case 2
                            PrintIt.SelectionFormula = "{QrySellingRpt.SoldDate} = '" & CStr(Date) & "'"
                    End Select
                    PrintIt.ReportFileName = App.Path & "\REPORTS\SELLING.RPT"
                    PrintIt.Action = 1
            Case "Receiving History Report"
                    Select Case choice
                        Case 0
                            PrintIt.SelectionFormula = ""
                        Case 1
                            PrintIt.SelectionFormula = "{QryReceivingRpt.DATENUM} >= " & CDbl(Format(FromDate.Value, "YYYYMMDD")) & " And {QryReceivingRpt.dATENUM} <= " & CDbl(Format(ToDate.Value, "YYYYMMDD"))
                        Case 2
                            PrintIt.SelectionFormula = "{QryReceivingRpt.DATENUM} = " & CDbl(Format(Date, "YYYYMMDD"))
                    End Select
                    PrintIt.ReportFileName = App.Path & "\REPORTS\PURCHASE.RPT"
                    PrintIt.Action = 1
            Case "Stocks Per Supplier Report"
                    Select Case choice
                        Case 0
                            PrintIt.SelectionFormula = ""
                        Case 1
                            PrintIt.SelectionFormula = "{QryProdBySupp.DATENUM} >= " & CDbl(Format(FromDate.Value, "YYYYMMDD")) & " And {QryProdBySupp.dATENUM} <= " & CDbl(Format(ToDate.Value, "YYYYMMDD"))
                        Case 2
                            PrintIt.SelectionFormula = "{QryProdBySupp.DATENUM} = " & CDbl(Format(Date, "YYYYMMDD"))
                    End Select
                    PrintIt.ReportFileName = App.Path & "\REPORTS\PRODSUPP.RPT"
                    PrintIt.Action = 1
    End Select
    Me.Hide
End Sub

Private Sub cmdPrint_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call MacButton("   Print", frmPrint.cmdPrint, 0, 0, 73, 50, frmLogin.Source, 74, 0, 1)
End Sub

Private Sub cmdPrint_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call MacButton("   Print", frmPrint.cmdPrint, 0, 0, 73, 50, frmLogin.Source, 0, 0, 1)
End Sub

Private Sub Form_Activate()
    FromDate.Value = Date - (Day(Date) - 1)
    ToDate.Value = Date
End Sub

Private Sub Form_Load()
    Call CreateMacOSTitleBar(titleBar, " Report Parameter ")
    Call MacButton("   Print", frmPrint.cmdPrint, 0, 0, 73, 50, frmLogin.Source, 0, 0, 1)
    Call MacButton("  Close", frmPrint.cmdClose, 0, 0, 73, 50, frmLogin.Source, 0, 0, 1)
    Call ColForm(BoxContainer, 217, 211, 213, 125)
End Sub

Private Sub RadioItem_Click(Index As Integer)
    choice = Index
    If Index = 1 Then
        FromDate.Enabled = True
        ToDate.Enabled = True
    Else
        FromDate.Enabled = False
        ToDate.Enabled = False
    End If
End Sub

Private Sub titleBar_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call DragForm(Me)
End Sub
