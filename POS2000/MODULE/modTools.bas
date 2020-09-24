Attribute VB_Name = "modTools"
'##############################################
'#          Coded by Walter A. Narvasa        #
'#       POS2000 - Point of Sales System      #
'#                                            #
'#           area :  modTools                 #
'#    description :  Database Tools           #
'#         e-mail :  walter@wancom.8k.com     #
'#            url :  http://wancom.8k.com     #
'#                                            #
'##############################################

Global Today As Variant
Global filename As String
Global CmdType As String
Global FindType As String
Global TableType As String
Global FieldType As String
Global ListType As Integer
Global DescType As Integer
Global FormType As String
Global EditMode As Boolean
Global blnAuto As Boolean
Global db As Database
Global rst As Recordset
Dim dummy As Recordset

' PAD STRING
Function Pad_Str(str As String, val_to_pad As String, strlength As Integer, Right As Boolean) As String
    Dim s1 As String
    s1 = ""
    For i = 1 To strlength - Len(str) Step 1
        s1 = s1 & val_to_pad
    Next i
    If Right Then
        Pad_Str = str & s1
    Else
        Pad_Str = s1 & str
    End If
End Function

' SEARCH VALIDATION
Function Search_Exist(str As String, fieldname As String, table As String) As Boolean
    strs = "select * from " & table & " where " & fieldname & " = '" & str & "'"
    Set dummy = frmLogin.db.OpenRecordset(strs)
    If dummy.AbsolutePosition <> -1 Then
         Search_Exist = True
    Else
        Search_Exist = False
    End If
End Function

' CONVERT STRING TO NUMERIC
Function Convert_Numeric(p As String, IsMoney As Boolean) As String
    If p <> "" Then
        If IsNumeric(p) Then
            If IsMoney Then
                Convert_Numeric = Format(p, "#,###,###,##0.00")
            Else
                Convert_Numeric = Format(p, "##,###,###,##0")
            End If
        Else
            MsgBox "Invalid numeric entered"
        End If
    Else
        Convert_Numeric = "0"
    End If
End Function

' UPDATE STOCKS
Public Sub Update_Stocks(code As String, IsAdd As Boolean, qty As Integer)
    On Error Resume Next
    strs = "select QUAN from PROD_STOCKS where PRODCODE = '" & code & "'"
    Set dummy = frmLogin.db.OpenRecordset(strs)
    If dummy.AbsolutePosition <> -1 Then
        dummy.Edit
        If IsAdd Then
            dummy("QUAN") = CStr(CInt(dummy("QUAN")) + qty)
        Else
            dummy("QUAN") = CStr(CInt(dummy("QUAN")) - qty)
        End If
        dummy.Update
    End If
End Sub

' UPDATE STOCKS 2 FOR RECEIVING TRANSACTION
Public Sub Receive_Stocks(code As String, qty As Integer)
    On Error Resume Next
    strs = "select QUAN,PDATE from PROD_STOCKS where PRODCODE = '" & code & "'"
    Set dummy = frmLogin.db.OpenRecordset(strs)
    If dummy.AbsolutePosition <> -1 Then
        dummy.Edit
        dummy("QUAN") = CStr(CInt(dummy("QUAN")) + qty)
        dummy("PDATE") = frmReceiving.DTPick.Value
        dummy.Update
    End If
End Sub

' GET NEW TOTAL
Function Get_New_Total(code As String) As Double
    strs = "select quan from prod_stocks where prodcode = '" & code & "'"
    Set dummy = frmLogin.db.OpenRecordset(strs)
    If dummy.AbsolutePosition <> -1 Then
        'Get_New_Total = Convert_Numeric(CDbl(txtField(5)) * CDbl(txtField(6)), True)
    Else
      Get_New_Total = 0
    End If
End Function

' DECODE PASSWORD
Function Decode_Pass(p_str As String) As String
    For i = 1 To Len(p_str) Step 1
        strs = strs + Chr(Asc(Mid(p_str, i, 1)) * 2)
    Next i
        Decode_Pass = strs
End Function

' UNCODE PASSWORD
Function UnCode_Pass(p_str As String) As String
    For i = 1 To Len(p_str) Step 1
        strs = strs + Chr(Asc(Mid(p_str, i, 1)) / 2)
    Next i
        UnCode_Pass = strs
End Function

' GET LAST DATE
Function Get_Last_Date(fieldname As String, table As String) As Date
    strs = "select max(" & fieldname & ") from " & table
    Set dummy = frmLogin.db.OpenRecordset(strs)
    If dummy.AbsolutePosition <> -1 Then
        Get_Last_Date = IIf(IsNull(dummy(0)), Date, dummy(0))
    Else
        Get_Last_Date = Date
    End If
End Function

' MESSAGE BOX
Function MessageBox(xType As String, xMessage, xMode As Integer)
    CmdType = xType
    If xMode = 0 Then
        frmMessageBox.MessageList = xMessage
        frmMessageBox.Show
    ElseIf xMode = 1 Then
        frmMessageBox2.MessageList = xMessage
        frmMessageBox2.Show
    End If
End Function

' FIND BOX
Function FindBox(xCategory As String, xTableType As String, _
                 xFieldType As String, xListType As Integer, _
                 xDescType As Integer, xFormType As String)
    FindType = xCategory
    TableType = xTableType
    FieldType = xFieldType
    ListType = xListType
    DescType = xDescType
    FormType = xFormType
    frmFind.Show
End Function

' SQL QUERY DATA FOR FIND BOX USE ONLY
Function QueryData(reqText As String, xTable As String, _
                   xField As String, xList As Integer, _
                   xDesc As Integer, xForm)
    On Error Resume Next
    Dim SQLtext As String, SQLText2 As String
        SQLtext = "SELECT * FROM " + xTable + " WHERE Left(" + xField + "," & _
                    Len(reqText) & ")='" & reqText & "';"
        Set rst = frmLogin.db.OpenRecordset(SQLtext)
        frmFind.lstWords.Clear
        If rst.RecordCount = 0 Then
            frmFind.lstWords.AddItem "No item was found"
            frmFind.txtExp.Text = ""
            Exit Function
        End If
        rst.MoveLast: rst.MoveFirst
        Do Until rst.EOF
            frmFind.lstWords.AddItem rst.Fields(xList)
            rst.MoveNext
        Loop
        If frmFind.lstWords.ListCount = 1 Then
            rst.MoveFirst
            frmFind.txtExp.Text = rst.Fields(xDesc)
            frmFind.txtWord.Text = frmFind.lstWords.List(xList)
            frmFind.txtWord.SelLength = Len(frmFind.txtWord.Text)
        Else
            frmFind.txtExp.Text = ""
        End If
End Function

' SQL QUERY DATA FOR SELLING TRANSACTION USE ONLY
Function QueryDetail(reqText As String)
    On Error Resume Next
    Dim SQLtext As String, SQLText2 As String
        SQLtext = "SELECT * FROM PROD_STOCKS WHERE Left(PRODDES," & _
                    Len(reqText) & ")='" & reqText & "';"
        Set rst = frmLogin.db.OpenRecordset(SQLtext)
        frmSelling.lstWords.Clear
        If rst.RecordCount = 0 Then
            frmSelling.lstWords.AddItem "No item was found"
            frmSelling.txtProductCode.Text = ""
            frmSelling.txtUnitCost.Text = ""
            frmSelling.txtAvailableStock = ""
            frmSelling.txtTotalCost.Text = ""
            Exit Function
        End If
        rst.MoveLast: rst.MoveFirst
        Do Until rst.EOF
            frmSelling.lstWords.AddItem rst.Fields(0)
            rst.MoveNext
        Loop
        If frmSelling.lstWords.ListCount = 1 Then
            rst.MoveFirst
            frmSelling.txtProductCode.Text = IIf(IsNull(rst.Fields("PRODCODE")), "", rst.Fields("PRODCODE"))
            frmSelling.txtUnitCost.Text = Convert_Numeric(IIf(IsNull(rst.Fields("UNIT_COST")), 0, rst.Fields("UNIT_COST")), True)
            frmSelling.txtAvailableStock = Convert_Numeric(IIf(IsNull(rst.Fields("QUAN")), 0, rst.Fields("QUAN")), False)
            frmSelling.txtTotalCost.Text = Convert_Numeric(CDbl(IIf(IsNull(rst.Fields("UNIT_COST")), 0, rst.Fields("UNIT_COST"))) * CDbl(IIf(IsNull(rst.Fields("QUAN")), 0, rst.Fields("QUAN"))), True)
            frmSelling.TerminalDescription = " " & frmSelling.txtWord & "/Available Stock=" & Convert_Numeric(IIf(IsNull(rst.Fields("QUAN")), 0, rst.Fields("QUAN")), False)
            frmSelling.TerminalCounter = Convert_Numeric(IIf(IsNull(rst.Fields("UNIT_COST")), 0, rst.Fields("UNIT_COST")), True)
            frmSelling.txtWord.Text = frmSelling.lstWords.List(0)
            frmSelling.txtWord.SelLength = Len(frmSelling.txtWord.Text)
            frmSelling.lstWords.Visible = False
        Else
            frmSelling.txtProductCode.Text = ""
            frmSelling.txtUnitCost.Text = ""
            frmSelling.txtAvailableStock = ""
            frmSelling.txtTotalCost.Text = ""
        End If
End Function

' SQL QUERY DATA FOR RECEIVING TRANSACTION USE ONLY
Function QueryReceiving(reqText As String)
    On Error Resume Next
    Dim SQLtext As String, SQLText2 As String
        SQLtext = "SELECT * FROM PROD_STOCKS WHERE Left(PRODDES," & _
                    Len(reqText) & ")='" & reqText & "';"
        Set rst = frmLogin.db.OpenRecordset(SQLtext)
        frmReceiving.lstWords.Clear
        If rst.RecordCount = 0 Then
            frmReceiving.lstWords.AddItem "No item was found"
            frmReceiving.txtProductCode.Text = ""
            frmReceiving.txtUnitCost.Text = ""
            frmReceiving.txtAvailableStock = ""
            frmReceiving.txtTotalCost.Text = ""
            frmReceiving.txtDate = ""
            Exit Function
        End If
        rst.MoveLast: rst.MoveFirst
        Do Until rst.EOF
            frmReceiving.lstWords.AddItem rst.Fields(0)
            rst.MoveNext
        Loop
        If frmReceiving.lstWords.ListCount = 1 Then
            rst.MoveFirst
            frmReceiving.txtProductCode.Text = IIf(IsNull(rst.Fields("PRODCODE")), "", rst.Fields("PRODCODE"))
            frmReceiving.txtUnitCost.Text = Convert_Numeric(IIf(IsNull(rst.Fields("UNIT_COST")), 0, rst.Fields("UNIT_COST")), True)
            frmReceiving.txtAvailableStock = Convert_Numeric(IIf(IsNull(rst.Fields("QUAN")), 0, rst.Fields("QUAN")), False)
            frmReceiving.txtTotalCost.Text = Convert_Numeric(CDbl(IIf(IsNull(rst.Fields("UNIT_COST")), 0, rst.Fields("UNIT_COST"))) * CDbl(IIf(IsNull(rst.Fields("QUAN")), 0, rst.Fields("QUAN"))), True)
            frmReceiving.txtDate = IIf(IsDate(rst.Fields("PDATE")), Format(rst.Fields("PDATE"), "MM" & "/" & "DD" & "/" & "YYYY "), Format(Date, "MM" & "/" & "DD" & "/" & "YYYY "))
            frmReceiving.txtWord.Text = frmReceiving.lstWords.List(0)
            frmReceiving.txtWord.SelLength = Len(frmReceiving.txtWord.Text)
            frmReceiving.lstWords.Visible = False
        Else
            frmReceiving.txtProductCode.Text = ""
            frmReceiving.txtUnitCost.Text = ""
            frmReceiving.txtAvailableStock = ""
            frmReceiving.txtTotalCost.Text = ""
            frmReceiving.txtDate = ""
        End If
End Function

' COPY PICTURE FILE TO A FIELDNAME
Function CopyFileToField(filename As String, fd As DAO.Field)
    Dim ChunkSize As Long
    Dim FileNum As Integer
    Dim Buffer()  As Byte
    Dim BytesNeeded As Long
    Dim Buffers As Long
    Dim Remainder As Long
    Dim i As Long
    If Len(filename) = 0 Then
        Exit Function
    End If
    If Dir(filename) = "" Then
        Err.Raise vbObjectError, , "File not found: """ & filename & """"
    End If
    ChunkSize = 65536
    FileNum = FreeFile
    Open filename For Binary As #FileNum
    BytesNeeded = LOF(FileNum)
    Buffers = BytesNeeded \ ChunkSize
    Remainder = BytesNeeded Mod ChunkSize
    For i = 0 To Buffers - 1
        ReDim Buffer(ChunkSize)
        Get #FileNum, , Buffer
        fd.AppendChunk Buffer
    Next
    ReDim Buffer(Remainder)
    Get #FileNum, , Buffer
    fd.AppendChunk Buffer
    Close #FileNum
End Function
