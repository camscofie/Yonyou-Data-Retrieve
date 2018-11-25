Option Explicit
Private user As New user
Private txtAcc As String
Private txtAccYear As String
Private sql As New sqlQuary
Private command As String
Private dbRecord As New ADODB.Recordset



Private Sub ComboBox3_Change()
    command = sql.sqlDict.Item(ComboBox3.value)
End Sub

Private Sub dbScan_Click()
    Call dbConnection.dbScan
    Dim count As Integer
    For count = LBound(Data.colAccid) To UBound(Data.colAccid)
        If colAccid(count).cAccid <> "" Then
            ComboBox1.AddItem Data.colAccid(count).cAccid
        End If
    Next count
End Sub

Private Sub UserForm_Initialize()
    txtUserName.value = "sa"
    txtServer.SetFocus
    sql.initSQL

    Dim key As Variant
    For Each key In sql.sqlDict.Keys
        ComboBox3.AddItem key
    Next key
End Sub


Private Sub ComboBox1_Change()
    ComboBox2.Clear
    Dim count As Integer
    For count = LBound(Data.colAccid) To UBound(Data.colAccid)
        If colAccid(count).cAccid = ComboBox1.value Then
            Dim year As Variant
            For Each year In colAccid(count).colYear
                ComboBox2.AddItem year
            Next year
        End If
    Next count
    txtAcc = ComboBox1.value
End Sub


Private Sub ComboBox2_Change()
    txtAccYear = ComboBox2.value
End Sub



Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdComfirm_Click()
    Set user = New user
    user.UserName = txtUserName.value
    Call user.SetDatabase(txtAcc, txtAccYear)

    dbRecord.Open command, dbConnection.initDB(user, command)
    ThisWorkbook.Sheets("Sheet2").Range("A1").CopyFromRecordset dbRecord

    MsgBox "信息检索成功"
    Unload Me

End Sub

Private Sub txtUserName_Change()
    user.UserName = txtUserName.value
End Sub


Private Sub txtServer_Change()
    user.UserServer = txtUserName.value
End Sub

Private Sub txtUserPass_Change()
    user.UserPass = txtUserPass.value
End Sub
