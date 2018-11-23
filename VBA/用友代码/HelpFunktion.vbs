Sub 帐务查询()

    Dim str_yer As String
    Dim str_new As String
    Dim str_acc As String
    Dim int_star As Integer
    Dim int_end As Integer

    str_acc = InputBox(prompt:="请输入帐套编号")
    int_star = InputBox(prompt:="请输入开始会计年度")
    int_end = InputBox(prompt:="请输入结束会计年度")

    For int_i = int_star To int_end
        str_yer = int_i
        str_accno = "UFDATA_" & str_acc & "_" & str_yer
        Call 用友明细帐(str_accno)
        Call 余额表(str_accno)
        Call 用友总帐(str_accno)
    Next
End Sub

'------------------------------------------------------------------------------'

Sub 用友明细帐(zt As String)
'@iAdjustPZ=1表示包括调整期凭证,=0表示不包括调整期凭证 xiaogj
'@bVouch=1表示只包括已经记帐凭证,=0表示包含末记帐凭证
    Dim rst As ADODB.Recordset
    Dim csqlstr As String
    csqlstr = "exec GL_SubsLedger " & _
              "N'" & 1001 & "'," & _
              Worksheets("余额查询").Cells(3, "b").Value & "," & _
              Worksheets("余额查询").Cells(3, "c").Value & "," & _
              "1"
    Set conn = New ADODB.Connection

    Set rst = New ADODB.Recordset
    With rst
        .ActiveConnection = obj_conn(zt)
        .Open csqlstr
    End With
    With Worksheets.Add(Worksheets("凭证"))
        If Worksheets("余额查询").Cells(6, "b").Value = 0 Then
            .Name = Right(zt, 4) & "年总帐明细表"
        Else
            .Name = Right(zt, 4) & "年末级科目明细表"
        End If

        .Range("a2").CopyFromRecordset rst
        ' .Range("b:C,g:m,o:P,s:AC").Delete Shift:=xlToLeft
        .Columns("a:M").Font.Size = 10

    End With

    ActiveWindow.DisplayZeros = False

    rst.Close

End Sub


'------------------------------------------------------------------------------'

Sub 用友总帐(zt As String)
'@iAdjustPZ=1表示包括调整期凭证,=0表示不包括调整期凭证 xiaogj
'@bVouch=1表示只包括已经记帐凭证,=0表示包含末记帐凭证
    Dim rst As ADODB.Recordset
    Dim csqlstr As String
    csqlstr = "exec GL_glCode " & _
              "N'" & 1001 & "'," & _
              Worksheets("余额查询").Cells(3, "b").Value & "," & _
              Worksheets("余额查询").Cells(3, "c").Value & "," & _
              "0,1"
    Set conn = New ADODB.Connection

    Set rst = New ADODB.Recordset
    With rst
        .ActiveConnection = obj_conn(zt)
        .Open csqlstr
    End With
    With Worksheets.Add(Worksheets("凭证"))
        If Worksheets("余额查询").Cells(6, "b").Value = 0 Then
            .Name = Right(zt, 4) & "年总帐表"
        Else
            .Name = Right(zt, 4) & "年总帐"
        End If

        .Range("a2").CopyFromRecordset rst
        ' .Range("b:C,g:m,o:P,s:AC").Delete Shift:=xlToLeft
        .Columns("a:M").Font.Size = 10

    End With

    ActiveWindow.DisplayZeros = False

    rst.Close

End Sub
