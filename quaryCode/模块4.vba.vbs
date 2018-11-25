Sub 导出用友表()
    Dim cn_pz As ADODB.Connection
    Dim rs_pz As ADODB.Recordset
    Dim sql_pz As String

    Dim str_yer As String
    Dim str_new As String
    Dim str_acc As String
    Dim str_biao As String
    Dim str_accno As String

    str_acc = InputBox(prompt:="请输入帐套编号")
    str_year = InputBox(prompt:="请输入会计年度")
    int_biao = InputBox("请选择要打开的表" & Chr(13) & Chr(10) & "1、科表表" & Chr(13) & Chr(10) & "2、凭证表" & Chr(13) & Chr(10) & "3、流量表" & Chr(13) & Chr(10) & "4、固定资产卡片" & Chr(13) & Chr(10), , 1)
    str_accno = "UFDATA_" & str_acc & "_" & str_year

    Application.DisplayAlerts = False
    Select Case int_biao
    Case 1
        str_biao = "code"
        str_new = str_year & "科表表"
    Case 2
        str_biao = "GL_accvouch"
        str_new = str_year & "凭证表"

    End Select
    sql_pz = "SELECT  * FROM " & str_biao
    On Error Resume Next
    If Worksheets(str_new) Is Nothing Then
        Worksheets.Add(after:=Worksheets(Worksheets.Count)).Name = str_new
    End If
    m = 2

    Set rs_pz = New ADODB.Recordset
    rs_pz.Open sql_pz, obj_conn(str_accno), adOpenStatic, adLockBatchOptimistic
    For int_II = 0 To rs_pz.Fields.Count - 1
        Worksheets(str_new).Cells(1, int_II + 1).Value = rs_pz(int_II).Name
    Next

    While Not rs_pz.EOF
        For int_II = 0 To rs_pz.Fields.Count - 1
            Worksheets(str_new).Cells(m, int_II + 1).Value = rs_pz(int_II)
        Next
        m = m + 1
        rs_pz.MoveNext
    Wend

    rs_pz.Close   '完成后要关闭
    obj_conn(str_accno).Close    '完成后要关闭

    Application.DisplayAlerts = True
End Sub
Sub 导出总帐及明细科目表()
    Dim cn_pz As ADODB.Connection
    Dim rs_pz As ADODB.Recordset
    Dim sql_pz As String

    Dim str_yer As String
    Dim str_new As String
    Dim str_acc As String
    Dim str_biao As String
    Dim str_accno As String

    str_acc = InputBox(prompt:="请输入帐套编号")
    str_year = InputBox(prompt:="请输入会计年度")
    int_biao = InputBox("请选择要打开的表" & Chr(13) & Chr(10) & "1、总帐科目" & Chr(13) & Chr(10) & "2、明细科目表" & Chr(13), , 1)
    str_accno = "UFDATA_" & str_acc & "_" & str_year

    Application.DisplayAlerts = False
    Select Case int_biao
    Case 1
        str_biao = "code"
        str_new = str_year & "总帐科目"
        sql_pz = "select  ccode,ccode_name from code   " & _
                 "where ccode  in ( select distinct( ccode) from gl_accsum) " & _
               "             and igrade=1"
    Case 2
        str_biao = "GL_accvouch"
        str_new = str_year & "明细科目表"
        sql_pz = "select  ccode,ccode_name from code   " & _
                 "where ccode  in ( select distinct( ccode) from gl_accsum) " & _
               "     and bend=1"
    Case Else
        str_biao = int_biao
        str_new = str_year & int_biao

    End Select

    On Error Resume Next
    If Worksheets(str_new) Is Nothing Then
        Worksheets.Add(after:=Worksheets(Worksheets.Count)).Name = str_new
    End If
    m = 2

    Set rs_pz = New ADODB.Recordset
    rs_pz.Open sql_pz, obj_conn(str_accno), adOpenStatic, adLockBatchOptimistic
    For int_II = 0 To rs_pz.Fields.Count - 1
        Worksheets(str_new).Cells(1, int_II + 1).Value = rs_pz(int_II).Name
    Next

    While Not rs_pz.EOF
        For int_II = 0 To rs_pz.Fields.Count - 1
            Worksheets(str_new).Cells(m, int_II + 1).Value = rs_pz(int_II)
        Next
        m = m + 1
        rs_pz.MoveNext
    Wend

    rs_pz.Close   '完成后要关闭
    obj_conn(str_accno).Close    '完成后要关闭

    Application.DisplayAlerts = True
End Sub

