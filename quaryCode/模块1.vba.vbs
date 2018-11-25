Option Explicit
Public str_accno As String
Public str_server As String
Public str_accno_star As String


Public Function obj_conn(str_dbs As String) As ADODB.Connection
    Set obj_conn = New ADODB.Connection
    With obj_conn
        .ConnectionString = "driver={SQL Server};server=" & str_server & ";uid=sa;pwd=;database=" & str_dbs
        .Open
    End With
End Function
Public Function 科目名称(str_kmbm As String, bln_zz As Boolean, lng_weishu As Integer) As String
    Dim sql_kmmz As String
    Dim rs_kmmz As ADODB.Recordset
    If bln_zz = True Then
        str_kmbm = Left(str_kmbm, lng_weishu)
    End If
    sql_kmmz = "SELECT code.ccode, code.ccode_name"
    sql_kmmz = sql_kmmz & " FROM code"
    sql_kmmz = sql_kmmz & " WHERE code.ccode='" & str_kmbm & "'"
    Set rs_kmmz = New ADODB.Recordset
    rs_kmmz.Open sql_kmmz, obj_conn(str_accno), adOpenStatic, adLockBatchOptimistic
    科目名称 = rs_kmmz!ccode_name.Value
    rs_kmmz.Close   '完成后要关闭
    obj_conn(str_accno).Close    '完成后要关闭
End Function
Public Function 往来单位名称(str_cus As String) As String
    Dim sql_cus As String

    Dim rs_cus As ADODB.Recordset
    sql_cus = "SELECT cCusname  FROM Customer "
    sql_cus = sql_cus & " WHERE cCuscode ='" & str_cus & "'"
    Set rs_cus = New ADODB.Recordset
    rs_cus.Open sql_cus, obj_conn(str_accno), adOpenStatic, adLockBatchOptimistic
    往来单位名称 = rs_cus!cCusname.Value

    rs_cus.Close   '完成后要关闭
    obj_conn(str_accno).Close    '完成后要关闭

End Function
Public Function 项目名称(str_citem As String) As String
    Dim sql_citem As String
    Dim rs_citem As ADODB.Recordset
    sql_citem = "SELECT citemname  FROM fitemss00 "
    sql_citem = sql_citem & " WHERE citemcode ='" & str_citem & "'"
    Set rs_citem = New ADODB.Recordset
    rs_citem.Open sql_citem, obj_conn(str_accno), adOpenStatic, adLockBatchOptimistic
    项目名称 = rs_citem!citemname.Value
    rs_citem.Close   '完成后要关闭
    obj_conn(str_accno).Close    '完成后要关闭

End Function
Public Function 一级科目() As String
    Dim sql_gradedef As String
    Dim rs_gradedef As ADODB.Recordset
    sql_gradedef = "select left(codingrule,1) from gradedef " & _
                   "where keyword='" & "code'"

    Set rs_gradedef = New ADODB.Recordset
    rs_gradedef.Open sql_gradedef, obj_conn(str_accno), adOpenStatic, adLockBatchOptimistic
    一级科目 = rs_gradedef(0).Value
    rs_gradedef.Close   '完成后要关闭
    obj_conn(str_accno).Close    '完成后要关闭

End Function
Public Function 明细期初余额(str_kmbm As String, int_kjqj As Integer) As Double
    Dim rs_kmye As ADODB.Recordset
    Dim str_zt As String
    Dim sql_kmye As String

    sql_kmye = "SELECT cbegind_c,mb"
    sql_kmye = sql_kmye & " FROM gl_accsum"
    sql_kmye = sql_kmye & " WHERE ccode=" & str_kmbm
    sql_kmye = sql_kmye & " and iperiod=" & int_kjqj
    Set rs_kmye = New ADODB.Recordset
    rs_kmye.Open sql_kmye, obj_conn(str_accno), adOpenStatic, adLockBatchOptimistic
    If rs_kmye.EOF Then
        明细期初余额 = 0
    Else

        'If rs_kmye!cbegind_c.Value = "借" Then
        明细期初余额 = rs_kmye!mb.Value
        'Else
        '明细期初余额 = 0 - rs_kmye!mb.Value
        'End If

    End If

    rs_kmye.Close   '完成后要关闭
    obj_conn(str_accno_star).Close    '完成后要关闭
End Function

Public Function 公司名称(str_acc As String) As String
    Dim rs As ADODB.Recordset
    Dim sql_text As String
    sql_text = "SELECT A.cAcc_Name as 公司名称" & _
             " FROM UA_Account as A" & _
             " WHERE A.cacc_id=" & str_acc
    Set rs = New ADODB.Recordset
    rs.Open sql_text, obj_conn("ufsystem"), adOpenStatic, adLockBatchOptimistic
    公司名称 = rs.Fields!公司名称
    rs.Close
    Set rs = Nothing
    obj_conn("ufsystem").Close
End Function
Public Function 会计基期() As Integer

    Dim rs As ADODB.Recordset
    Dim sql_text As String
    sql_text = "select min(iperiod) from gl_accvouch" & _
             " where iperiod<>0 "
    Set rs = New ADODB.Recordset
    rs.Open sql_text, obj_conn(str_accno), adOpenStatic, adLockBatchOptimistic
    会计基期 = rs.Fields(0).Value

    rs.Close
    Set rs = Nothing
    obj_conn(str_accno_star).Close
End Function

Sub 插入期初数(newsheet As String)
    With Worksheets(newsheet)
        .Range("k2").Select
        Selection.EntireRow.Insert
        .Range("e2").Value = "期初余额"
        If 余额方向(.Range("f3").Value) = False Then
            .Range("k2").Value = Abs(明细期初余额(.Range("f3").Value, 会计基期))
            .Range("k3").FormulaR1C1 = "=R[-1]C-RC[-2]+RC[-1]"
        Else
            .Range("k2").Value = 明细期初余额(.Range("f3").Value, 会计基期)
            .Range("k3").FormulaR1C1 = "=R[-1]C+RC[-2]-RC[-1]"
        End If
        .Range("k3").Select
        If .UsedRange.Rows.Count > 4 Then
            Selection.AutoFill Destination:=Range(.Cells(3, "K"), .Cells(.UsedRange.Rows.Count - 1, "K")), Type:=xlFillDefault
        End If
        .Range("k:k").NumberFormatLocal = "#,##0.00_ "
    End With
End Sub
Public Function 余额方向(str_kmbm As String) As String
    Dim sql_bproperty As String
    Dim rs_bproperty As ADODB.Recordset
    sql_bproperty = "select bproperty from code " & _
                  " where ccode=" & str_kmbm
    Set rs_bproperty = New ADODB.Recordset
    rs_bproperty.Open sql_bproperty, obj_conn(str_accno), adOpenStatic, adLockBatchOptimistic
    余额方向 = rs_bproperty(0).Value
    rs_bproperty.Close   '完成后要关闭
    obj_conn(str_accno).Close    '完成后要关闭

End Function
Sub 删除表()
    Dim xlSht As Worksheet
    Application.DisplayAlerts = False
    For Each xlSht In Sheets
        With xlSht
            If .Name <> Worksheets("凭证").Name Then
                If .Name <> Worksheets("检查表").Name Then
                    If .Visible = xlSheetVisible Then .Delete
                End If
            End If
        End With
    Next
    Application.DisplayAlerts = True
    Call 凭证筛选
End Sub

Sub 帐的期初期末金额(str_newsheet As String, str_zangsheet As String, int_col As Integer)
    Dim Dbl_yu    'str_kmdm as String
    With Worksheets(str_newsheet)
        If 余额方向(.Range("f2").Value) = False Then
            Worksheets(str_zangsheet).Cells(int_col, 7).Value = Abs(明细期初余额(.Range("f2").Value, 会计基期))

        Else
            Worksheets(str_zangsheet).Cells(int_col, 6).Value = 明细期初余额(.Range("f2").Value, 会计基期)
        End If
        Dbl_yu = (Worksheets(str_zangsheet).Cells(int_col, 6).Value + Worksheets(str_zangsheet).Cells(int_col, 8).Value) - (Worksheets(str_zangsheet).Cells(int_col, 7).Value + Worksheets(str_zangsheet).Cells(int_col, 9).Value)
        If Dbl_yu > 0 Then
            Worksheets(str_zangsheet).Cells(int_col, 10).Value = Dbl_yu
        Else
            Worksheets(str_zangsheet).Cells(int_col, 11).Value = Abs(Dbl_yu)
        End If
    End With
End Sub
Sub 帐合计(str_zangsheet As String, int_col As Integer)
    With Worksheets(str_zangsheet)
        .Cells(int_col, 3).Rows.Value = "合　　　计"
        .Cells(int_col, 4).Value = WorksheetFunction.Sum(.Range(.Cells(2, 4), .Cells(int_col - 1, 4)))
        .Cells(int_col, 5).Value = WorksheetFunction.Sum(.Range(.Cells(2, 5), .Cells(int_col - 1, 5)))
        .Cells(int_col, 6).Value = WorksheetFunction.Sum(.Range(.Cells(2, 6), .Cells(int_col - 1, 6)))
        .Cells(int_col, 7).Value = WorksheetFunction.Sum(.Range(.Cells(2, 7), .Cells(int_col - 1, 7)))
        .Cells(int_col, 8).Value = WorksheetFunction.Sum(.Range(.Cells(2, 8), .Cells(int_col - 1, 8)))
        .Cells(int_col, 9).Value = WorksheetFunction.Sum(.Range(.Cells(2, 9), .Cells(int_col - 1, 9)))
        .Cells(int_col, 10).Value = WorksheetFunction.Sum(.Range(.Cells(2, 10), .Cells(int_col - 1, 10)))
        .Cells(int_col, 11).Value = WorksheetFunction.Sum(.Range(.Cells(2, 11), .Cells(int_col - 1, 11)))
        .Cells.EntireColumn.AutoFit
        If str_zangsheet = "明细帐" Then
            .Cells(int_col, 4).Value = ""
        Else
            .Cells(int_col, 10).Value = ""
            .Cells(int_col, 11).Value = ""
        End If
        With .Range(.Cells(2, 5), .Cells(int_col, 11))
            .Style = "Comma"

        End With
    End With
End Sub

Sub 凭证筛选()
    Dim rs_pz As ADODB.Recordset
    Dim R As Integer
    Dim sql_pz, day1, linenumber
    Dim cnnstr As String
    Dim str_yer As String
    Dim str_new As String
    Dim str_acc As String
    Dim int_zzkm As Integer
    Dim m As Integer
    Dim int_star As Integer
    Dim int_end As Integer
    Dim int_i As Integer
    Dim dbl_cxje As Double
    Worksheets("凭证").Columns("A:Z").Delete
    str_acc = InputBox(prompt:="请输入帐套编号")
    int_star = InputBox(prompt:="请输入开始会计年度")
    int_end = InputBox(prompt:="请输入结束会计年度")
    dbl_cxje = InputBox(prompt:="请输入筛选金额", Default:=-100000000)

    With Worksheets("凭证")
        .Cells(1, 1).Value = "月份"
        .Cells(1, 2).Value = "凭证类别"
        .Cells(1, 3).Value = "凭证编号"
        .Cells(1, 4).Value = "日期"
        .Cells(1, 5).Value = "摘要"
        .Cells(1, 6).Value = "科目编码"
        .Cells(1, 7).Value = "总帐科目"
        .Cells(1, 8).Value = "末级科目"
        .Cells(1, 9).Value = "借方金额"
        .Cells(1, 10).Value = "贷方金额"
        .Cells(1, 11).Value = "余额"
        .Cells(1, 13).Value = "项目编号"
        .Cells(1, 14).Value = "项目名称"
        .Cells(1, 15).Value = "往来"
        .Cells(1, 16).Value = "往来编码"
        .Cells(1, 18).Value = dbl_cxje
        .Cells(1, 17).Value = 公司名称(str_acc)
        .Cells(1, 19).Value = "UFDATA_" & str_acc & "_" & int_star

        .Columns("f:f").NumberFormatLocal = "@"

    End With
    str_accno_star = "UFDATA_" & str_acc & "_" & int_star
    Worksheets("检查表").Cells(4, 1).Value = "编制单位：" & 公司名称(str_acc)
    R = Worksheets("凭证").UsedRange.Rows.Count + 1
    For int_i = int_star To int_end
        str_yer = int_i
        str_accno = "UFDATA_" & str_acc & "_" & str_yer
        int_zzkm = 一级科目
        'Call 余额表(str_accno)
        Application.DisplayAlerts = False
        str_new = str_yer & "凭证"
        Worksheets.Add(before:=Worksheets("凭证")).Name = str_new

        m = 2
        With Worksheets(str_new)
            .Columns("f:f").NumberFormatLocal = "@"
            .Cells(1, 1).Value = "月份"
            .Cells(1, 2).Value = "凭证类别"
            .Cells(1, 3).Value = "凭证编号"
            .Cells(1, 4).Value = "日期"
            .Cells(1, 5).Value = "摘要"
            .Cells(1, 6).Value = "科目编码"
            .Cells(1, 7).Value = "总帐科目"
            .Cells(1, 8).Value = "末级科目"
            .Cells(1, 9).Value = "借方金额"
            .Cells(1, 10).Value = "贷方金额"
            .Cells(1, 11).Value = "余额"
            .Cells(1, 13).Value = "项目编号"
            .Cells(1, 14).Value = "项目名称"
            .Cells(1, 15).Value = "往来"
            .Cells(1, 16).Value = "往来编码"
        End With

        sql_pz = "select * "
        sql_pz = sql_pz & " from gl_accvouch where csign + '_' + cast(iperiod as varchar(2))+'_' + cast(ino_id as varchar(10))"
        sql_pz = sql_pz & " in (select csign + '_' + cast(iperiod as varchar(2))+'_' + cast(ino_id as varchar(10)) "
        sql_pz = sql_pz & " from gl_accvouch  "
        sql_pz = sql_pz & " where IsNull(md, 0) + IsNull(mc, 0) > " & dbl_cxje
        'sql_pz = sql_pz & " and citem_id =" & InputBox(prompt:="请输入项目编号")
        'sql_pz = sql_pz & " and ( citem_id ='400100402' or ccode like'" & "%1511%" & "')"
        'sql_pz = sql_pz & " where ctext1 ='" & "川续断" & "'"
        sql_pz = sql_pz & " group by csign + '_' + cast(iperiod as varchar(2))+'_' + cast(ino_id as varchar(10)) ) order by iperiod,ino_id"
        Set rs_pz = New ADODB.Recordset
        rs_pz.Open sql_pz, obj_conn(str_accno), adOpenStatic, adLockBatchOptimistic
        While Not rs_pz.EOF
            With Worksheets(str_new)
                .Cells(m, 1).Value = rs_pz.Fields!iperiod.Value
                .Cells(m, 2).Value = rs_pz.Fields!csign.Value
                .Cells(m, 3).Value = rs_pz.Fields!ino_id.Value
                .Cells(m, 4).Value = rs_pz.Fields!dbill_date.Value
                .Cells(m, 5).Value = rs_pz.Fields!cdigest.Value
                .Cells(m, 6).Value = rs_pz.Fields!ccode.Value
                .Cells(m, 7).Value = 科目名称(rs_pz.Fields!ccode.Value, True, int_zzkm)
                .Cells(m, 8).Value = 科目名称(rs_pz.Fields!ccode.Value, False, 0)
                .Cells(m, 9).Value = rs_pz.Fields!md.Value
                .Cells(m, 10).Value = rs_pz.Fields!mc.Value
                If IsNull(rs_pz.Fields!ccus_id.Value) = False Then
                    .Cells(m, 15).Value = 往来单位名称(rs_pz.Fields!ccus_id.Value)
                    .Cells(m, 16).Value = rs_pz.Fields!ccus_id.Value
                End If
                If IsNull(rs_pz.Fields!citem_id.Value) = False Then
                    .Cells(m, 13).Value = 项目名称(rs_pz.Fields!citem_id.Value)
                    .Cells(m, 14).Value = rs_pz.Fields!citem_id.Value
                End If
            End With
            With Worksheets("凭证")
                .Cells(R, 1).Value = rs_pz.Fields!iperiod.Value
                .Cells(R, 2).Value = rs_pz.Fields!csign.Value
                .Cells(R, 3).Value = rs_pz.Fields!ino_id.Value
                .Cells(R, 4).Value = rs_pz.Fields!dbill_date.Value
                .Cells(R, 5).Value = rs_pz.Fields!cdigest.Value
                .Cells(R, 6).Value = rs_pz.Fields!ccode.Value
                .Cells(R, 7).Value = 科目名称(rs_pz.Fields!ccode.Value, True, int_zzkm)
                .Cells(R, 8).Value = 科目名称(rs_pz.Fields!ccode.Value, False, 0)
                .Cells(R, 9).Value = rs_pz.Fields!md.Value
                .Cells(R, 10).Value = rs_pz.Fields!mc.Value

                If IsNull(rs_pz.Fields!ccus_id.Value) = False Then
                    .Cells(R, 15).Value = 往来单位名称(rs_pz.Fields!ccus_id.Value)
                    .Cells(R, 16).Value = rs_pz.Fields!ccus_id.Value

                End If
                If IsNull(rs_pz.Fields!citem_id.Value) = False Then
                    .Cells(R, 13).Value = 项目名称(rs_pz.Fields!citem_id.Value)
                    .Cells(R, 14).Value = rs_pz.Fields!citem_id.Value
                End If
                .Cells(R, 19).Value = str_accno
                .Cells(1, 20).Value = int_zzkm
            End With

            R = R + 1
            m = m + 1
            rs_pz.MoveNext
        Wend
        Worksheets("凭证").Cells(2, 11).NumberFormatLocal = "#,##0.00_ "
        Worksheets("凭证").Cells(2, 11).Value = 1000000000
        With Worksheets(str_new)
            .Cells(m, 1).Value = rs_pz.RecordCount
            .Range("a:d").HorizontalAlignment = xlCenter
            .Range("f:f").HorizontalAlignment = xlCenter
            .Range("h:h").HorizontalAlignment = xlLeft
            .Range("I:J").Style = "Comma"
            .Cells.EntireColumn.AutoFit
        End With
        rs_pz.Close   '完成后要关闭
        obj_conn(str_accno).Close    '完成后要关闭
    Next
    With Worksheets("凭证")
        .Cells.EntireColumn.AutoFit
        .Range("a:d").HorizontalAlignment = xlCenter
        .Range("f:f").HorizontalAlignment = xlCenter
        .Range("h:h").HorizontalAlignment = xlLeft
        .Range("I:J").Style = "Comma"
    End With
    Application.DisplayAlerts = True
    Call 手动总帐
    Call 手动明细帐

    Worksheets("凭证").Cells(2, 11).Value = ""
    ActiveWorkbook.SaveAs str_acc & Worksheets("凭证").Cells(1, 17).Value & int_star & "年-" & int_end & "年凭证帐表"
    Application.CommandBars("Visual Basic").Visible = True
    Worksheets("凭证").Shapes("CommandButton1").Select
    Selection.Delete
    Application.CommandBars("Visual Basic").Visible = False
End Sub


