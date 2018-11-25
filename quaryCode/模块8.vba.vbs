Sub 康定县公安局凭证筛选()
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
    Dim str_citem As String

    Worksheets("凭证").Columns("A:o").Delete
    str_acc = InputBox(prompt:="请输入帐套编号")
    int_star = InputBox(prompt:="请输入开始会计年度")
    int_end = InputBox(prompt:="请输入结束会计年度")
    'int_zzkm = 一级科目
    str_citem = InputBox(prompt:="请输入筛选项目编号", Default:=2)

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
        .Cells(1, 11).Value = "往来"
        .Cells(1, 12).Value = "往来编码"
        .Cells(1, 13).Value = "项目编号"
        .Cells(1, 14).Value = "项目名称"
        .Cells(1, 15).Value = int_end
        .Cells(1, 16).Value = int_zzkm
        .Cells(1, 18).Value = str_citem
        .Cells(1, 17).Value = 公司名称(str_acc)
        .Cells(1, 19).Value = "UFDATA_" & str_acc & "_" & int_star
        .Columns("f:f").NumberFormatLocal = "@"

    End With
    Worksheets("检查表").Cells(4, 1).Value = "编制单位：" & 公司名称(str_acc)
    R = Worksheets("凭证").UsedRange.Rows.Count + 1
    For int_i = int_star To int_end
        str_yer = int_i
        str_accno = "UFDATA_" & str_acc & "_" & str_yer
        int_zzkm = 一级科目
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
            .Cells(1, 11).Value = "往来"
            .Cells(1, 12).Value = "往来编码"
            .Cells(1, 13).Value = "项目编号"
            .Cells(1, 14).Value = "项目名称"
        End With

        sql_pz = "select * "

        sql_pz = sql_pz & " from gl_accvouch  "
        sql_pz = sql_pz & " where  citem_id = " & str_citem

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
                    .Cells(m, 11).Value = 往来单位名称(rs_pz.Fields!ccus_id.Value)
                    .Cells(m, 12).Value = rs_pz.Fields!ccus_id.Value
                End If
                .Cells(m, 13).Value = 项目名称(rs_pz.Fields!citem_id.Value)
                .Cells(m, 14).Value = rs_pz.Fields!citem_id.Value
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
                    .Cells(R, 11).Value = 往来单位名称(rs_pz.Fields!ccus_id.Value)
                    .Cells(R, 12).Value = rs_pz.Fields!ccus_id.Value

                End If
                .Cells(R, 13).Value = 项目名称(rs_pz.Fields!citem_id.Value)
                .Cells(R, 14).Value = rs_pz.Fields!citem_id.Value
                .Cells(R, 15).Value = str_accno
            End With

            R = R + 1
            m = m + 1
            rs_pz.MoveNext
        Wend

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
    Call 总帐
    Call 明细帐
    ActiveWorkbook.SaveAs str_acc & Worksheets("凭证").Cells(1, 17).Value & "凭证帐表"
End Sub
