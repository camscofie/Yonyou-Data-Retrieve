Sub 经济分类汇总()
    Dim rs_pz As ADODB.Recordset
    Dim R As Integer
    Dim sql_pz As String
    Dim cnnstr As String
    Dim str_yer As String
    Dim str_new As String
    Dim str_acc As String
    Dim int_zzkm As Integer
    Dim m As Integer
    Dim int_star, int_i As Integer
    Dim int_end As Integer

    Worksheets("凭证").Columns("A:L").Delete
    str_acc = InputBox(prompt:="请输入帐套编号")
    int_star = InputBox(prompt:="请输入开始会计年度")
    int_end = InputBox(prompt:="请输入结束会计年度")

    R = 2
    With Worksheets("凭证")
        .Cells(1, 1).Value = "经济分类代码"
        .Cells(1, 2).Value = "经济分类"
        .Cells(1, 3).Value = "借方合计,"
        .Cells(1, 4).Value = "贷方合计"

    End With

    For int_i = int_star To int_end
        str_yer = int_i
        str_accno = "UFDATA_" & str_acc & "_" & str_yer

        Application.DisplayAlerts = False
        str_new = str_yer & "凭证"
        Worksheets.Add(before:=Worksheets("凭证")).Name = str_new

        m = 2
        With Worksheets(str_new)
            .Columns("f:f").NumberFormatLocal = "@"
            .Cells(1, 1).Value = "经济分类代码"
            .Cells(1, 2).Value = "经济分类"
            .Cells(1, 3).Value = "借方合计,"
            .Cells(1, 4).Value = "贷方合计"

        End With

        sql_pz = "select v.cAssistant2_id as 经济分类代码,"

        sql_pz = sql_pz & "( SELECT f.citemname"

        sql_pz = sql_pz & " FROM fitemss96 as f "

        sql_pz = sql_pz & " where  f.citemcode=v.cAssistant2_id) as 经济分类"

        sql_pz = sql_pz & " , Sum(v.md) AS 借方合计, Sum(v.mc) AS 贷方合计 from gl_accvouch as v"
        sql_pz = sql_pz & " where v.ccode like '" & "5001%'"
        sql_pz = sql_pz & " GROUP BY v.cAssistant2_id"
        Set rs_pz = New ADODB.Recordset
        rs_pz.Open sql_pz, obj_conn(str_accno), adOpenStatic, adLockBatchOptimistic
        While Not rs_pz.EOF
            With Worksheets(str_new)
                .Cells(m, 1).Value = rs_pz.Fields!经济分类代码
                .Cells(m, 2).Value = rs_pz.Fields!经济分类
                .Cells(m, 3).Value = rs_pz.Fields!借方合计
                .Cells(m, 4).Value = rs_pz.Fields!贷方合计

            End With
            With Worksheets("凭证")
                .Cells(R, 1).Value = rs_pz.Fields!经济分类代码
                .Cells(R, 2).Value = rs_pz.Fields!经济分类
                .Cells(R, 3).Value = rs_pz.Fields!借方合计
                .Cells(R, 4).Value = rs_pz.Fields!贷方合计

            End With

            R = R + 1
            m = m + 1
            rs_pz.MoveNext
        Wend


        rs_pz.Close   '完成后要关闭
        obj_conn(str_accno).Close    '完成后要关闭
    Next

    Application.DisplayAlerts = True
End Sub


Sub 经济分类透视表汇总()
    Dim rs_pz As ADODB.Recordset
    Dim R As Integer
    Dim sql_pz As String
    Dim cnnstr As String
    Dim str_yer As String
    Dim str_new As String
    Dim str_acc As String
    Dim int_zzkm As Integer
    Dim m As Integer
    Dim int_star, int_i As Integer
    Dim int_end As Integer

    Worksheets("凭证").Columns("A:L").Delete
    str_acc = InputBox(prompt:="请输入帐套编号")
    int_star = InputBox(prompt:="请输入开始会计年度")
    int_end = InputBox(prompt:="请输入结束会计年度")

    R = 2
    With Worksheets("凭证")
        .Cells(1, 1).Value = "经济分类代码"
        .Cells(1, 2).Value = "经济分类"
        .Cells(1, 3).Value = "借方合计,"
        .Cells(1, 4).Value = "贷方合计"

    End With

    For int_i = int_star To int_end
        str_yer = int_i
        str_accno = "UFDATA_" & str_acc & "_" & str_yer

        Application.DisplayAlerts = False
        str_new = str_yer & "凭证"
        Worksheets.Add(before:=Worksheets("凭证")).Name = str_new

        m = 2
        With Worksheets(str_new)
            .Columns("f:f").NumberFormatLocal = "@"
            .Cells(1, 1).Value = "经济分类代码"
            .Cells(1, 2).Value = "经济分类"
            .Cells(1, 3).Value = "借方合计,"
            .Cells(1, 4).Value = "贷方合计"

        End With

        sql_pz = "TRANSFORM Sum(GL_accvouch.md) AS 总计"


        sql_pz = sql_pz & "select cAssistant1_id  FROM GL_accvouch"

        sql_pz = sql_pz & " where ccode like '" & "5001%'"


        sql_pz = sql_pz & " group by a.cAssistant1_id"

        sql_pz = sql_pz & "PIVOT GL_accvouch.cAssistant2_id"

        Set rs_pz = New ADODB.Recordset
        rs_pz.Open sql_pz, obj_conn(str_accno), adOpenStatic, adLockBatchOptimistic
        While Not rs_pz.EOF
            With Worksheets(str_new)
                .Cells(m, 1).Value = rs_pz.Fields!经济分类代码
                .Cells(m, 2).Value = rs_pz.Fields!经济分类
                .Cells(m, 3).Value = rs_pz.Fields!借方合计
                .Cells(m, 4).Value = rs_pz.Fields!贷方合计

            End With
            With Worksheets("凭证")
                .Cells(R, 1).Value = rs_pz.Fields!经济分类代码
                .Cells(R, 2).Value = rs_pz.Fields!经济分类
                .Cells(R, 3).Value = rs_pz.Fields!借方合计
                .Cells(R, 4).Value = rs_pz.Fields!贷方合计

            End With

            R = R + 1
            m = m + 1
            rs_pz.MoveNext
        Wend


        rs_pz.Close   '完成后要关闭
        obj_conn(str_accno).Close    '完成后要关闭
    Next
    Application.DisplayAlerts = True
End Sub
Sub 甘孜职业凭证筛选()
    Dim rs_pz As ADODB.Recordset
    Dim R As Integer
    Dim int_acc As Integer
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
    int_star = 2010
    int_end = 2015
    int_zzkm = 3
    dbl_cxje = -100000000
    R = Worksheets("凭证").UsedRange.Rows.Count + 1

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
        .Cells(1, 13).Value = "UFDATA_" & str_acc & "_" & int_star
        .Cells(1, 14).Value = int_star
        .Cells(1, 15).Value = int_end
        .Cells(1, 16).Value = int_zzkm
        .Cells(1, 18).Value = dbl_cxje
        '.Cells(1, 17).Value = 公司名称(str_acc)
        .Columns("f:f").NumberFormatLocal = "@"

    End With
    'Call 凭证筛选
    For int_acc = 601 To 614
        str_acc = Right("000" & int_acc, 3)
        For int_i = int_star To int_end
            str_yer = int_i
            str_accno = "UFDATA_" & str_acc & "_" & str_yer

            Application.DisplayAlerts = False
            str_new = str_yer & "凭证"
            'Worksheets.Add(Before:=Worksheets("凭证")).Name = str_new


            sql_pz = "select * "
            sql_pz = sql_pz & " from gl_accvouch where csign + '_' + cast(iperiod as varchar(2))+'_' + cast(ino_id as varchar(10))"
            sql_pz = sql_pz & " in (select csign + '_' + cast(iperiod as varchar(2))+'_' + cast(ino_id as varchar(10)) "
            sql_pz = sql_pz & " from gl_accvouch  "
            sql_pz = sql_pz & " where IsNull(md, 0) + IsNull(mc, 0) > " & dbl_cxje

            sql_pz = sql_pz & " group by csign + '_' + cast(iperiod as varchar(2))+'_' + cast(ino_id as varchar(10)) ) order by iperiod,ino_id"
            Set rs_pz = New ADODB.Recordset
            rs_pz.Open sql_pz, obj_conn(str_accno), adOpenStatic, adLockBatchOptimistic
            While Not rs_pz.EOF

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
                    .Cells(R, 13).Value = str_accno
                    If IsNull(rs_pz.Fields!ccus_id.Value) = False Then
                        .Cells(m, 11).Value = 往来单位名称(rs_pz.Fields!ccus_id.Value)
                        .Cells(m, 12).Value = rs_pz.Fields!ccus_id.Value

                    End If
                End With

                R = R + 1
                m = m + 1
                rs_pz.MoveNext
            Wend


            rs_pz.Close   '完成后要关闭
            obj_conn(str_accno).Close    '完成后要关闭
        Next
    Next
    With Worksheets("凭证")
        .Cells.EntireColumn.AutoFit
        .Range("a:d").HorizontalAlignment = xlCenter
        .Range("f:f").HorizontalAlignment = xlCenter
        .Range("h:h").HorizontalAlignment = xlLeft
        .Range("I:J").Style = "Comma"
    End With
    Application.DisplayAlerts = True

End Sub


