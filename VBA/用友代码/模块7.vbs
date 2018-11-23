Dim str_sht As String
Dim log_i, log_row As Long
Dim log_ii As Long
Dim dat_rq, lon_bh
Dim dat_rqb, lon_bhb
Dim log_star As Long
Dim log_end As Long
Dim log_add As Long

Sub 插入表()
    Sheets("检查表").Copy after:=Sheets(Sheets.Count)
    str_sht = "检查表" & log_add
    ActiveSheet.Name = str_sht
    ActiveSheet.Visible = -1
    Worksheets(str_sht).Cells(4, 1) = "编制单位：" & Worksheets("凭证").Cells(1, 17).Value
    Worksheets(str_sht).Cells(1, 12) = "索引号：" & Abs(Int(-(Sheets("凭证").UsedRange.Rows.Count / 35))) & "-" & log_add
End Sub
Sub 生成检查表()
    Dim str_zhongmi As String    '总帐加明细帐
    log_row = 8
    log_add = 1
    Call 插入表

    For log_i = 2 To Sheets("凭证").UsedRange.Rows.Count
        With Worksheets(str_sht)
            .Cells(log_row, 1) = Worksheets("凭证").Cells(log_i, 4)
            .Cells(log_row, 2) = Worksheets("凭证").Cells(log_i, 3)
            .Cells(log_row, 3) = Worksheets("凭证").Cells(log_i, 5) & Worksheets("凭证").Cells(log_i, 11)
            If Worksheets("凭证").Cells(log_i, 7) = Worksheets("凭证").Cells(log_i, 8) Then
                str_zhongmi = Worksheets("凭证").Cells(log_i, 7)
            Else
                str_zhongmi = Worksheets("凭证").Cells(log_i, 7) & "-" & Worksheets("凭证").Cells(log_i, 8)
            End If
            If Worksheets("凭证").Cells(log_i, 9) = 0 Then
                .Cells(log_row, 4) = "  贷：" & str_zhongmi
                .Cells(log_row, 5) = Worksheets("凭证").Cells(log_i, 10)
            Else
                .Cells(log_row, 4) = "借：" & str_zhongmi
                .Cells(log_row, 5) = Worksheets("凭证").Cells(log_i, 9)
            End If
        End With
        log_row = log_row + 1
        If log_row > 42 Then
            'Call 合并单元格
            log_row = 8
            log_add = log_add + 1
            Call 插入表
        End If

    Next
    'Call 合并单元格
End Sub
Sub 合并单元格()
    log_star = 8
    log_end = 9
    dat_rq = Worksheets(str_sht).Cells(8, 1).Value
    lon_bh = Worksheets(str_sht).Cells(8, 2).Value
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    For log_ii = 9 To 43
        dat_rqb = Worksheets(str_sht).Cells(log_ii, 1).Value
        lon_bhb = Worksheets(str_sht).Cells(log_ii, 2).Value
        If dat_rq = dat_rqb And lon_bh = lon_bhb Then
            log_end = log_ii
        Else

            If log_end - log_star > 0 Then

                Worksheets(str_sht).Range(Cells(log_star, 1), Cells(log_end, 1)).Select
                Call 合并
                Worksheets(str_sht).Range(Cells(log_star, 2), Cells(log_end, 2)).Select
                Call 合并


            End If

            log_star = log_ii
            dat_rq = Worksheets(str_sht).Cells(log_ii, 1)
            lon_bh = Worksheets(str_sht).Cells(log_ii, 2)
        End If
    Next
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
End Sub
Sub 合并()

    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .WrapText = True
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    Selection.Merge

End Sub

Sub 逆序打印()
    Dim int_star As Integer
    Dim int_end As Integer
    Dim int_new As Integer
    int_star = InputBox("请输入开始打印页")
    int_end = Worksheets.Count
    For i = int_end To int_star Step -1
        Worksheets(i).PrintOut
    Next
End Sub

Sub 凭证月编号()
'

    Columns("D:D").Select
    Selection.Insert Shift:=xlToRight
    Range("D2").Select
    ActiveCell.FormulaR1C1 = "=RC[-3]&""月第""&RC[-1]&""号"""
    Range("D2").Select
    Selection.AutoFill Destination:=Range("D2:D15"), Type:=xlFillDefault
    Columns("E:E").Select
    Selection.Cut
    Columns("D:D").Select
    Selection.Insert Shift:=xlToRight
    Columns("D:D").Select
    Selection.NumberFormatLocal = "[$-F800]dddd, mmmm dd, yyyy"
    Columns("J:J").Select
    Selection.Cut
    Columns("F:F").Select
    Selection.Insert Shift:=xlToRight
End Sub

