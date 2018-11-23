Sub 手动总帐()
    Dim xlSht As Worksheet
    Dim oCnn As Object
    Dim sTab, arr, c
    Dim str_sht As String
    Dim lon_pz As Long
    Dim lon_row As Long, lon_zzkm As Long
    Dim Dbl_yu
    lon_zzkm = Worksheets("凭证").Cells(1, 20).Value

    Set oCnn = CreateObject("adodb.connection")
    oCnn.Open "provider=microsoft.jet.oledb.4.0;extended properties=excel 8.0;data source=" & ThisWorkbook.FullName

    With Worksheets("凭证")
        sTab = "[" & .Name & "$a1:j" & .[a65536].End(xlUp).Row & "]"
        lon_pz = .UsedRange.Rows.Count
    End With
    arr = oCnn.Execute("select distinct 总帐科目 from " & sTab).GetRows
    Worksheets.Add(after:=Worksheets("凭证")).Name = "总帐"
    Worksheets("总帐").Cells(1, 1).Rows.Value = "序号"
    Worksheets("总帐").Cells(1, 2).Rows.Value = "总帐编码"
    Worksheets("总帐").Cells(1, 3).Rows.Value = "总帐科目"
    Worksheets("总帐").Cells(1, 4).Rows.Value = "期初借方额"
    Worksheets("总帐").Cells(1, 5).Rows.Value = "期初贷方额"
    Worksheets("总帐").Cells(1, 6).Rows.Value = "借方发生额"
    Worksheets("总帐").Cells(1, 7).Rows.Value = "贷方发生额"
    Worksheets("总帐").Cells(1, 8).Rows.Value = "期未借方金额"
    Worksheets("总帐").Cells(1, 9).Rows.Value = "期未贷方金额"
    R = 2
    For Each c In arr

        Worksheets("总帐").Cells(R, 1).Rows.Value = R - 1
        Worksheets("总帐").Cells(R, 3).Rows.Value = c
        With Worksheets("凭证")
            .Range(.Cells(1, 1), .Cells(.UsedRange.Rows.Count, 15)).AutoFilter Field:=7, Criteria1:=c
            .Range(.Cells(1, 1), .Cells(.UsedRange.Rows.Count, 15)).Copy
        End With
        Worksheets.Add(after:=Sheets(Sheets.Count)).Name = c
        Worksheets(c).Range("A1").Select
        ActiveSheet.Paste

        With Worksheets(c)    '求合计
            lon_row = .UsedRange.Rows.Count
            .Cells(lon_row + 1, "e").Value = "合   计"
            .Cells(lon_row + 1, "i").Value = WorksheetFunction.Sum(.Range(.Cells(2, 9), .Cells(lon_row, 9)))
            .Cells(lon_row + 1, "j").Value = WorksheetFunction.Sum(.Range(.Cells(2, 10), .Cells(lon_row, 10)))
            .Cells(lon_row + 1, "e").Select
            .Range(.Cells(2, 9), .Cells(lon_row + 1, "j")).Style = "comma"
            .Cells.EntireColumn.AutoFit
        End With

        With Worksheets(c)    '统计借方　贷方金额

            Worksheets("总帐").Cells(R, 6).Value = .Cells(lon_row + 1, "i")
            Worksheets("总帐").Cells(R, 7).Value = .Cells(lon_row + 1, "j").Value
            Worksheets("总帐").Cells(R, 3).Hyperlinks.Add Anchor:=Worksheets("总帐").Cells(R, 3), Address:="", SubAddress:=.Name & Chr("33") & "a2"
            Worksheets("总帐").Cells(R, 2).Value = Mid(.Cells(2, "f"), 1, lon_zzkm)
            Dbl_yu = (Worksheets("总帐").Cells(R, 4).Value + Worksheets("总帐").Cells(R, 6).Value) - (Worksheets("总帐").Cells(R, 5).Value + Worksheets("总帐").Cells(R, 7).Value)
            If Dbl_yu > 0 Then
                Worksheets("总帐").Cells(R, 8).Value = Dbl_yu
            Else
                Worksheets("总帐").Cells(R, 9).Value = Abs(Dbl_yu)
            End If

        End With
        With Worksheets("凭证")    '帐本链接凭证
            For W = 2 To lon_row
                For x = 2 To lon_pz
                    If .Cells(x, 1).Value = Worksheets(c).Cells(W, 1).Value And .Cells(x, 3).Value = Worksheets(c).Cells(W, 3).Value Then
                        Worksheets(c).Cells(W, 3).Hyperlinks.Add Anchor:=Worksheets(c).Cells(W, 3), Address:="", SubAddress:=.Name & Chr("33") & "e" & x
                        GoTo ok
                    End If
                Next x
ok:
            Next W
        End With
        R = R + 1

    Next
    With Worksheets("总帐")
        .Cells(R, 3).Rows.Value = "合　　　计"
        .Cells(R, 4).Value = WorksheetFunction.Sum(.Range(.Cells(2, 4), .Cells(R - 1, 4)))
        .Cells(R, 5).Value = WorksheetFunction.Sum(.Range(.Cells(2, 5), .Cells(R - 1, 5)))
        .Cells(R, 6).Value = WorksheetFunction.Sum(.Range(.Cells(2, 6), .Cells(R - 1, 6)))
        .Cells(R, 7).Value = WorksheetFunction.Sum(.Range(.Cells(2, 7), .Cells(R - 1, 7)))
        .Cells(R, 8).Value = WorksheetFunction.Sum(.Range(.Cells(2, 8), .Cells(R - 1, 8)))
        .Cells(R, 9).Value = WorksheetFunction.Sum(.Range(.Cells(2, 9), .Cells(R - 1, 9)))
        .Cells.EntireColumn.AutoFit

        With .Range(.Cells(2, 4), .Cells(R, 9))
            .Style = "Comma"

        End With
    End With
    oCnn.Close
    Set oCnn = Nothing
    Worksheets("凭证").Select
    Selection.AutoFilter
    Worksheets("总帐").Select
    Worksheets("总帐").Range(Cells(2, 1), Cells(R, 9)).Select
    Call 排序
    For i = 2 To R - 1
        Worksheets("总帐").Cells(i, 1).Rows.Value = i - 1
    Next

    Worksheets("总帐").Range("f2").Select
End Sub
Sub 手动明细帐()
    Dim xlSht As Worksheet
    Dim oCnn As Object
    Dim sTab, arr, c
    Dim str_sht As String
    Dim lon_pz As Long
    Dim lon_row As Long, lon_zzkm As Long
    Dim Dbl_yu
    lon_zzkm = Worksheets("凭证").Cells(1, 20).Value
    Set oCnn = CreateObject("adodb.connection")
    oCnn.Open "provider=microsoft.jet.oledb.4.0;extended properties=excel 8.0;data source=" & ThisWorkbook.FullName
    With Worksheets("凭证")
        lon_pz = .UsedRange.Rows.Count
        sTab = "[" & .Name & "$a1:j" & .[a65536].End(xlUp).Row & "]"
    End With
    arr = oCnn.Execute("select distinct 科目编码 from " & sTab).GetRows
    Worksheets.Add(after:=Worksheets("凭证")).Name = "明细帐"
    With Worksheets("明细帐")
        .Cells(1, 1).Rows.Value = "序号"
        .Cells(1, 2).Rows.Value = "总帐编码"
        .Cells(1, 3).Rows.Value = "总帐科目"
        .Cells(1, 4).Rows.Value = "未级编码"
        .Cells(1, 5).Rows.Value = "未级科目"
        .Cells(1, 6).Rows.Value = "期初借方余额"
        .Cells(1, 7).Rows.Value = "期初贷方余额"
        .Cells(1, 8).Rows.Value = "借方发生额"
        .Cells(1, 9).Rows.Value = "贷方发生额"
        .Cells(1, 10).Rows.Value = "期未借方余额"
        .Cells(1, 11).Rows.Value = "期未贷方余额"
    End With
    R = 2
    For Each c In arr
        Worksheets("明细帐").Cells(R, 1).Rows.Value = R - 1
        Worksheets("明细帐").Cells(R, 3).Rows.Value = c
        Worksheets.Add(after:=Sheets(Sheets.Count)).Name = c
        str_sht = ActiveSheet.Name
        With Worksheets("凭证")
            .Range(.Cells(1, 1), .Cells(.UsedRange.Rows.Count, 15)).AutoFilter Field:=6, Criteria1:=str_sht
            .Range(.Cells(1, 1), .Cells(.UsedRange.Rows.Count, 15)).Copy
        End With
        Worksheets(str_sht).Range("A1").Select
        ActiveSheet.Paste
        With Worksheets(str_sht)    '求合计
            lon_row = .UsedRange.Rows.Count
            .Cells(lon_row + 1, "e").Value = "合   计"
            .Cells(lon_row + 1, "i").Value = WorksheetFunction.Sum(.Range(.Cells(2, 9), .Cells(lon_row, 9)))
            .Cells(lon_row + 1, "j").Value = WorksheetFunction.Sum(.Range(.Cells(2, 10), .Cells(lon_row, 10)))
            .Cells(lon_row + 1, "e").Select
            .Range(.Cells(2, 9), .Cells(lon_row + 1, "j")).Style = "comma"
            .Cells.EntireColumn.AutoFit
        End With

        With Worksheets(str_sht)    '统计借方　贷方金额
            Worksheets("明细帐").Cells(R, 1).Rows.Value = R - 1
            Worksheets("明细帐").Cells(R, 2).Value = Mid(str_sht, 1, lon_zzkm)
            Worksheets("明细帐").Cells(R, 3).Value = .Cells(2, "g")
            Worksheets("明细帐").Cells(R, 4).Value = str_sht
            Worksheets("明细帐").Cells(R, 4).Hyperlinks.Add Anchor:=Worksheets("明细帐").Cells(R, 4), _
                                                         Address:="", SubAddress:=.Name & Chr("33") & "a2"
            Worksheets("明细帐").Cells(R, 5).Value = .Cells(2, "h")
            Worksheets("明细帐").Cells(R, 8).Value = .Cells(lon_row + 1, "i").Value
            Worksheets("明细帐").Cells(R, 9).Value = .Cells(lon_row + 1, "j").Value
            Dbl_yu = (Worksheets("明细帐").Cells(R, 6).Value + Worksheets("明细帐").Cells(R, 8).Value) - (Worksheets("明细帐").Cells(R, 7).Value + Worksheets("明细帐").Cells(R, 9).Value)
            If Dbl_yu > 0 Then
                Worksheets("明细帐").Cells(R, 10).Value = Dbl_yu
            Else
                Worksheets("明细帐").Cells(R, 11).Value = Abs(Dbl_yu)
            End If
        End With
        With Worksheets("凭证")    '帐本链接凭证
            For W = 2 To lon_row
                For x = 2 To lon_pz
                    If .Cells(x, 1).Value = Worksheets(str_sht).Cells(W, 1).Value And .Cells(x, 3).Value = Worksheets(str_sht).Cells(W, 3).Value Then
                        Worksheets(str_sht).Cells(W, 3).Hyperlinks.Add Anchor:=Worksheets(str_sht).Cells(W, 3), Address:="", SubAddress:=.Name & Chr("33") & "e" & x
                        GoTo ok
                    End If
                Next x
ok:
            Next W
        End With
        'If Worksheets("明细帐").Cells(R, "h").Value > Worksheets("明细帐").Cells(R, "i").Value Then '期初余额
        ' Worksheets("明细帐").Cells(R, 6).Value = 期初余额(str_sht, 1)
        ' Worksheets("明细帐").Cells(R, "j").FormulaR1C1 = "=RC[-4]+RC[-2]-RC[-1]"

        '  Else
        'Worksheets("明细帐").Cells(R, 7).Value = -期初余额(str_sht, 1)
        'Worksheets("明细帐").Cells(R, "k").FormulaR1C1 = "=RC[-4]+RC[-2]-RC[-3]"
        ' End If

        R = R + 1
    Next
    With Worksheets("明细帐")
        .Cells(R, 3).Rows.Value = "合　　　计"
        .Cells(R, 6).Rows.Value = WorksheetFunction.Sum(.Range(.Cells(2, 6), .Cells(R - 1, 6)))
        .Cells(R, 7).Rows.Value = WorksheetFunction.Sum(.Range(.Cells(2, 7), .Cells(R - 1, 7)))
        .Cells(R, 8).Rows.Value = WorksheetFunction.Sum(.Range(.Cells(2, 8), .Cells(R - 1, 8)))
        .Cells(R, 9).Rows.Value = WorksheetFunction.Sum(.Range(.Cells(2, 9), .Cells(R - 1, 9)))
        .Cells(R, 10).Rows.Value = WorksheetFunction.Sum(.Range(.Cells(2, 10), .Cells(R - 1, 10)))

        .Cells(R, 11).FormulaR1C1 = "=SUM(R[-" & R - 2 & "]C:R[-1]C)"
        .Cells.EntireColumn.AutoFit
        With .Range(.Cells(2, 6), .Cells(R, 11))
            .Style = "Comma"
        End With
    End With
    Worksheets("凭证").Select
    Selection.AutoFilter
    Worksheets("明细帐").Select
    Worksheets("明细帐").Range(Cells(2, 1), Cells(R, 11)).Select
    Call 排序
    For i = 2 To R - 1
        Worksheets("明细帐").Cells(i, 1).Rows.Value = i
    Next
    Worksheets("明细帐").Range("f2").Select
    oCnn.Close
    Set oCnn = Nothing
    'Call 余额("明细帐")
End Sub

Sub 排序()
    Selection.Sort Key1:=Range("B2"), Order1:=xlAscending, Key2:=Range("D2") _
                 , Order2:=xlAscending, Header:=xlGuess, OrderCustom:=1, MatchCase:= _
                   False, Orientation:=xlTopToBottom, SortMethod:=xlPinYin, DataOption1:= _
                   xlSortNormal, DataOption2:=xlSortNormal

End Sub

