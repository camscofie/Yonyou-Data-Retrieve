Sub 帐套目录()

    Dim rs As ADODB.Recordset
    Dim sql_text As String
    Dim str_new As String
    Dim int_col, m, int_row As Integer
    str_new = "帐套一览表"
    int_col = 2
    On Error Resume Next
    If Worksheets(str_new) Is Nothing Then
        Worksheets.Add(after:=Worksheets(Worksheets.Count)).Name = str_new
    End If
    sql_text = "SELECT *" & _
             " FROM UA_Account "
    Set rs = New ADODB.Recordset
    rs.Open sql_text, obj_conn("ufsystem"), adOpenStatic, adLockBatchOptimistic
    For int_row = 1 To rs.Fields.Count - 1
        Worksheets(str_new).Cells(1, int_row).Value = rs(int_row).Name
    Next
    While Not rs.EOF
        For int_row = 1 To rs.Fields.Count - 1
            Worksheets(str_new).Cells(int_col, int_row).Value = rs(int_row).Value
        Next
        int_col = int_col + 1

        rs.MoveNext
    Wend
    rs.Close
    Set rs = Nothing
    obj_conn("ufsystem").Close

End Sub

