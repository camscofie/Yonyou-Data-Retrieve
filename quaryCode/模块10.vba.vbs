Sub 录入服务单位名称()
'update  code set  ccode_name='广元市剑阁建工建材有限公司'
'where ccode_name='施工单位'
    Dim rs As New ADODB.Recordset
    Set rs = New ADODB.Recordset
    Dim obj_sheet As Worksheet
    Dim sql_text As String
    Dim sql_text_one As String

    Dim sql_text_two As String
    Dim sql_text_go As String
    sql_text_one = "update  code set  ccode_name='"
    sql_text_two = "where ccode_name='"
    sql_text_go = "go"
    Set obj_sheet = Worksheets("服务单位")

    For i = 2 To obj_sheet.UsedRange.Rows.Count
        If obj_sheet.Cells(i, "h") <> "" Then
            sql_text = sql_text & sql_text_one & obj_sheet.Cells(i, "h").Value & "'" & Chr(13)
            sql_text = sql_text & sql_text_two & obj_sheet.Cells(i, "g").Value & "'" & Chr(13)
            sql_text = sql_text & sql_text_go & Chr(13)
        End If
        MsgBox sql_text
    Next
End Sub

