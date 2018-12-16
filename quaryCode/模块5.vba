Sub 删除结转分录()
    Dim int_i As Integer
    Dim int_row As Integer
    Dim wks_scb As Worksheet    'wks_scb 要删除数据的表
    Set wks_scb = ActiveSheet
    Dim str_sc As String  '删除的字符
    int_row = 2
    str_sc = InputBox(prompt:="请输入要删除会计分录中的关健字")
    With wks_scb
        Do While .Cells(int_row, 1).Value <> ""
            'MsgBox .Cells(int_row, 5).Value & str_sc
            If .Cells(int_row, 5).Value = str_sc Then
                .Rows(int_row).Delete
                int_row = int_row - 1
            Else
                int_row = int_row + 1
            End If

        Loop

    End With
End Sub
Sub 删除结转()
    Dim ran_c
    Dim str_adr As String
    Dim int_len As Integer
    Dim int_row As Integer
    Dim wks_scb As Worksheet    'wks_scb 要删除数据的表
    Set wks_scb = ActiveSheet
    Dim str_sc As String  '删除的字符

    str_sc = InputBox(prompt:="请输入要删除会计分录中的关健字")
    With wks_scb.UsedRange

        Set c = .Find(str_sc, LookIn:=xlValues)
        If Not c Is Nothing Then
            firstAddress = c.Address(ReferenceStyle:=xlR1C1)
            Do
                str_adr = c.Address(ReferenceStyle:=xlR1C1)
                int_len = Len(str_adr)
                int_instr = InStr(str_adr, "C")
                int_row = Right(str_adr, int_len - int_instr)
                .Rows(int_row).Delete
                Set c = .FindNext(c)
            Loop While Not c Is Nothing And c.Address <> firstAddress
        End If

    End With
End Sub

