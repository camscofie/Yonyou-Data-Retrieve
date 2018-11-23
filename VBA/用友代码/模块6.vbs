

Sub 用友凭证连续编号第二版()
    Dim cn_pz As ADODB.Connection
    Dim rs_pz As ADODB.Recordset
    Dim R As Integer
    Dim sql_pz, day1, linenumber
    Dim cnnstr As String
    Dim str_acc As String
    Dim str_accno As String
    Dim int_new, int_new_xh As Integer
    Worksheets("凭证").Columns("A:L").Delete
    R = 2
    str_acc = InputBox(prompt:="请输入帐套编号")
    int_star = InputBox(prompt:="请输入开始会计年度")
    int_end = InputBox(prompt:="请输入结束会计年度")
    int_new_xh = InputBox(prompt:="请输入拟重新编号的开始数字", Default:=1)
    For int_i = int_star To int_end
        int_new = int_new_xh
        str_accno = "UFDATA_" & str_acc & "_" & int_i
        Set cn_pz = New ADODB.Connection
        sql_pz = "select distinct ino_id ,iperiod,dbill_date " & _
                 "from gl_accvouch " & _
                 "where  csign + '_' + cast(iperiod as varchar(2))+'_' + cast(ino_id as varchar(10))" & _
               " in (" & _
                 "select csign + '_' + cast(iperiod as varchar(2))+'_' + cast(ino_id as varchar(10)) " & _
               " from gl_accvouch " & _
               " group by csign + '_' + cast(iperiod as varchar(2))+'_' + cast(ino_id as varchar(10))  " & _
                 ")" & _
                 "order by dbill_date,iperiod,ino_id"

        Set rs_pz = New ADODB.Recordset
        rs_pz.Open sql_pz, obj_conn(str_accno), adOpenStatic, 3
        Application.DisplayAlerts = False
        While Not rs_pz.EOF

            With Worksheets("凭证")
                .Cells(R, 1).Value = int_new
                .Cells(R, 3).Value = rs_pz.Fields!ino_id.Value
                .Cells(R, 4).Value = rs_pz.Fields!iperiod.Value
            End With
            rs_pz.CancelUpdate
            rs_pz("ino_id") = int_new

            rs_pz.Update
            R = R + 1
            int_new = int_new + 1
            rs_pz.MoveNext

        Wend
        rs_pz.Close   '完成后要关闭
        obj_conn(str_accno).Close    '完成后要关闭
    Next

    MsgBox "恭喜你更改" & R - 2 & "个会计凭证成功"


    Application.DisplayAlerts = True
End Sub


