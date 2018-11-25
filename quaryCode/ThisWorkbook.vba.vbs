
Private Sub Workbook_Open()
'2015年12月25日 修改版
    Dim str_lj As String
    Dim str_excel As String
    On Error Resume Next

    ActiveWindow.Caption = ActiveWorkbook.Name & "  四川宏康会计师事务所  友情制作"

    str_server = InputBox(prompt:="请输入你要登录的数据库服务器", Default:="VIRSCHKS")

End Sub


