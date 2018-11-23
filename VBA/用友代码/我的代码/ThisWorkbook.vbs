Public str_server As String


Private Sub Workbook_Open()
'第一课 hello
'一、目标：在excll中加载登录窗体，以便用户选择帐套及会计年度
'二、需要用到的数据库表
'1、  ufsystem中UA-account
'2、  ufsystem中UA-user(
'3、  UA_Account_sub(账套年度表)
'三、我做的程序
'我编的程序是后置连接数据库文件，我用得比较顺手，我知道打开要客户的那套帐，但是客户觉得非常难用，因为他们不知道帐所对应的数据库文件，你要解决就是要把数据库连接前置

    Dim str_lj As String
    Dim str_excel As String
    Dim bool_ZhaoTao As Integer
    Dim Zhaotao As String

    On Error Resume Next

    ActiveWindow.Capion = ActiveWorkbook.Name & "       四川宏康会计师事务所  友情制作"

    str_server = InputBox(prompt:="请输入你要登录的数据库服务器", Default:="VIRSCHKS")

    If Not str_server = "" Then

        bool_ZhaoTao = MsgBox(prompt:="是否现在打开帐套?", Buttons:=vbYesNo + vbQuestion)

        If bool_ZhaoTao = vbYes Then
            帐套.帐务查询
        Else
            MsgBox "帐套可在之后的的窗体内打开"
        End If

    End If

End Sub









