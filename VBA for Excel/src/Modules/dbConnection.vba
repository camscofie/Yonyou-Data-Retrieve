Option Explicit

Private Const serverErr = -2147467259
Private Const accountErr = -2147217843
Private dbRecord As ADODB.Recordset


Public Function dbScan()
    Dim user As New user
    Dim sql As New sqlQuary
    Call user.SetDatabase("UFSystem")
    Call dbConnection.initDB(user, "SELECT cAcc_Id, iYear FROM UA_Account")
    Call Data.parseRecordset(dbRecord)
    MsgBox "数据库扫描成功"
End Function



' this function initialize a DB connection
' need to close after use this method
Public Function initDB(user As user, quary As String) As ADODB.Connection
    Set initDB = New ADODB.Connection
    Set dbRecord = New ADODB.Recordset
    Dim field As Variant

    On Error GoTo Catch     ' try-catch

    With initDB
        .ConnectionTimeout = 3
        .Provider = "MSDASQL"
        .ConnectionString = "driver={SQL Server};" & _
                            "server=" & user.UserServer & _
                            ";uid=" & user.UserName & _
                            ";pwd=" & user.UserPass & _
                            ";database=" & user.GetDatabase
        .Open
    End With

    If Not initDB.State = adStateOpen Then
        MsgBox "数据库加载失败"
        Exit Function
    End If

    dbRecord.Open quary, initDB


Done:
    Exit Function



    ' Error handeling
Catch:
    If Err.Number = serverErr Then
        MsgBox "没有找到指定数据库，请检查数据库名称，帐套编号以及会计年度有无错误。", vbCritical, "错误"
    ElseIf Err.Number = accountErr Then
        MsgBox "输入的用户名或密码有误，请重新输入。", vbCritical, "错误"
    Else
        MsgBox Err.Description, vbCritical, "错误"
    End If


End Function

