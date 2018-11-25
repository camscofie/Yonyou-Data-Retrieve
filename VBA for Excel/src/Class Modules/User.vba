Option Explicit


'Class User


Private uName As String
Private uPass As String
Private uServer As String
Private uDatabase As String
Public colDatabase As New Collection


'
' default constructor
Private Sub Class_Initialize()
    uServer = ""
    uName = "sa"
    uPass = ""
    uDatabase = ""
End Sub


'
' getter and setter for current opened database
Public Function SetDatabase(accNumberOrName As String, Optional accYear As String)
    If accYear = "" Then
        uDatabase = accNumberOrName
        colDatabase.Add uDatabase
        Exit Function
    End If
    uDatabase = "UFDATA_" & accNumberOrName & "_" & accYear
    colDatabase.Add uDatabase
End Function


Public Function GetDatabase() As String
    GetDatabase = uDatabase
End Function


' Collection for all the databases from user
'getter setter for user database
Private Function InsertDatabase(database As String)
    colDatabase.Add database
End Function


Public Function GetDatabaseAt(count As Integer) As String
    uDatabase = colDatabase(count)
    GetUserDatabase = colDatabase(count)
End Function


' getter setter for user name
Property Let UserName(name As String)
    uName = name
End Property

Property Get UserName() As String
    UserName = uName
End Property

' getter setter for user password

Property Let UserPass(pass As String)
    uPass = pass
End Property

Property Get UserPass() As String
    UserPass = uPass
End Property


' getter setter for user Server
Property Let UserServer(server As String)
    uServer = server
End Property

Property Get UserServer() As String
    UserServer = uServer
End Property




