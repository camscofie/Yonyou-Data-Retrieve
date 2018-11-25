Option Explicit

Private Accid As String
Public colYear As New Collection


Property Get getYear()
    getYear = colYear
End Property

Public Function yearExist(year As String) As Boolean
    Dim yr As Variant
    For Each yr In colYear
        If yr = year Then
            yearExist = True
            Exit Function
        End If
    Next yr
    yearExist = False
End Function

Public Function addYear(year As String)
    If Not yearExist(year) Then
        colYear.Add year
    End If
End Function

Property Let cAccid(id As String)
    Accid = id
End Property

Property Get cAccid() As String
    cAccid = Accid
End Property




Public Function toString() As String
    Dim i As Long
    Dim str As String
    For i = 1 To colYear.count
        str = str + colYear(i) + ","
    Next i
    If Len(str) <> 0 Then
        str = Left(str, Len(str) - 1)
    End If

    MsgBox "cAccid: " & Accid & " has " & colYear.count & " years." & vbCrLf & _
           " They are: " & str

End Function
