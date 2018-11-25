Option Explicit
Public colAccid(50) As New cAccid

Function parseRecordset(dbRecord As Recordset)
    If Not dbRecord.EOF Then
        Do While Not dbRecord.EOF
            Call cAccidRegistry(dbRecord.Fields!cAcc_Id.value, dbRecord.Fields!iYear.value)
            dbRecord.MoveNext
        Loop
    End If
End Function



Private Function cAccidRegistry(id As String, year As String)
' check if the cAccid is regitered
    Dim iter As Integer
    For iter = LBound(colAccid) To UBound(colAccid)
        If colAccid(iter).cAccid = id Then
            ' check if year is also regitered
            If colAccid(iter).yearExist(year) Then
                Exit Function
            End If
        End If
    Next iter


    ' at here means either is year not exist, or user not exist
    Dim count As Integer
    For count = LBound(colAccid) To UBound(colAccid)
        If colAccid(count).cAccid = "" Then
            With colAccid(count)
                .cAccid = id
                .addYear year
            End With
            Exit Function
        End If
    Next count

End Function
