Option Explicit

Public sqlDict As Scripting.Dictionary


Public Function initSQL()
    Set sqlDict = New Scripting.Dictionary

    '   sql.Add "dbScan", "SELECT cAcc_Id, iYear FROM UA_Account"

    sqlDict.Add "accVouch", "SELECT iperiod,csign, ino_id,dbill_date,cdigest, ccode, md, mc   FROM GL_accvouch"
End Function

Function addSQL(sqlName As String, sqlQuary As String)
    initSQL
    sql.Add sqlName, sqlQuary
End Function

Function quary(str As String)
    initSQL
    quary = sql(str)
End Function


