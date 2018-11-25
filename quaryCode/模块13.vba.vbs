Sub 科目导入()
    Dim year As Integer
    Dim sql_text As String
    Dim sql_text_old As String
    Dim rs_code_old As ADODB.Recordset
    Dim obj_rs As ADODB.Recordset
    Dim str_pc As String
    Dim str_accno As String
    Dim str_no As String
    year = 2010
    str_no = InputBox(prompt:="请输入帐套编号")
    str_accno = "UFDATA_" & str_no
    str_accno = str_accno & "_" & year
    sql_text_old = "select cclass,cclass_engl,ccode,ccode_name,ccode_engl,igrade,bproperty,cbook_type,cbook_type_engl,bitem,cass_item,bend,bd_c from code2010"
    Set rs_code_old = New ADODB.Recordset
    rs_code_old.Open sql_text_old, obj_conn(str_accno), adOpenKeyset, adLockOptimistic



    sql_text = "SELECT i_id,cclass,cclass_engl, ccode,ccode_name,ccode_engl,igrade,"
    sql_text = sql_text & "bproperty,cbook_type,cbook_type_engl,bend,bd_c,bitem,cass_item  FROM code "

    Set obj_rs = New ADODB.Recordset
    obj_rs.Open sql_text, obj_conn(str_accno), adOpenKeyset, adLockOptimistic



    'For int_II = 2 To Worksheets("金财科目").UsedRange.Rows.Count

    Do While Not rs_code_old.EOF    '当数据指针未移到记录集末尾时，循环下列操作

        obj_rs.AddNew
        obj_rs!cclass = rs_code_old!cclass
        obj_rs!cclass_engl = rs_code_old!cclass_engl
        obj_rs!ccode = rs_code_old!ccode
        obj_rs!ccode_name = rs_code_old!ccode_name
        obj_rs!ccode_engl = rs_code_old!ccode_engl
        obj_rs!igrade = rs_code_old!igrade
        obj_rs!bproperty = rs_code_old!bproperty
        obj_rs!bend = rs_code_old!bend
        obj_rs!cbook_type = "金额式"
        obj_rs!cbook_type_engl = "JES"
        obj_rs!bd_c = rs_code_old!bd_c
        obj_rs!bitem = rs_code_old!bitem
        obj_rs!cass_item = rs_code_old!cass_item
        obj_rs.Update
        'MsgBox rs_code_old!ccode_name
        rs_code_old.MoveNext
    Loop
    'Next
    MsgBox "输入完成"
    Set obj_rs = Nothing
    Set obj_rs_old = Nothing
    obj_conn(str_accno).Close


End Sub
