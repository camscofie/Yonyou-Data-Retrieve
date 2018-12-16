Sub Macro1()
'
' Macro1 Macro
' 宏由 ZHENG 录制，时间: 2018-1-18
'

'
    Range("C:C,E:E,G:G,H:H,I:I,J:J").Select
    Range("J1").Activate
    ActiveWindow.ScrollColumn = 2
    ActiveWindow.ScrollColumn = 3
    ActiveWindow.ScrollColumn = 4
    ActiveWindow.ScrollColumn = 5
    Range("C:C,E:E,G:G,H:H,I:I,J:J,M:M,L:L,K:K,O:O,P:P").Select
    Range("P1").Activate
    ActiveWindow.ScrollColumn = 6
    ActiveWindow.ScrollColumn = 7
    ActiveWindow.ScrollColumn = 8
    ActiveWindow.ScrollColumn = 9
    ActiveWindow.ScrollColumn = 10
    ActiveWindow.ScrollColumn = 11
    ActiveWindow.ScrollColumn = 12
    ActiveWindow.ScrollColumn = 13
    ActiveWindow.ScrollColumn = 14
    Range("C:C,E:E,G:G,H:H,I:I,J:J,M:M,L:L,K:K,O:O,P:P,U:AC").Select
    Range("U1").Activate
    Selection.Delete Shift:=xlToLeft
    ActiveWindow.ScrollColumn = 14
    ActiveWindow.ScrollColumn = 13
    ActiveWindow.ScrollColumn = 12
    ActiveWindow.ScrollColumn = 11
    ActiveWindow.ScrollColumn = 10
    ActiveWindow.ScrollColumn = 9
    ActiveWindow.ScrollColumn = 8
    ActiveWindow.ScrollColumn = 6
    ActiveWindow.ScrollColumn = 5
    ActiveWindow.ScrollColumn = 4
    ActiveWindow.ScrollColumn = 3
    ActiveWindow.ScrollColumn = 2
    ActiveWindow.ScrollColumn = 1
    Range("F24").Select
End Sub
