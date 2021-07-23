Attribute VB_Name = "Modulesort"
Option Explicit

Sub Demo()
Attribute Demo.VB_Description = "口罩遞減"
Attribute Demo.VB_ProcData.VB_Invoke_Func = "q\n14"
'
' Demo 巨集
' 口罩遞減
'
' 快速鍵: Ctrl+q
'
    Range("B1").Select
    ActiveWorkbook.Worksheets("工作表1").Sort.SortFields.Clear
    ActiveWorkbook.Worksheets("工作表1").Sort.SortFields.Add Key:=Range("B2:B414"), _
        SortOn:=xlSortOnValues, Order:=xlDescending, DataOption:=xlSortNormal
    With ActiveWorkbook.Worksheets("工作表1").Sort
        .SetRange Range("A1:B414")
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    Range("E1").Select
    ActiveCell.FormulaR1C1 = "=SUM(R2C2:R414C2)"
    Range("G1").Select
    ActiveCell.FormulaR1C1 = "=AVERAGE(R2C2:R[413]C[-5])"
    Range("G2").Select
End Sub
Sub demo2()
Attribute demo2.VB_Description = "口罩數量遞增"
Attribute demo2.VB_ProcData.VB_Invoke_Func = "w\n14"
'
' demo2 巨集
' 口罩數量遞增
'
' 快速鍵: Ctrl+w
'
    Range("B1").Select
    ActiveWorkbook.Worksheets("工作表1").Sort.SortFields.Clear
    ActiveWorkbook.Worksheets("工作表1").Sort.SortFields.Add Key:=Range("B2:B414"), _
        SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
    With ActiveWorkbook.Worksheets("工作表1").Sort
        .SetRange Range("A1:B414")
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    Range("G1").Select
End Sub
