Attribute VB_Name = "Modulesort"
Option Explicit

Sub Demo()
Attribute Demo.VB_Description = "�f�n����"
Attribute Demo.VB_ProcData.VB_Invoke_Func = "q\n14"
'
' Demo ����
' �f�n����
'
' �ֳt��: Ctrl+q
'
    Range("B1").Select
    ActiveWorkbook.Worksheets("�u�@��1").Sort.SortFields.Clear
    ActiveWorkbook.Worksheets("�u�@��1").Sort.SortFields.Add Key:=Range("B2:B414"), _
        SortOn:=xlSortOnValues, Order:=xlDescending, DataOption:=xlSortNormal
    With ActiveWorkbook.Worksheets("�u�@��1").Sort
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
Attribute demo2.VB_Description = "�f�n�ƶq���W"
Attribute demo2.VB_ProcData.VB_Invoke_Func = "w\n14"
'
' demo2 ����
' �f�n�ƶq���W
'
' �ֳt��: Ctrl+w
'
    Range("B1").Select
    ActiveWorkbook.Worksheets("�u�@��1").Sort.SortFields.Clear
    ActiveWorkbook.Worksheets("�u�@��1").Sort.SortFields.Add Key:=Range("B2:B414"), _
        SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
    With ActiveWorkbook.Worksheets("�u�@��1").Sort
        .SetRange Range("A1:B414")
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    Range("G1").Select
End Sub
