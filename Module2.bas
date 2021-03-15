Attribute VB_Name = "Module2"
Sub sort2()
Attribute sort2.VB_Description = "由小~大"
Attribute sort2.VB_ProcData.VB_Invoke_Func = "j\n14"
'
' sort2 巨集
' 由小~大
'
' 快速鍵: Ctrl+j
'Create by ?郁柔 2021/3/15
    Range("B1").Select '動作1-選擇B1儲存格
    ActiveWorkbook.Worksheets("工作表1").Sort.SortFields.Clear '動作2-資料排序設定 ，根據口罩數量B欄位遞增排序
    ActiveWorkbook.Worksheets("工作表1").Sort.SortFields.Add Key:=Range("B2:B414"), _
        SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
    With ActiveWorkbook.Worksheets("工作表1").Sort '對全範圍逐行排序
        .SetRange Range("A1:B414")
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    'End of create
End Sub
Sub cal()
Attribute cal.VB_Description = "sum average"
Attribute cal.VB_ProcData.VB_Invoke_Func = " \n14"
'
' cal 巨集
' sum average
'

'
    Range("E1").Select
    ActiveCell.FormulaR1C1 = "=SUM(R[1]C[-3]:R[413]C[-3])"
    Range("E2").Select
    ActiveWindow.SmallScroll Down:=-36
    Range("G1").Select
    ActiveCell.FormulaR1C1 = "=AVERAGE(R[1]C[-5]:R[413]C[-5])"
    Range("G2").Select
    ActiveWindow.SmallScroll Down:=-27
    ActiveWorkbook.Save
    ActiveWorkbook.Save
    ActiveWorkbook.RunAutoMacros Which:=xlAutoClose
End Sub
