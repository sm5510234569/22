Attribute VB_Name = "Module2"
Sub sort2()
Attribute sort2.VB_Description = "�Ѥp~�j"
Attribute sort2.VB_ProcData.VB_Invoke_Func = "j\n14"
'
' sort2 ����
' �Ѥp~�j
'
' �ֳt��: Ctrl+j
'Create by ?���X 2021/3/15
    Range("B1").Select '�ʧ@1-���B1�x�s��
    ActiveWorkbook.Worksheets("�u�@��1").Sort.SortFields.Clear '�ʧ@2-��ƱƧǳ]�w �A�ھڤf�n�ƶqB��컼�W�Ƨ�
    ActiveWorkbook.Worksheets("�u�@��1").Sort.SortFields.Add Key:=Range("B2:B414"), _
        SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
    With ActiveWorkbook.Worksheets("�u�@��1").Sort '����d��v��Ƨ�
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
' cal ����
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
