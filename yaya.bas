Attribute VB_Name = "Module1"
Option Explicit
Sub �ʺA�X��1()


Application.DisplayAlerts = False '�@�~�t�δ�����r�A�Y�S���]�w�|�̭ȴ���
Dim i, j As Long '�ŧii�̫�Aj������� i���̫�@�C j����e�C����
Dim myrng As Range '�ŧi�d���ܼ�
'�ʺA�M��A��즳��Ƴ̫�@�C���C����
i = Cells(Rows.Count, 1).End(xlUp).Row
MsgBox "A��즳��Ƴ̫�@�C����" & i '������
    For j = i To 2 Step -1 '�q�̫�@�C��ĤG�C����A step-1���˼�
    Set myrng = Cells(j, "A") '�ثe�d��
    If myrng = myrng.Offset(-1, 0) Then '�Y�ثe��A���ȩM�e�@�C�ۦP
    myrng.Offset(-1, 0).Resize(2, 1).Merge '�h�ݥѤU�ӤW�X��
    End If
Next
Application.DisplayAlerts = True '���s�}�Ҧ۰ʴ���

End Sub


Sub �ʺA�X��2()
'�ĤG���q-�Ҧ��u�@��B�z
Dim shtIdx As Integer
'�]���Ĥ@�i�����ҥH�q�ĤG�i��
For shtIdx = 2 To Sheets.Count
Sheets(shtIdx).Activate


Application.DisplayAlerts = False '�@�~�t�δ�����r�A�Y�S���]�w�|�̭ȴ���
Dim i, j As Long '�ŧii�̫�Aj������� i���̫�@�C j����e�C����
Dim myrng As Range '�ŧi�d���ܼ�
'�ʺA�M��A��즳��Ƴ̫�@�C���C����
i = Cells(Rows.Count, 1).End(xlUp).Row
MsgBox "A��즳��Ƴ̫�@�C����" & i '������
    For j = i To 2 Step -1 '�q�̫�@�C��ĤG�C����A step-1���˼�
    Set myrng = Cells(j, "A") '�ثe�d��
    If myrng = myrng.Offset(-1, 0) Then '�Y�ثe��A���ȩM�e�@�C�ۦP
    myrng.Offset(-1, 0).Resize(2, 1).Merge '�h�ݥѤU�ӤW�X��
    End If
Next
Next
Application.DisplayAlerts = True '���s�}�Ҧ۰ʴ���

End Sub

Sub �ʺA�X��3()

Application.DisplayAlerts = False '�@�~�t�δ�����r�A�Y�S���]�w�|�̭ȴ���
Dim i, j As Long '�ŧii�̫�Aj������� i���̫�@�C j����e�C����
Dim myrng As Range '�ŧi�d���ܼ�
'�ʺA�M��A��즳��Ƴ̫�@�C���C����
i = Cells(Rows.Count, 1).End(xlUp).Row
MsgBox "A��즳��Ƴ̫�@�C����" & i '������
    For j = i To 2 Step -1 '�q�̫�@�C��ĤG�C����A step-1���˼�
    Set myrng = Cells(j, "A") '�ثe�d��
    If myrng = myrng.Offset(-1, 0) Then '�Y�ثe��A���ȩM�e�@�C�ۦP
    myrng.Offset(-1, 0).Resize(2, 1).Merge '�h�ݥѤU�ӤW�X��
    End If
Next
Application.DisplayAlerts = True '���s�}�Ҧ۰ʴ���

End Sub



