Attribute VB_Name = "Module1"
Option Explicit

Sub queryStatusNew()
'created by yujou 2021/4/26
Dim qMan As String '�d�ߪ��ܼƬ���r
Dim rowNum As Integer '�C���ܼƬ����
Dim content As String '������ܤ��e����r���A�ܼ�
Dim paySatus As Boolean '�I�ڪ��A�O���L
qMan = Range("G1").Value '�d�ߪ̬�g1�x�s�檺���e

For rowNum = 2 To 7 '�q�ĤG�C���ĤC�C
If (Cells(rowNum, "A").Value = qMan) Then '�qa�����ŦX����Ʈ�
Range("G2").Value = Cells(rowNum, 2).Value 'g2�x�s��h���ӤH���q��
If (Cells(rowNum, 3).Value = 0) Then '�P�_�I�ڪ��A�����I��
paySatus = False '�ݽᤩ�ܼƭȬ�false
Else '�w�I��
paySatus = True '�ݽᤩ�ܼƭȬ�true
End If '�����I�ڪ��A�P�_
content = qMan & "�I�ڪ��A" & paySatus '�u��������ܤ��e
MsgBox (content) '�μu���������
Else
End If '�����v��d��
Next '�j�鵲��


End Sub
 
