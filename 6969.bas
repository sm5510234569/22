Attribute VB_Name = "Module1"
Option Explicit

Sub queryStatusNew()
'created by yujou 2021/4/26
Dim qMan As String '查詢者變數為文字
Dim rowNum As Integer '列數變數為整數
Dim content As String '視窗顯示內容為文字型態變數
Dim paySatus As Boolean '付款狀態是布林
qMan = Range("G1").Value '查詢者為g1儲存格的內容

For rowNum = 2 To 7 '從第二列找到第七列
If (Cells(rowNum, "A").Value = qMan) Then '從a欄位找到符合的資料時
Range("G2").Value = Cells(rowNum, 2).Value 'g2儲存格則為該人員電話
If (Cells(rowNum, 3).Value = 0) Then '判斷付款狀態為未付款
paySatus = False '需賦予變數值為false
Else '已付款
paySatus = True '需賦予變數值為true
End If '結束付款狀態判斷
content = qMan & "付款狀態" & paySatus '彈跳視窗顯示內容
MsgBox (content) '用彈跳視窗顯示
Else
End If '結束逐行查詢
Next '迴圈結尾


End Sub
 
