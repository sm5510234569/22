Attribute VB_Name = "Module1"
Option Explicit

Sub ResizeDemo()
Dim numRows, numColumns As Integer
numRows = Selection.Rows.Count
numColumns = Selection.Columns.Count
MsgBox "�ثe�C��" & numRows
MsgBox "�ثe���" & numColumns
Selection.Resize(numRows + 1, numColumns + 1).Select

End Sub
