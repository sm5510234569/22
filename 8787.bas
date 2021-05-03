Attribute VB_Name = "Module1"
Option Explicit

Sub ResizeDemo()
Dim numRows, numColumns As Integer
numRows = Selection.Rows.Count
numColumns = Selection.Columns.Count
MsgBox "目前列數" & numRows
MsgBox "目前欄數" & numColumns
Selection.Resize(numRows + 1, numColumns + 1).Select

End Sub
