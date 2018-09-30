VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CellRange"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private rng As Range
Private rowVal As Long
Private colVal As Long
Public rowCount As Long
Public colCount As Long
Private values() As Variant
Private sheet As Worksheet
'range object wrapper. Increases performance when repeatidly reading and writing to cells
Public Function init_params(ByRef sheetInput As Worksheet, ByVal row As Long, ByVal col As Long) As Integer
    Set sheet = sheetInput
    colCount = Utility.getColCount(sheet, row, col)
    rowCount = Utility.getRowCount(sheet, row, col)
    If colCount > 0 And rowCount > 0 Then
        values = sheet.Range(sheet.Cells(row, col), sheet.Cells(row + rowCount - 1, col + colCount - 1)).value
        init_params = 0
    Else
        init_params = 1
    End If
    rowVal = row
    colVal = col
    
End Function

'retrieve cell value from range object
Public Property Get Cells(ByVal i As Long, ByVal j As Long) As Variant
On Error Resume Next:
    Cells = values(i - rowVal + 1, j - colVal + 1)
If Err.Number <> 0 Then
    Err.Clear
    Cells = ""
End If
End Property

'retrieve cell value from range object
Public Property Let Cells(ByVal i As Long, ByVal j As Long, ByVal value As Variant)
    values(i - rowVal + 1, j - colVal + 1) = value
End Property


'retrieve cell value from range object
Public Sub flush()
    Dim i As Long
    Dim j As Long
    For i = 1 To rowCount
        For j = 1 To colCount
            sheet.Cells(i - rowVal + 1, j - colVal + 1) = values(i, j)
        Next j
    Next i
End Sub
