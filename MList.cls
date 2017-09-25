VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "MList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private values() As Variant
Private size As Long
Public count As Long
Private colCount As Long
Private currColIndex2Add As Long

Public Sub Class_Initialize()
    size = 0
    count = 0
    currColIndex2Add = 0
End Sub

Public Sub init_params(ByVal colCountInput As Long)
    colCount = colCountInput
    resize
End Sub

Private Sub resize()
    Dim temp() As Variant
    Dim i As Integer
    If count = size Then
        size = size + 100
        ReDim Preserve values(colCount - 1, size - 1)
    End If
End Sub

Public Sub add(ByRef value() As Variant)
    Dim i As Long
    For i = 0 To colCount - 1
        values(i, count + 1) = value(i)
    Next i
    count = count + 1
    resize
End Sub


Public Sub add2(ByVal value As Variant)
    Dim i As Long
    values(currColIndex2Add, count) = value
    currColIndex2Add = currColIndex2Add + 1
    If currColIndex2Add = colCount Then
        currColIndex2Add = 0
        count = count + 1
        resize
    End If
    
    
End Sub

Public Function Get_Val(ByVal i As Long, ByVal j As Long) As Variant
    Get_Val = values(i, j)
End Function






