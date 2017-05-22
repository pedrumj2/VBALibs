VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "List"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private values() As Variant
Private size As Long
Public count As Long
Public Sub Class_Initialize()
    size = 0
    count = 0
    resize
  
End Sub

Private Sub resize()
    Dim temp() As Variant
    Dim i As Integer
    If count = size Then
        size = size + 100
        ReDim temp(1 To size)
        For i = 1 To size - 100
            temp(i) = values(i)
        Next i
        ReDim values(1 To size)
        For i = 1 To size - 100
            values(i) = temp(i)
        Next i
    End If
End Sub

Public Sub add(ByVal value As Variant)
    values(count + 1) = value
    count = count + 1
    resize
End Sub

Public Function Get_Val(ByVal index As Long) As Variant
    Get_Val = values(index)
End Function





