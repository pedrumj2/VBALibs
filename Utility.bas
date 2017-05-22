Attribute VB_Name = "Utility"
'returns the number of rows of data from the input cell
Public Function getRowCount(ByRef sheet As Worksheet, ByVal row As Long, ByVal col As Long) As Long
Dim rng As Range
Dim flag As Boolean
Dim i As Long
Set rng = sheet.Range(Strings.Trim(Str(col)) & ":" & Strings.Trim(Str(col)))
flag = True
i = 0
While flag = True
    If rng.Cells(row + i, 1) <> "" Then
        i = i + 1
    Else
        getRowCount = i
        flag = False
    End If
    
Wend
End Function

'returns the number of columns of data from the input cell
Public Function getColCount(ByRef sheet As Worksheet, ByVal row As Long, ByVal col As Long) As Long
Dim rng As Range
Dim flag As Boolean
Dim i As Long
Set rng = sheet.Rows(Strings.Trim(Str(row)) & ":" & Strings.Trim(Str(row)))
flag = True
i = 0
While flag = True
    If rng.Cells(1, col + i) <> "" Then
        i = i + 1
    Else
        getColCount = i
        flag = False
    End If
    
Wend
End Function

'promts the user for a file and returns the full path of the file as output
Public Function getFile(ByVal prompt As String) As String

Dim fdgOpen As FileDialog
Set fdgOpen = Application.FileDialog(msoFileDialogOpen)
fdgOpen.Title = prompt
fdgOpen.Show
If fdgOpen.SelectedItems.count = 0 Then
    getFile = ""
Else
    getFile = fdgOpen.SelectedItems(1)
End If


End Function

