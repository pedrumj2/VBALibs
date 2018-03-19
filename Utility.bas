Attribute VB_Name = "Utility"
'returns the number of rows of data from the input cell
Public Function getRowCount(ByRef sheet As Worksheet, ByVal row As Long, ByVal col As Long, Optional ByVal tol As Integer = 0) As Long
Dim rng As Range
Dim flag As Boolean
Dim i As Long
Set rng = sheet.Range(Strings.Trim(str(col)) & ":" & Strings.Trim(str(col)))
Dim tolLeft As Integer
flag = True
i = 0
tolLeft = tol
While flag = True
    If rng.Cells(row + i, 1) <> "" Then
        i = i + 1
        tolLeft = tol
    Else
        tolLeft = tolLeft - 1
        If tolLeft <= 0 Then
            getRowCount = i - tol + 1 + tolLeft
            flag = False
        Else
             i = i + 1
        End If
    End If
    
Wend

End Function


'returns the number of columns of data from the input cell
Public Function getColCount(ByRef sheet As Worksheet, ByVal row As Long, ByVal col As Long) As Long
Dim rng As Range
Dim flag As Boolean
Dim i As Long
Set rng = sheet.Rows(Strings.Trim(str(row)) & ":" & Strings.Trim(str(row)))
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
Public Function selectFile(ByVal prompt As String) As String

Dim fdgOpen As FileDialog
Set fdgOpen = Application.FileDialog(msoFileDialogOpen)
fdgOpen.Title = prompt
fdgOpen.Show
If fdgOpen.SelectedItems.count = 0 Then
    selectFile = ""
Else
    selectFile = fdgOpen.SelectedItems(1)
End If

End Function


'promts the user for a file and returns the full path of the file as output
Public Function selectFolder(ByVal prompt As String) As String

Dim fdgOpen As FileDialog
Set fdgOpen = Application.FileDialog(msoFileDialogFolderPicker)
fdgOpen.Title = prompt
fdgOpen.Show
If fdgOpen.SelectedItems.count = 0 Then
    selectFolder = ""
Else
    selectFolder = fdgOpen.SelectedItems(1)
End If

End Function


Public Function getAllFiles(ByVal folder As String) As MList
    Dim output As MList
    Dim objFSO As Object
    Dim objFolder As Object
    Dim objFile As Object
    Dim i As Integer

    Set output = New MList
    Call output.init_params(2)
       
    'Create an instance of the FileSystemObject
    Set objFSO = CreateObject("Scripting.FileSystemObject")
    'Get the folder object
    Set objFolder = objFSO.GetFolder(folder)
    i = 1
    'loops through each file in the directory and prints their names and path
    For Each objFile In objFolder.files
        'print file name
        output.add2 (objFile.name)
        'print file path
        output.add2 (objFile.Path)
    Next objFile
    Set getAllFiles = output

End Function

Public Sub clear_sheet(ByRef sheet As Worksheet)
    Dim sheetcur As Worksheet
    Set sheetcur = ActiveSheet
    sheet.Activate
    sheet.Cells.Clear
    sheetcur.Activate
End Sub
Public Sub insert_text_file(ByVal Path As String, ByRef sheet As Worksheet, ByVal row As Long, ByVal column As Long)
    Dim strTemp As String
    strTemp = "TEXT;" & Path
    sheet.Activate
    sheet.Range("A1").Select
    Application.CutCopyMode = False
    With ActiveSheet.QueryTables.add(Connection:= _
        strTemp _
        , Destination:=Range("$A$1"))
        .name = "CS1_1"
        .Refresh BackgroundQuery:=False
    End With
    ActiveWorkbook.Connections(ActiveWorkbook.Connections.count).Delete
End Sub

Public Sub delete_rows(ByRef sheet As Worksheet, ByVal row1 As Long, ByVal row2 As Long)
    sheet.Activate
    sheet.Rows(Strings.Trim(str(row1)) & ":" & Strings.Trim(str(row2))).Select
    Selection.Delete Shift:=xlUp
End Sub


