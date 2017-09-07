Attribute VB_Name = "PublicStuff"
Option Compare Database
Option Explicit

Public blnSelectFile As Boolean
Public strPathSave As String
Public varFiles As Variant
Public appXL As Excel.Application

Public Sub GetFile(strTitle As String)

Dim intCount1 As Integer
Dim fObj As Object
Dim varItem As Variant

blnSelectFile = False
Set fObj = Application.FileDialog(msoFileDialogFilePicker)

With fObj
    .Filters.Clear
    .AllowMultiSelect = True
    .InitialFileName = "C:\"
    .Title = strTitle
    .Show
End With

If fObj.SelectedItems.Count > 0 Then
    ReDim varFiles(1 To fObj.SelectedItems.Count)
    intCount1 = 1

    For Each varItem In fObj.SelectedItems
        varFiles(intCount1) = varItem
        intCount1 = intCount1 + 1
    Next varItem

    blnSelectFile = True
End If

Set fObj = Nothing


End Sub

Public Sub SelectSaveDirectory(strTitle As String)

Dim fObj As Object

strPathSave = ""
Set fObj = Application.FileDialog(msoFileDialogFolderPicker)
fObj.Title = strTitle
fObj.Show

If fObj.SelectedItems.Count > 0 Then
    strPathSave = fObj.SelectedItems(1) & "\"
End If

Set fObj = Nothing
End Sub



