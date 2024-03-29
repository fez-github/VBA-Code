Option Explicit
Sub All_Files_In_Folder()
'Purpose: To run procedures on all files in a folder.
'Test capabilities of opening files with a name.
'New FileSelect procedure makes this code less useful, but this could still be used as a base for later.
'   Ex. We need to run procedures on all files in a folder, and it's easier to not have the user select it.
'Last Updated: July 22, 2021

'Declare variables
Dim Folder As String, file As String, SaveName As String
Dim FileCount As Long, SuccessCount As Long, FailureCount As Long, FolderLoop As Long, CountLastRow As Long
Dim checkWB As Workbook, currentWB As Workbook
Dim confirm As Variant

'Confirmation MsgBox
    confirm = MsgBox("Once activated, this macro cannot be undone.  Please ensure you meet the following conditions: " & vbCrLf & _
        "-Keep a backup copy of the folder you wish to operate on, just in case." & vbCrLf & _
        "-No changes have been made to the Settlement Report.", vbYesNo, "Warning")
        If confirm <> vbYes Then
            Exit Sub
        End If

'Pull folder name
Folder = FolderSelect(Folder, "Please select the folder in which you want to operate on.", "Select the folder you want to operate on.")
        If Folder = "" Then
            Exit Sub
        End If

'Hard-code folder name here
'Folder = "Insert folder filepath here."
file = Dir(Folder & "\*")

'Disable screen updating and display alerts so the procedure can act with no disrupting visuals.
Application.ScreenUpdating = False
Application.DisplayAlerts = False

'Determine number of files within folder, and use that number to create an array.
    Do Until file = ""
        FileCount = FileCount + 1
        file = Dir
    Loop
    Dim CompletedList() As String
    ReDim CompletedList(1 To FileCount, 1 To 2)
    file = Dir(Folder & "\*")

'Open all files.  Error handler is for any files that could not open.
For FolderLoop = 1 To FileCount
    On Error Resume Next
        Set currentWB = Workbooks.Open(Folder & "\" & file)
        If Err <> 0 Then
            'Adds a note to the array that this file did not open.
            On Error GoTo 0
            CompletedList(FolderLoop, 1) = file
            CompletedList(FolderLoop, 2) = "Failure"
            file = Dir
        Else
            'After closing the error handler, we run the code you want on the file, then close the workbook & note the file as a success.
            On Error GoTo 0
'---------Insert block of code on next line--------

'---------Inserted block of code ends here---------
            'Saves a new copy of the file as a CSV, and adds a note to the array that this file did not open.
            currentWB.SaveAs filename:="CSV" & GetFilenameFromPath(currentWB.name), FileFormat:=xlCSVUTF8
            currentWB.Close
            CompletedList(FolderLoop, 1) = file
            CompletedList(FolderLoop, 2) = "Success"
            file = Dir
        End If
Next FolderLoop
Application.DisplayAlerts = True
Application.ScreenUpdating = True
'Create a new workbook and use array to list what files were and were not opened.
    Set checkWB = Workbooks.Add
     With checkWB.Sheets(1)
            .Cells(1, 1).value = "FileName"
            .Cells(1, 2).value = "Status"
        For FolderLoop = 1 To FileCount
            .Cells(FolderLoop + 1, 1).value = CompletedList(FolderLoop, 1)
            .Cells(FolderLoop + 1, 2).value = CompletedList(FolderLoop, 2)
        Next FolderLoop
        CountLastRow = .Range("A" & rows.Count).End(xlUp).row
        SuccessCount = Application.WorksheetFunction.CountIf(.Range("B2:B" & CountLastRow), "Success")
        FailureCount = Application.WorksheetFunction.CountIf(.Range("B2:B" & CountLastRow), "Failure")
     End With
'Inform user that the process is done.
    MsgBox "Procedure is finished with " & SuccessCount & " successes and " & FailureCount & " failures out of " & FileCount & " files.", vbCritical, "Complete"
End Sub

Private Function FolderSelect(ByVal folderpath As String, message As String, Title As String)
'Modified version of Fileselect meant to work with folders.
'Automates the process of selecting an external workbook.
Dim WK As FileDialog
        Set WK = Application.FileDialog(msoFileDialogFolderPicker)
    MsgBox (message)
         With WK
            .AllowMultiSelect = False
            .Title = Title
             If .Show = False Then
                Exit Function
             End If
         FolderSelect = .selectedItems.Item(1)
         End With
'Cannot use multiple .SelectedItems.Item values, so multi file select is not possible.
    'Index number assigned to the file is dependent on their order on the dropdown list, not the order you pick them.
End Function

Private Function GetFilenameFromPath(ByVal strPath As String) As String
' Returns the rightmost characters of a string upto but not including the rightmost '\'
'Done in order to use the filename for declaring variables, as the Filepath itself cannot be used in a workbook/sheet name declaration.
' e.g. 'c:\winnt\win.ini' returns 'win.ini'
'Code from Stackoverflow: https://stackoverflow.com/questions/1743328/how-to-extract-file-name-from-path

    If Right$(strPath, 1) <> "\" And Len(strPath) > 0 Then
        GetFilenameFromPath = GetFilenameFromPath(Left$(strPath, Len(strPath) - 1)) + Right$(strPath, 1)
    End If
End Function
