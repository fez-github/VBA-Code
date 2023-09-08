'EXTERNAL WORKBOOK SELECT: PART 1A
Public Function FileSelect(message As String, Title As String)
'Purpose: Automates the process of selecting an external workbook.
'Last Updated: July 26, 2021

Dim WK As FileDialog
        Set WK = Application.FileDialog(msoFileDialogFilePicker)
    MsgBox (message)
         With WK
            .AllowMultiSelect = False
            .Title = Title
            If .Show = False Then
                Exit Function
            End If
                FileSelect = .selectedItems.Item(1)
         End With
End Function

'EXTERNAL WORKBOOK SELECT: PART 1B
Public Function MultiFileSelect(filename() As String, message As String, Title As String)
'Purpose: Automates the process of selecting multiple external workbooks.
'Last Updated: July 26, 2021

Dim i As Integer
Dim WK As FileDialog
        Set WK = Application.FileDialog(msoFileDialogFilePicker)
    MsgBox (message)
         With WK
            .AllowMultiSelect = True
            .Title = Title
            If .Show = False Then
                Exit Function
            End If
            ReDim filename(.selectedItems.Count - 1)
            For i = 1 To .selectedItems.Count
                filename(i - 1) = .selectedItems.Item(i)
            Next i
         End With
         MultiFileSelect = filename()
End Function

'EXTERNAL WORKBOOK SELECT: PART 2
Public Sub CheckOpenWorkbook(filename As String)
'Source: https://www.mrexcel.com/board/threads/vba-code-to-open-an-excel-file-only-if-not-already-open.431458/
If CheckFileIsOpen(filename) = False Then
    Workbooks.Open filename
End If
End Sub

'EXTERNAL WORKBOOK SELECT: PART 3
Function CheckFileIsOpen(chkSumfile As String) As Boolean
'Source: https://www.mrexcel.com/board/threads/vba-code-to-open-an-excel-file-only-if-not-already-open.431458/
    On Error Resume Next
    CheckFileIsOpen = (Workbooks(chkSumfile).name = chkSumfile)
    On Error GoTo 0
End Function

'EXTERNAL WORKBOOK SELECT: PART 4
Public Function GetFilenameFromPath(strPath As String) As String
' Returns the rightmost characters of a string upto but not including the rightmost '\'
'Done in order to use the filename for declaring variables, as the filename itself cannot be used in a workbook/sheet name declaration.
' e.g. 'c:\winnt\win.ini' returns 'win.ini'
'Code from Stackoverflow: https://stackoverflow.com/questions/1743328/how-to-extract-file-name-from-path

    If Right$(strPath, 1) <> "\" And Len(strPath) > 0 Then
        GetFilenameFromPath = GetFilenameFromPath(Left$(strPath, Len(strPath) - 1)) + Right$(strPath, 1)
    End If
End Function
