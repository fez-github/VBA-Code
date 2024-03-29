Public Enum PasteSpecialType
'Enum C&P'd from: https://docs.microsoft.com/en-us/office/vba/api/excel.xlpastetype
    xlPasteAll = -4104    'Everything will be pasted.
    xlPasteAllExceptBorders = 7    'Everything except borders will be pasted.
    xlPasteAllMergingConditionalFormats = 14   'Everything will be pasted and conditional formats will be merged.
    xlPasteAllUsingSourceTheme = 13  'Everything will be pasted using the source theme.
    xlPasteColumnWidths = 8   'Copied column width is pasted.
    xlPasteComments = -4144   'Comments are pasted.
    xlPasteFormats = -4122    'Copied source format is pasted.
    xlPasteFormulas = -4123   'Formulas are pasted.
    xlPasteFormulasAndNumberFormats = 11  'Formulas and Number formats are pasted.
    xlPasteValidation = 6     'Validations are pasted.
    xlPasteValues = -4163     'Values are pasted.
    xlPasteValuesAndNumberFormats = 12    'Values and Number formats are pasted.
End Enum

Public Enum PasteSpecialOperation
'Enum C&P'd from https://docs.microsoft.com/en-us/office/vba/api/excel.xlpastespecialoperation
    xlPasteSpecialOperationAdd = 2    'Copied data will be added to the value in the destination cell.
    xlPasteSpecialOperationDivide = 5     'Copied data will divide the value in the destination cell.
    xlPasteSpecialOperationMultiply = 4   'Copied data will multiply the value in the destination cell.
    xlPasteSpecialOperationNone = -4142   'No calculation will be done in the paste operation.
    xlPasteSpecialOperationSubtract = 3   'Copied data will be subtracted from the value in the destination cell.
End Enum

Public Sub Paste_New_Range(sourceRange As Range, destinationRange As Range, headers As Boolean, _
    Optional typePaste As PasteSpecialType, Optional operator As PasteSpecialOperation, _
    Optional skipBlankRows As Boolean, Optional transposeRange As Boolean)
'Purpose: To group multiple ranges from across workbooks/worksheets together more easily.
'Last Updated: July 21, 2021
    
'Declare variables
    Dim wb As Workbook
        Set wb = destinationRange.Worksheet.Parent
    Dim ws As Worksheet
        Set ws = destinationRange.Worksheet
'Check for blank values
    If typePaste = 0 Then
        typePaste = xlPasteAll
    End If
    If operator = 0 Then
        operator = xlPasteSpecialOperationNone
    End If
'Copy & Paste ranges.  Checks if destination range is already occupied, and offsets to compensate.
    If headers = True Then
        sourceRange.Offset(1, 0).Resize(sourceRange.rows.Count - 1, sourceRange.columns.Count).Copy
    Else
        sourceRange.Copy
    End If
    If IsEmpty(destinationRange) = False Then
        destinationRange.End(xlDown).Offset(1, 0).PasteSpecial _
            Paste:=typePaste, operation:=operator, skipblanks:=skipBlankRows, transpose:=transposeRange
    Else
        destinationRange.PasteSpecial _
            Paste:=typePaste, operation:=operator, skipblanks:=skipBlankRows, transpose:=transposeRange
    End If
End Sub
