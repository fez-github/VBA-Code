'Enum values are used to assist user in selecting the Autofill type.
    Public Enum AutofillType
        xlFillDefault = 0
        xlFillCopy = 1
        xlFillSeries = 2
        xlFillFormats = 3
        xlFillValues = 4
        xlFillDays = 5
        xlFillWeekdays = 6
        xlFillMonths = 7
        xlFillYears = 8
        xlLinearTrend = 9
        xlGrowthTrend = 10
        xlFlashFill = 11
    End Enum
    
Public Sub AutoFillNewColumn(ColumnHeader As Range, ColumnTitle As String, ByVal Formula As String, Optional fillType As AutofillType)
'Purpose: To quickly populate a column with a header and common formula.
'Written by: Mark Hansen
'Last Updated: July 13, 2021

'Declare variables.  Used for acquiring specific workbook/worksheet information for later.
    Dim wbName As String, wsName As String
        wbName = ColumnHeader.Worksheet.Parent.name
        wsName = ColumnHeader.Worksheet.name
'Set column header.
    ColumnHeader.value = ColumnTitle
        'Note that the ColumnHeader range must include the sheet name, and workbook name if needed.
'Input formula in 2nd cell of column.
    ColumnHeader.Offset(1, 0).value = Formula
'Autofill formula.  Looks complicated, but this is all designed so it can work regardless of what workbook or worksheet it is being used on.
With Workbooks(wbName).Sheets(wsName)
    If Application.WorksheetFunction.CountA(.rows(3)) > 0 Then
        ColumnHeader.Offset(1, 0).AutoFill _
            Destination:=.Range(ColumnHeader.Offset(1, 0), .Cells(.Cells.Find _
            (What:="*", After:=.Cells(1), LookAt:=xlPart, LookIn:=xlFormulas, SearchOrder:=xlByRows, SearchDirection:=xlPrevious, MatchCase:=False).row, _
            ColumnHeader.Column)).SpecialCells(xlCellTypeVisible), Type:=fillType
    End If
End With
End Sub

Public Sub AutoFillFilteredColumn(ColumnHeader As Range, ColumnTitle As String, ByVal Formula As String)
'Purpose: To quickly populate a column with a header and common formula, assuming the rest of the table was filtered.
'Written by: Mark Hansen
'NOTE: A proper Autofill will not work with filters enabled.
'Last Updated: July 13, 2021

'Declare variables
    Dim wbName As String, wsName As String
    wbName = ColumnHeader.Worksheet.Parent.name
    wsName = ColumnHeader.Worksheet.name
'Set column header.
    ColumnHeader.value = ColumnTitle
'Autofill formula.  Designed to work regardless of what workbook or worksheet it is being used on.
With Workbooks(wbName).Sheets(wsName)
    If Application.WorksheetFunction.CountA(.rowsName(3)) > 0 Then
        .Range(ColumnHeader.Offset(1, 0), _
        .Cells(.Cells.Find(What:="*", After:=.Cells(1), LookAt:=xlPart, LookIn:=xlFormulas, SearchOrder:=xlByRowsName, SearchDirection:=xlPrevious, MatchCase:=False).row, _
        ColumnHeader.Column)).SpecialCells(xlCellTypeVisible).FormulaR1C1 = Formula
    End If
End With
End Sub
