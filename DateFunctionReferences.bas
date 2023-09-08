'Timezone differences:
'Can be calculated for by subtracting a fraction from a date/time field.
   'Range("A2").Value - 5/24  'This subtracts 5 hours from the date/time field
    'A2's value would equal 06/01/2020 20:00PM

'Pull date from generic text
'Can use a formula to identify the proper m/d/y values

    '"=DATEVALUE(MID(LEFT(D2,10),6,2)&""-""&MID(LEFT(D2,10),9,2)&""-""&MID(LEFT(D2,10),1,4))")
    
    'Or use Text to Columns, delimit properly and remove any additional columns.

'Autofilter to only show past certain date
   ' .Range("$A$1:$L$" & lastRow).AutoFilter field:=5, _
    Operator:=xlFilterValues, Criteria1:=">" & deadline.  deadline = 01/01/2021
    
Sub MonthAdjustment()
'Correct values that are not from this month.  Latest dates could be from later than the right month, and need to be adjusted to the last day of this month.
'Dates were in yyyy-mm-dd format.
'Not a complete procedure.  Merely used as reference material.
    LastRow = .Range("A" & rows.Count).End(xlUp).row
    .Range("A1:AH" & LastRow).Sort Key1:=.Range("E1:E" & LastRow), order1:=xlDescending, Header:=xlYes 'Sort range of dates from latest to oldest.
    ActVal = .Range("E2").value 'Acquire latest date.
    FirstDay = .Range("E" & LastRow) 'Acquire earliest date
    i = 0
    If Mid(ActVal, 6, 2) <> Mid(FirstDay, 6, 2) Then 'Check if the month of the latest date matches the month of the earliest date.
        Do Until Mid(ActVal, 6, 2) = Mid(FirstDay, 6, 2) 'Loop through range until months match.
            i = i + 1
            ActVal = .Range("E2").Offset(i, 0)
        Loop
    currentRow = .Range("E2").Offset(i, 0).row 'This is the last date of the desired month.  Declare it as variable.
    .Range("E" & currentRow).AutoFill Destination:=.Range("E" & currentRow & ":E2"), Type:=xlFillCopy 'Autofill the last date upward to replace the dates that are too far ahead.
End Sub
