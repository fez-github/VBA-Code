Function OFFSETF(reference As Range, rows As Long, columns As Long) As Range
'Purpose: User-Defined Function that replicates Offset, but ignores filtered rows.
'Have attempted to use Resize, but has caused complications.  At the moment it cannot be done.

'Declare integers used to control looping
    Dim i As Long, j As Long, k As Long
'Check for filtered rows for Rows value.
    If rows <> 0 Then
        j = LoopStartingPoint(rows)
        i = j
        rows = LoopInitiator(i, j, reference, rows, True)
    End If
'Check for filtered columns for Columns value.
    If columns <> 0 Then
        k = LoopStartingPoint(columns)
        i = k
        columns = LoopInitiator(i, k, reference, columns, False)
    End If
'Use adjusted values to create range that determines new offset position.
    Set OFFSETF = reference.Offset(rows, columns)
End Function
Function LoopStartingPoint(value As Long) As Integer
'Determine what value and direction to loop in, depending on whether the value is positive or negative.
    If value >= 1 Then
        LoopStartingPoint = 1
    ElseIf value <= -1 Then
        LoopStartingPoint = -1
    End If
End Function

Function LoopInitiator(loopPosition As Long, loopDirection As Long, reference As Range, field As Long, row As Boolean) As Long
'Purpose: Increase the size of the original offset argument value in order to accommodate for the filtered cells.

If loopDirection > 0 Then
    Do While loopPosition <= field
        If row = True Then
            If reference.Offset(i, 0).EntireRow.Hidden = True Then
                field = field + loopDirection
            End If
        Else
            If reference.Offset(0, i).EntireColumn.Hidden = True Then
                field = field + loopDirection
            End If
        End If
        loopPosition = loopPosition + loopDirection
    Loop
Else
    Do While i >= field
        If row = False Then
            If reference.Offset(i, 0).EntireRow.Hidden = True Then 'Procedure ended here because it was looping left, and went past row 1.
                field = field + loopDirection
            End If
        Else
            If reference.Offset(0, i).EntireColumn.Hidden = True Then
                field = field + loopDirection
            End If
        End If
        loopPosition = loopPosition + loopDirection
    Loop
End If
    LoopInitiator = field
End Function
