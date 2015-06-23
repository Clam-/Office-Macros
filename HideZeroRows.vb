Sub HideZeroRows()
    Dim zeros, nonzeros As Boolean
    
    For Each Row In ActiveSheet.UsedRange.Rows
        zeros = False
        nonzeros = False
        For Each cell In Row.Cells
            If Not IsEmpty(cell) And IsNumeric(cell) Then
                If cell.Value2 <> 0 Then
                    nonzeros = True
                ElseIf cell.Value2 = 0 Then
                    zeros = True
                End If
            End If
        Next
        If zeros And Not nonzeros Then
            Row.EntireRow.Hidden = True
        End If
    Next
End Sub
