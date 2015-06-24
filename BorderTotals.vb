Sub BorderTotals()
    Dim hasTotal As Boolean
    Dim cellStart, cellEnd As String
    Dim borderRange As Range

    For Each Row In ActiveSheet.UsedRange.Rows
        hasTotal = False
        For Each cell In Row.Cells
            If Not IsEmpty(cell) Then
                If cell.Value2 Like "*Total*" Then
                    hasTotal = True
                    cellStart = cell.Address
                    cellEnd = cell.Address
                End If
            End If
            If hasTotal Then
                If IsEmpty(cell) Then
                    Exit For
                Else
                    cellEnd = cell.Address
                End If
            End If
        Next
        If hasTotal Then
            Set borderRange = ActiveSheet.Range(cellStart, cellEnd)
            borderRange.Borders.LineStyle = xlContinuous
            borderRange.Borders.Weight = xlMedium
			borderRange.Borders.Color = RGB(0, 0, 0)
            borderRange.Borders(xlInsideVertical).LineStyle = xlLineStyleNone
        End If
    Next
End Sub