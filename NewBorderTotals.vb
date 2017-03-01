
Sub NewBorders()
     ' xlEdgeRight xlContinuous
    Dim hasTotal, padding As Boolean
    Dim cellStart, cellEnd As String
    Dim bordStart, boardEnd As String
    Dim borderRange As Range
    Dim r As Integer
    
    For Each Row In ActiveSheet.UsedRange.Rows
        hasTotal = False
        bordStart = ""
        bordEnd = ""
        padding = False
        For Each cell In Row.Cells
            ' detect existing borders
            If cell.Borders(xlEdgeRight).LineStyle = xlContinuous Then
                If bordStart <> "" Then
                    bordEnd = cell.Address
                End If
            End If
            If cell.Borders(xlEdgeLeft).LineStyle = xlContinuous Then
                If bordStart = "" Then
                    bordStart = cell.Address
                End If
            End If
            ' Detect cell contents
            If Not IsEmpty(cell) And Not IsError(cell) Then
                If cell.Value2 Like "*Total*" Then
                    hasTotal = True
                    cellStart = cell.Address
                    cellEnd = cell.Address
                    If Not (cell.Value2 Like "Total*") Then
                        padding = True
                    End If
                End If
            End If
            If hasTotal Then
                If IsEmpty(cell) Then
                    ' pass
                Else
                    cellEnd = cell.Address
                End If
            End If
        Next
        
        If hasTotal Then
            If bordEnd <> "" Then
                cellStart = bordStart
                cellEnd = bordEnd
            End If
            Set borderRange = ActiveSheet.Range(cellStart, cellEnd)
            borderRange.Borders.LineStyle = xlContinuous
            If padding Then
                borderRange.Borders.Weight = xlThin
            Else
                borderRange.Borders.Weight = xlMedium
            End If
            borderRange.Borders.Color = RGB(0, 0, 0)
            borderRange.Borders(xlInsideVertical).LineStyle = xlLineStyleNone
        End If
    Next
End Sub
