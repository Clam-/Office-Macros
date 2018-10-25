Attribute VB_Name = "CitationProcessing"
Sub FootnoteConvert()
Dim lngIndex As Long
Dim oFN As Footnote
Dim oRng As Word.Range
  For lngIndex = ActiveDocument.Footnotes.Count To 1 Step -1
    Set oFN = ActiveDocument.Footnotes(lngIndex)
    oFN.Reference.InsertBefore " "
    oFN.Reference.InsertAfter ""
    Set oRng = oFN.Reference
    With oRng
      .Move wdCharacter, -1
      .FormattedText = oFN.Range.FormattedText
    End With
    oFN.Delete
  Next
lbl_Exit:
  Exit Sub
End Sub

Sub CitationCheckMove()
Dim oRng As Word.Range
Dim iBoxVal As Integer
Dim bModified As Boolean
Dim bPause As Boolean
Dim iPause As Integer

  iPause = MsgBox("Pause on every citation? (Otherwise only modified)", 36) ' 6 yes 7 no
  bPause = IIf(iPause = 6, True, False)

  Selection.EndKey Unit:=wdStory, Extend:=wdExtend
  Set tField = Selection.Range.Fields
  For Each fld In tField
    bModified = False
    If fld.Type = 81 Then
      fld.Select
      Set oRng = Selection.Range
      oRng.Start = oRng.Start - 1
      If Left(oRng.Text, 1) = "." Then
        oRng.Start = oRng.Start + 1
        oRng.Cut
        oRng.Start = oRng.Start - 1
        oRng.InsertBefore " "
        oRng.Collapse
        oRng.Start = oRng.Start + 1
        oRng.Paste
        ActiveWindow.SmallScroll Down:=5
        If (oRng.TextVisibleOnScreen <= 0) Then
          ActiveWindow.SmallScroll Up:=5
        End If
        iBoxVal = MsgBox("Moved Citation. Continue?", 49)
        bModified = True
        If iBoxVal = 2 Then
            End
        End If
      Else
        oRng.Start = oRng.Start - 1
        If Left(oRng.Text, 2) = ". " Then
          ' Odd case for . (citation).
          oRng.End = oRng.End + 1
          If Right(oRng.Text, 1) = "." Then
            oRng.End = oRng.End - 1
            oRng.Start = oRng.Start + 1
            oRng.Cut
            oRng.Delete 1
            oRng.Start = oRng.Start - 2
            oRng.Collapse
            oRng.Start = oRng.Start + 1
            oRng.Paste
          Else
            oRng.End = oRng.End - 1
            oRng.Start = oRng.Start + 1
            oRng.Cut
            oRng.Start = oRng.Start - 2
            oRng.Collapse
            oRng.Start = oRng.Start + 1
            oRng.Paste
          End If
          ActiveWindow.SmallScroll Down:=5
          If (oRng.TextVisibleOnScreen <= 0) Then
            ActiveWindow.SmallScroll Up:=5
          End If
          iBoxVal = MsgBox("Moved Citation. Continue?", 49)
          bModified = True
          If iBoxVal = 2 Then
            End
          End If
        End If
      End If
      If Not bModified And bPause Then
        ActiveWindow.SmallScroll Down:=5
        ActiveWindow.ScrollIntoView oRng
        iBoxVal = MsgBox("No preceeding period found. Continue? ", 33)
        If iBoxVal = 2 Then
          End
        End If
      End If
    End If
  Next
End Sub
