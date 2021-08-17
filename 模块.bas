Attribute VB_Name = "Ä£¿é1"
Sub ScratchMaco()
    Dim oRev As Revision
    For Each oRev In ActiveDocument.Revisions
        If oRev.Type = wdRevisionInsert Then
            oRev.Range.Font.Color = wdColorRed
            oRev.Range.Font.Shading.BackgroundPatternColor = wdColorYellow
        End If
        
        If oRev.Type = wdRevisionDelete Then
            'Texts = oRev.Range.Text
            Set Ranges = oRev.Range
            Ranges.InsertAfter Text:=oRev.Range.Text
            Ranges.Font.Color = wdColorGreen
            Ranges.Font.StrikeThrough = True
            Ranges.Font.Shading.BackgroundPatternColor = wdColorYellow
        End If
        oRev.Accept
    Next
End Sub

