Sub ProcessScriptTable()
  Dim oTbl As Table
  Dim lastCellText As String
  Dim currentHeader As String
  Dim currentPage As String
  Dim headerHasBeenPrintedToDocument As Boolean
  Debug.Print "START"
  For Each oTable In ActiveDocument.Tables
    For Each oRow In oTable.Rows
        For Each ocell In oRow.Cells
            If ocell.Range.Font.Size = 24 Then
                currentPage = ocell.Range.Text
                ActiveDocument.Content.InsertAfter Text:="Page: " & currentPage
            End If
            If ocell.Range.Font.Size = 12 Then
             currentHeader = ocell.Range.Text
             headerHasBeenPrintedToDocument = False
            End If
            If ocell.Range.Font.Color = 255 Or ocell.Range.Font.Color = 13395456 Then
            If InStr(lastCellText, "TRUE") > 0 Then
                Debug.Print ocell.Range.Text
                Debug.Print ocell.Range.Font.Color
                If Not headerHasBeenPrintedToDocument Then
                    ActiveDocument.Content.InsertAfter Text:=currentHeader
                    headerHasBeenPrintedToDocument = True
                End If
                ActiveDocument.Content.InsertAfter Text:=ocell.Range.Text
            End If
         lastCellText = ocell.Range.Text
         End If
         If ocell.Range.Font.Size = 10 Then
            If Len(ocell.Range.Text) > 3 Then
                    If Not InStr(ocell.Range.Text, "Comments:") > 0 Then
                        Debug.Print Len(ocell.Range.Text)
                        Debug.Print ocell.Range.Text
                        ActiveDocument.Content.InsertAfter Text:="Comment: " & ocell.Range.Text
                        If Not headerHasBeenPrintedToDocument Then
                            ActiveDocument.Content.InsertAfter Text:=currentHeader
                            headerHasBeenPrintedToDocument = True
                        End If
                    End If
            End If
         End If
        Next
    Next
  Next
End Sub
