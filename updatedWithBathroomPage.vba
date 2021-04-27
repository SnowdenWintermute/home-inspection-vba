Sub ProcessScriptTable()
' Normal Vars
  Dim oTbl As Table
  Dim lastCellText As String
  Dim currentHeader As String
  Dim currentPage As String
  Dim headerHasBeenPrintedToDocument As Boolean
                      ' Col 8
                    Dim bathroomALoc As String
                    Dim bathroomAToilet As String
                    Dim bathroomASink As String
                    Dim bathroomAShower As String
                    Dim bathroomABathtub As String
                    Dim bathroomAVentLight As String
                    Dim bathroomAComments As String
                    ' Col 10
                    Dim bathroomBLoc As String
                    Dim bathroomBToilet As String
                    Dim bathroomBSink As String
                    Dim bathroomBShower As String
                    Dim bathroomBBathtub As String
                    Dim bathroomBVentLight As String
                    Dim bathroomBComments As String
                    ' Col 12
                    Dim bathroomCLoc As String
                    Dim bathroomCToilet As String
                    Dim bathroomCSink As String
                    Dim bathroomCShower As String
                    Dim bathroomCBathtub As String
                    Dim bathroomCVentLight As String
                    Dim bathroomCComments As String
                    ' Col 14
                    Dim bathroomDLoc As String
                    Dim bathroomDToilet As String
                    Dim bathroomDSink As String
                    Dim bathroomDShower As String
                    Dim bathroomDBathtub As String
                    Dim bathroomDVentLight As String
                    Dim bathroomDComments As String
                    ' Col 16
                    Dim bathroomELoc As String
                    Dim bathroomEToilet As String
                    Dim bathroomESink As String
                    Dim bathroomEShower As String
                    Dim bathroomEBathtub As String
                    Dim bathroomEVentLight As String
                    Dim bathroomEComments As String
                    ' Col 18
                    Dim bathroomFLoc As String
                    Dim bathroomFToilet As String
                    Dim bathroomFSink As String
                    Dim bathroomFShower As String
                    Dim bathroomFBathtub As String
                    Dim bathroomFVentLight As String
                    Dim bathroomFComments As String
  Debug.Print "START"
' End Vars
  For Each oTable In ActiveDocument.Tables
    For Each oRow In oTable.Rows
        For Each ocell In oRow.Cells
            If ocell.Range.Font.Size = 24 Then
                currentPage = ocell.Range.Text
                ActiveDocument.Content.InsertAfter Text:="Page: " & currentPage
                ' Electrical Page
                If InStr(currentPage, "ELECTRICAL INSPECTION REPORT") > 0 Then
                    Dim lastCheckedBox As Boolean
                    For Each eRow In oTable.Rows
                        For Each eCell In eRow.Cells
                            If InStr(eCell.Range.Text, "TRUE") > 0 Then
                                lastCheckedBox = True
                            ElseIf InStr(eCell.Range.Text, "FALSE") > 0 Then
                                lastCheckedBox = False
                            End If
                            If eCell.Range.Font.Size = 11 Then
                                Debug.Print eCell.Range.Text
                            End If
                        Next
                    Next
                End If
                
                ' Bathroom Page
                If InStr(currentPage, "BATHROOM INSPECTION REPORT") > 0 Then
                    Dim currentBathroomHeader As String
                    Dim currentBathroomHeaderPrinted As Boolean
                    
                    Dim currentRow As Integer
                    currentRow = 0
                    Dim currentCol As Integer
                    currentCol = 0
                    For Each bathroomRow In oTable.Rows
                        currentRow = currentRow + 1
                        currentCol = 0
                        ' Location names
                        For Each bathroomCell In bathroomRow.Cells
                            currentCol = currentCol + 1
                            
                            If currentRow = 8 Then
                                If currentCol = 3 Then
                                    If Len(bathroomCell.Range.Text) > 0 Then
                                        bathroomALoc = bathroomCell.Range.Text
                                    End If
                                End If
                            End If
                            If currentRow = 8 Then
                                If currentCol = 5 Then
                                    If Len(bathroomCell.Range.Text) > 0 Then
                                        bathroomBLoc = bathroomCell.Range.Text
                                    End If
                                End If
                            End If
                            If currentRow = 8 Then
                                If currentCol = 7 Then
                                    If Len(bathroomCell.Range.Text) > 0 Then
                                        bathroomCLoc = bathroomCell.Range.Text
                                    End If
                                End If
                            End If
                            If currentRow = 9 Then
                                If currentCol = 3 Then
                                    If Len(bathroomCell.Range.Text) > 0 Then
                                        bathroomDLoc = bathroomCell.Range.Text
                                    End If
                                End If
                            End If
                            If currentRow = 9 Then
                                If currentCol = 5 Then
                                    If Len(bathroomCell.Range.Text) > 0 Then
                                        bathroomELoc = bathroomCell.Range.Text
                                    End If
                                End If
                            End If
                            If currentRow = 9 Then
                                If currentCol = 7 Then
                                    If Len(bathroomCell.Range.Text) > 0 Then
                                        bathroomFLoc = bathroomCell.Range.Text
                                    End If
                                End If
                            End If
                            ' Comments
                            If bathroomCell.Range.Font.Size = 10 Then
                                If currentRow = 57 And Len(bathroomCell.Range.Text) > 2 Then
                                    bathroomAComments = bathroomAComments & bathroomCell.Range.Text
                                End If
                                If currentRow = 58 And Len(bathroomCell.Range.Text) > 2 Then
                                    bathroomBComments = bathroomBComments & bathroomCell.Range.Text
                                End If
                                If currentRow = 59 And Len(bathroomCell.Range.Text) > 2 Then
                                    bathroomCComments = bathroomCComments & bathroomCell.Range.Text
                                End If
                                If currentRow = 60 And Len(bathroomCell.Range.Text) > 2 Then
                                    bathroomDComments = bathroomDComments & bathroomCell.Range.Text
                                End If
                                If currentRow = 61 And Len(bathroomCell.Range.Text) > 2 Then
                                    bathroomEComments = bathroomEComments & bathroomCell.Range.Text
                                End If
                                If currentRow = 62 And Len(bathroomCell.Range.Text) > 2 Then
                                    bathroomFComments = bathroomFComments & bathroomCell.Range.Text
                                End If
                            End If
                            If bathroomCell.Range.Font.Size = 12 Then
                                currentBathroomHeader = bathroomCell.Range.Text
                            End If
                            ' Checked items
                            ' Change 0 to 255 for red
                            If bathroomCell.Range.Font.Color = 0 And InStr(bathroomCell.Range.Text, "TRUE") Then
                                ' Toilet: Row 15 to 18
                                ' Sink: Row 23 to 26
                                ' Shower: Row 32 to 37
                                ' Bathtub: Row 43 to 48
                                ' Vent/Light: Row 53 to 54
                                ' Debug.Print "Col: " & currentCol & " Row:" & currentRow
                                If currentCol = 8 Or currentCol = 7 Then
                                    If currentRow < 19 And currentRow > 14 Then
                                        bathroomAToilet = bathroomAToilet & bathroomRow.Cells(2).Range.Text
                                        
                                    End If
                                    If currentRow < 27 And currentRow > 22 Then
                                        Dim newValue As String
                                        newValue = bathroomASink & bathroomRow.Cells(2).Range.Text
                                        bathroomASink = newValue
                                    End If
                                    If currentRow < 39 And currentRow > 31 Then
                                        bathroomAShower = bathroomAShower & bathroomRow.Cells(2).Range.Text
                                    End If
                                    If currentRow < 50 And currentRow > 42 Then
                                        bathroomABathtub = bathroomABathtub & bathroomRow.Cells(2).Range.Text
                                    End If
                                    If currentRow < 56 And currentRow > 52 Then
                                        bathroomAVentLight = bathroomAVentLight & bathroomRow.Cells(2).Range.Text
                                    End If
                                End If
                                If currentCol = 10 Or currentCol = 9 Then
                                    If currentRow < 19 And currentRow > 14 Then
                                        bathroomBToilet = bathroomBToilet & bathroomRow.Cells(2).Range.Text
                                    End If
                                    If currentRow < 27 And currentRow > 22 Then
                                        bathroomBSink = bathroomBSink & bathroomRow.Cells(2).Range.Text
                                    End If
                                    If currentRow < 38 And currentRow > 31 Then
                                        bathroomBShower = bathroomBShower & bathroomRow.Cells(2).Range.Text
                                    End If
                                    If currentRow < 49 And currentRow > 42 Then
                                        bathroomBBathtub = bathroomBBathtub & bathroomRow.Cells(2).Range.Text
                                    End If
                                    If currentRow < 55 And currentRow > 52 Then
                                        bathroomBVentLight = bathroomBVentLight & bathroomRow.Cells(2).Range.Text
                                    End If
                                End If
                                If currentCol = 12 Or currentCol = 11 Then
                                    If currentRow < 19 And currentRow > 14 Then
                                        bathroomCToilet = bathroomCToilet & bathroomRow.Cells(2).Range.Text
                                    End If
                                    If currentRow < 27 And currentRow > 22 Then
                                        bathroomCSink = bathroomCSink & bathroomRow.Cells(2).Range.Text
                                    End If
                                    If currentRow < 38 And currentRow > 31 Then
                                        bathroomCShower = bathroomCShower & bathroomRow.Cells(2).Range.Text
                                    End If
                                    If currentRow < 49 And currentRow > 42 Then
                                        bathroomCBathtub = bathroomCBathtub & bathroomRow.Cells(2).Range.Text
                                    End If
                                    If currentRow < 55 And currentRow > 52 Then
                                        bathroomCVentLight = bathroomCVentLight & bathroomRow.Cells(2).Range.Text
                                    End If
                                End If
                                If currentCol = 14 Or currentCol = 13 Then
                                    If currentRow < 19 And currentRow > 14 Then
                                        bathroomDToilet = bathroomDToilet & bathroomRow.Cells(2).Range.Text
                                    End If
                                    If currentRow < 27 And currentRow > 22 Then
                                        bathroomDSink = bathroomDSink & bathroomRow.Cells(2).Range.Text
                                    End If
                                    If currentRow < 38 And currentRow > 31 Then
                                        bathroomDShower = bathroomDShower & bathroomRow.Cells(2).Range.Text
                                    End If
                                    If currentRow < 49 And currentRow > 42 Then
                                        bathroomDBathtub = bathroomDBathtub & bathroomRow.Cells(2).Range.Text
                                    End If
                                    If currentRow < 55 And currentRow > 52 Then
                                        bathroomDVentLight = bathroomDVentLight & bathroomRow.Cells(2).Range.Text
                                    End If
                                End If
                                If currentCol = 16 Or currentCol = 15 Then
                                    If currentRow < 19 And currentRow > 14 Then
                                        bathroomEToilet = bathroomEToilet & bathroomRow.Cells(2).Range.Text
                                    End If
                                    If currentRow < 27 And currentRow > 22 Then
                                        bathroomESink = bathroomESink & bathroomRow.Cells(2).Range.Text
                                    End If
                                    If currentRow < 38 And currentRow > 31 Then
                                        bathroomEShower = bathroomEShower & bathroomRow.Cells(2).Range.Text
                                    End If
                                    If currentRow < 49 And currentRow > 42 Then
                                        bathroomEBathtub = bathroomEBathtub & bathroomRow.Cells(2).Range.Text
                                    End If
                                    If currentRow < 55 And currentRow > 52 Then
                                        bathroomEVentLight = bathroomEVentLight & bathroomRow.Cells(2).Range.Text
                                    End If
                                End If
                                If currentCol = 18 Or currentCol = 17 Then
                                    If currentRow < 19 And currentRow > 14 Then
                                        bathroomFToilet = bathroomFToilet & bathroomRow.Cells(2).Range.Text
                                    End If
                                    If currentRow < 27 And currentRow > 22 Then
                                        bathroomFSink = bathroomFSink & bathroomRow.Cells(2).Range.Text
                                    End If
                                    If currentRow < 38 And currentRow > 31 Then
                                        bathroomFShower = bathroomFShower & bathroomRow.Cells(2).Range.Text
                                    End If
                                    If currentRow < 49 And currentRow > 42 Then
                                        bathroomFBathtub = bathroomFBathtub & bathroomRow.Cells(2).Range.Text
                                    End If
                                    If currentRow < 55 And currentRow > 52 Then
                                        bathroomFVentLight = bathroomFVentLight & bathroomRow.Cells(2).Range.Text
                                    End If
                                End If
                            End If
                        Next
                    Next
                    If Len(bathroomALoc) > 0 Then
                        ActiveDocument.Content.InsertAfter Text:="LOCATION: " & bathroomALoc
                        If Len(bathroomAToilet) > 0 Then
                            ActiveDocument.Content.InsertAfter Text:="TOILET" & vbNewLine
                            ActiveDocument.Content.InsertAfter Text:=bathroomAToilet
                        End If
                        If Len(bathroomASink) > 0 Then
                            ActiveDocument.Content.InsertAfter Text:="SINK" & vbNewLine
                            ActiveDocument.Content.InsertAfter Text:=bathroomASink
                        End If
                        If Len(bathroomAShower) > 0 Then
                            ActiveDocument.Content.InsertAfter Text:="SHOWER" & vbNewLine
                            ActiveDocument.Content.InsertAfter Text:=bathroomAShower
                        End If
                        If Len(bathroomABathtub) > 0 Then
                            ActiveDocument.Content.InsertAfter Text:="BATHTUB" & vbNewLine
                            ActiveDocument.Content.InsertAfter Text:=bathroomABathtub
                        End If
                        If Len(bathroomAVentLight) > 0 Then
                            ActiveDocument.Content.InsertAfter Text:="VENT / LIGHT" & vbNewLine
                            ActiveDocument.Content.InsertAfter Text:=bathroomAVentLight
                        End If
                        If Len(bathroomAComments) > 0 Then
                            ActiveDocument.Content.InsertAfter Text:="COMMENTS" & vbNewLine
                            ActiveDocument.Content.InsertAfter Text:=bathroomAComments
                        End If
                    End If
                    If Len(bathroomBLoc) > 0 Then
                        ActiveDocument.Content.InsertAfter Text:="LOCATION: " & bathroomBLoc
                        If Len(bathroomBToilet) > 0 Then
                            ActiveDocument.Content.InsertAfter Text:="TOILET" & vbNewLine
                            ActiveDocument.Content.InsertAfter Text:=bathroomBToilet
                        End If
                        If Len(bathroomBSink) > 0 Then
                            ActiveDocument.Content.InsertAfter Text:="SINK" & vbNewLine
                            ActiveDocument.Content.InsertAfter Text:=bathroomBSink
                        End If
                        If Len(bathroomBShower) > 0 Then
                            ActiveDocument.Content.InsertAfter Text:="SHOWER" & vbNewLine
                            ActiveDocument.Content.InsertAfter Text:=bathroomBShower
                        End If
                        If Len(bathroomBBathtub) > 0 Then
                            ActiveDocument.Content.InsertAfter Text:="BATHTUB" & vbNewLine
                            ActiveDocument.Content.InsertAfter Text:=bathroomABathtub
                        End If
                        If Len(bathroomBVentLight) > 0 Then
                            ActiveDocument.Content.InsertAfter Text:="VENT / LIGHT" & vbNewLine
                            ActiveDocument.Content.InsertAfter Text:=bathroomBVentLight
                        End If
                        If Len(bathroomBComments) > 0 Then
                            ActiveDocument.Content.InsertAfter Text:="COMMENTS" & vbNewLine
                            ActiveDocument.Content.InsertAfter Text:=bathroomBComments
                        End If
                    End If
                    If Len(bathroomCLoc) > 0 Then
                        ActiveDocument.Content.InsertAfter Text:="LOCATION: " & bathroomCLoc
                        If Len(bathroomCToilet) > 0 Then
                            ActiveDocument.Content.InsertAfter Text:="TOILET" & vbNewLine
                            ActiveDocument.Content.InsertAfter Text:=bathroomCToilet
                        End If
                        If Len(bathroomCSink) > 0 Then
                            ActiveDocument.Content.InsertAfter Text:="SINK" & vbNewLine
                            ActiveDocument.Content.InsertAfter Text:=bathroomCSink
                        End If
                        If Len(bathroomCShower) > 0 Then
                            ActiveDocument.Content.InsertAfter Text:="SHOWER" & vbNewLine
                            ActiveDocument.Content.InsertAfter Text:=bathroomCShower
                        End If
                        If Len(bathroomCBathtub) > 0 Then
                            ActiveDocument.Content.InsertAfter Text:="BATHTUB" & vbNewLine
                            ActiveDocument.Content.InsertAfter Text:=bathroomCBathtub
                        End If
                        If Len(bathroomCVentLight) > 0 Then
                            ActiveDocument.Content.InsertAfter Text:="VENT / LIGHT" & vbNewLine
                            ActiveDocument.Content.InsertAfter Text:=bathroomCVentLight
                        End If
                        If Len(bathroomCComments) > 0 Then
                            ActiveDocument.Content.InsertAfter Text:="COMMENTS" & vbNewLine
                            ActiveDocument.Content.InsertAfter Text:=bathroomCComments
                        End If
                    End If
                    If Len(bathroomDLoc) > 0 Then
                        ActiveDocument.Content.InsertAfter Text:="LOCATION: " & bathroomDLoc
                        If Len(bathroomDToilet) > 0 Then
                            ActiveDocument.Content.InsertAfter Text:="TOILET" & vbNewLine
                            ActiveDocument.Content.InsertAfter Text:=bathroomDToilet
                        End If
                        If Len(bathroomDSink) > 0 Then
                            ActiveDocument.Content.InsertAfter Text:="SINK" & vbNewLine
                            ActiveDocument.Content.InsertAfter Text:=bathroomDSink
                        End If
                        If Len(bathroomDShower) > 0 Then
                            ActiveDocument.Content.InsertAfter Text:="SHOWER" & vbNewLine
                            ActiveDocument.Content.InsertAfter Text:=bathroomDShower
                        End If
                        If Len(bathroomDBathtub) > 0 Then
                            ActiveDocument.Content.InsertAfter Text:="BATHTUB" & vbNewLine
                            ActiveDocument.Content.InsertAfter Text:=bathroomDBathtub
                        End If
                        If Len(bathroomDVentLight) > 0 Then
                            ActiveDocument.Content.InsertAfter Text:="VENT / LIGHT" & vbNewLine
                            ActiveDocument.Content.InsertAfter Text:=bathroomDVentLight
                        End If
                        If Len(bathroomDComments) > 0 Then
                            ActiveDocument.Content.InsertAfter Text:="COMMENTS" & vbNewLine
                            ActiveDocument.Content.InsertAfter Text:=bathroomDComments
                        End If
                    End If
                    If Len(bathroomELoc) > 0 Then
                        ActiveDocument.Content.InsertAfter Text:="LOCATION: " & bathroomELoc
                        If Len(bathroomEToilet) > 0 Then
                            ActiveDocument.Content.InsertAfter Text:="TOILET" & vbNewLine
                            ActiveDocument.Content.InsertAfter Text:=bathroomEToilet
                        End If
                        If Len(bathroomESink) > 0 Then
                            ActiveDocument.Content.InsertAfter Text:="SINK" & vbNewLine
                            ActiveDocument.Content.InsertAfter Text:=bathroomESink
                        End If
                        If Len(bathroomEShower) > 0 Then
                            ActiveDocument.Content.InsertAfter Text:="SHOWER" & vbNewLine
                            ActiveDocument.Content.InsertAfter Text:=bathroomEShower
                        End If
                        If Len(bathroomEBathtub) > 0 Then
                            ActiveDocument.Content.InsertAfter Text:="BATHTUB" & vbNewLine
                            ActiveDocument.Content.InsertAfter Text:=bathroomEBathtub
                        End If
                        If Len(bathroomEVentLight) > 0 Then
                            ActiveDocument.Content.InsertAfter Text:="VENT / LIGHT" & vbNewLine
                            ActiveDocument.Content.InsertAfter Text:=bathroomEVentLight
                        End If
                        If Len(bathroomEComments) > 0 Then
                            ActiveDocument.Content.InsertAfter Text:="COMMENTS" & vbNewLine
                            ActiveDocument.Content.InsertAfter Text:=bathroomEComments
                        End If
                    End If
                    If Len(bathroomFLoc) > 0 Then
                        ActiveDocument.Content.InsertAfter Text:="LOCATION: " & bathroomFLoc
                        If Len(bathroomFToilet) > 0 Then
                            ActiveDocument.Content.InsertAfter Text:="TOILET" & vbNewLine
                            ActiveDocument.Content.InsertAfter Text:=bathroomFToilet
                        End If
                        If Len(bathroomFSink) > 0 Then
                            ActiveDocument.Content.InsertAfter Text:="SINK" & vbNewLine
                            ActiveDocument.Content.InsertAfter Text:=bathroomFSink
                        End If
                        If Len(bathroomFShower) > 0 Then
                            ActiveDocument.Content.InsertAfter Text:="SHOWER" & vbNewLine
                            ActiveDocument.Content.InsertAfter Text:=bathroomFShower
                        End If
                        If Len(bathroomFBathtub) > 0 Then
                            ActiveDocument.Content.InsertAfter Text:="BATHTUB" & vbNewLine
                            ActiveDocument.Content.InsertAfter Text:=bathroomFBathtub
                        End If
                        If Len(bathroomFVentLight) > 0 Then
                            ActiveDocument.Content.InsertAfter Text:="VENT / LIGHT" & vbNewLine
                            ActiveDocument.Content.InsertAfter Text:=bathroomFVentLight
                        End If
                        If Len(bathroomFComments) > 0 Then
                            ActiveDocument.Content.InsertAfter Text:="COMMENTS" & vbNewLine
                            ActiveDocument.Content.InsertAfter Text:=bathroomFComments
                        End If
                    End If
                End If
                ' End Bathroom Page
            End If
            If ocell.Range.Font.Size = 12 Then
             currentHeader = ocell.Range.Text
             headerHasBeenPrintedToDocument = False
            End If
            If ocell.Range.Font.Color = 255 Or ocell.Range.Font.Color = 12611584 Then
            If InStr(lastCellText, "TRUE") > 0 Then
                If Not headerHasBeenPrintedToDocument Then
                    ActiveDocument.Content.InsertAfter Text:=currentHeader
                    headerHasBeenPrintedToDocument = True
                End If
                ActiveDocument.Content.InsertAfter Text:=ocell.Range.Text
            End If
         lastCellText = ocell.Range.Text
         End If
         If ocell.Range.Font.Size = 10 And InStr(currentPage, "BATHROOM INSPECTION REPORT") = 0 Then
            If Len(ocell.Range.Text) > 3 Then
                    If Not InStr(ocell.Range.Text, "Comments:") > 0 Then
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


