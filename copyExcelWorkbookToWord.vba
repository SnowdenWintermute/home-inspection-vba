Sub CopyWorksheetsToWord()
' requires a reference to the Word Object library:
' in the VBE select Tools, References and check the Microsoft Word X.X object library

Dim wdApp As Word.Application, wdDoc As Word.Document, ws As Worksheet

    Application.ScreenUpdating = False
    Application.StatusBar = "Creating new document..."
    Set wdApp = New Word.Application
    Set wdDoc = wdApp.Documents.Add
    For Each ws In ActiveWorkbook.Worksheets
   
Application.StatusBar = "Copying data from " & ws.Name & "..."
        ws.UsedRange.Copy ' or edit to the range you want to copy
        wdDoc.Paragraphs(wdDoc.Paragraphs.Count).Range.InsertParagraphAfter
        wdDoc.Paragraphs(wdDoc.Paragraphs.Count).Range.Paste
        Application.CutCopyMode = False
        wdDoc.Paragraphs(wdDoc.Paragraphs.Count).Range.InsertParagraphAfter
        ' insert page break after all worksheets except the last one
        If Not ws.Name = Worksheets(Worksheets.Count).Name Then
            With wdDoc.Paragraphs(wdDoc.Paragraphs.Count).Range
                .InsertParagraphBefore
                .Collapse Direction:=wdCollapseEnd
                .InsertBreak Type:=wdPageBreak
            End With
        End If
    Next ws
   
Set ws = Nothing
    Application.StatusBar = "Cleaning up..."
    ' apply normal view
    With wdApp.ActiveWindow
        If .View.SplitSpecial = wdPaneNone Then
            .ActivePane.View.Type = wdNormalView
        Else
            .View.Type = wdNormalView
        End If
    End With
    Set wdDoc = Nothing
    wdApp.Visible = True
    Set wdApp = Nothing
    Application.StatusBar = False
   
End Sub