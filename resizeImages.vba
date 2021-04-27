Sub ResizeImages()
Dim i As Long
    With ActiveDocument
     For i = 1 To .InlineShapes.Count
     If i > 3 Then
        With .InlineShapes(i)
            .Width = 267#
            .Height = 200
          End With
     End If
     Next i
     i = i - 2
            With .InlineShapes(i)
            .Width = 200#
            .Height = 75
          End With
          i = i + 1
          With .InlineShapes(i)
            .Width = 125#
            .Height = 125
          End With
End With
End Sub