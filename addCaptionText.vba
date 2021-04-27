Sub AddCaptionText1()
Dim i As Long
   With ActiveDocument
     For i = 1 To .InlineShapes.Count
         If i > 3 Then
            If i <> .InlineShapes.Count Then
               If i Mod 2 = 1 Then
                  With .InlineShapes(i)
                    .Range.InsertAfter Chr(13)
                    .Range.InsertAfter Chr(13)
                    .Range.InsertAfter Chr(13)
                  End With
               End If
            End If
         End If
      Next i
   End With
End Sub