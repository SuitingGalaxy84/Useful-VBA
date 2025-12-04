Attribute VB_Name = "ConvertFigureListToText"
Sub ConvertFigureListToText()
    Dim para As Paragraph
    Dim ListStr As String
    
    For Each para In ActiveDocument.Content.Paragraphs
        If para.Range.ListFormat.ListType <> wdListNoNumbering Then
            ListStr = para.Range.ListFormat.ListString
            
            If InStr(ListStr, "Í¼") > 0 Then
                para.Range.ListFormat.ConvertNumbersToText
            End If
        End If
    Next para
    
    MsgBox "´íÎóÍ¼Æ¬ÐòºÅ×ª»»ÎªÎÄ×Ö"
End Sub

