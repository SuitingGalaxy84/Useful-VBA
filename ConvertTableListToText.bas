Attribute VB_Name = "Module1"
Sub ConvertTableListToText()
    Dim para As Paragraph
    Dim ListStr As String
    
    For Each para In ActiveDocument.Content.Paragraphs
        If para.Range.ListFormat.ListType <> wdListNoNumbering Then
            ListStr = para.Range.ListFormat.ListString
            
            If InStr(ListStr, "表") > 0 Then
                para.Range.ListFormat.ConvertNumbersToText
            End If
        End If
    Next para
    
    MsgBox "错误表格序号转换为文字"
End Sub
