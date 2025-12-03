Attribute VB_Name = "Module4"
Sub ChangeTextAfterFigureFieldToSimHei()
    Dim doc As Document
    Dim fld As Field
    Dim rng As Range
    Dim fieldCode As String
    
    Set doc = ActiveDocument
    
    For Each fld In doc.Fields
        
        fieldCode = Trim(fld.code.Text)
        
        If fld.Type = wdFieldSequence And InStr(1, fieldCode, "Figure", vbTextCompare) > 0 Then
            Set rng = doc.Range(Start:=fld.Result.End, End:=fld.Result.Paragraphs(1).Range.End - 1)
            
            If rng.Start < rng.End Then
                rng.Font.NameFarEast = "黑体"
                rng.Font.NameAscii = "Times New Roman"
                rng.Font.Size = 12
            End If
            
        End If
    
    Next fld
    
    MsgBox "图片Caption字体正则化处理完成"
    
End Sub

