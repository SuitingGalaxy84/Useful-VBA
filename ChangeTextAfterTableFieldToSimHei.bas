Attribute VB_Name = "ChangeTextAfterTableFieldToSimHei"
Sub ChangeTextAfterTableFieldToSimHei()
    Dim doc As Document
    Dim fld As Field
    Dim rng As Range
    Dim fieldCode As String
    
    Set doc = ActiveDocument
    
    For Each fld In doc.Fields
        
        fieldCode = Trim(fld.code.Text)
        
        If fld.Type = wdFieldSequence And InStr(1, fieldCode, "Table", vbTextCompare) > 0 Then
            Set rng = doc.Range(Start:=fld.Result.End, End:=fld.Result.Paragraphs(1).Range.End - 1)
            
            If rng.Start < rng.End Then
                rng.Font.NameFarEast = "黑体"
                rng.Font.NameAscii = "Times New Roman"
            End If
            
        End If
    
    Next fld
    
    MsgBox "表格Caption正则化处理完成"
    
End Sub
