Attribute VB_Name = "CenterAndResetIndentForCaptions"
Sub CenterAndResetIndentForCaptions()
    Dim doc As Document
    Dim fld As Field
    Dim p As Paragraph
    Dim code As String
    Dim count As Integer
    
    Set doc = ActiveDocument
    count = 0
    
    ' 关闭屏幕更新，加速处理
    Application.ScreenUpdating = False
    
    For Each fld In doc.Fields
        ' 1. 检查是否为序列域 (wdFieldSequence)，这是题注的核心
        If fld.Type = wdFieldSequence Then
            code = Trim(fld.code.Text)
            
            ' 2. 检查域代码中是否包含 Table(表) 或 Figure(图) 关键字
            ' 通常 Word 默认是 "SEQ Table" 或 "SEQ Figure"
            ' 为了保险，同时也检查中文 "表" 和 "图"（防止自定义标签）
            If InStr(1, code, "Table", vbTextCompare) > 0 Or _
               InStr(1, code, "Figure", vbTextCompare) > 0 Or _
               InStr(1, code, "表", vbTextCompare) > 0 Or _
               InStr(1, code, "图", vbTextCompare) > 0 Then
               
                ' 获取该域所在的段落
                Set p = fld.Result.Paragraphs(1)
                
                With p.Format
                    ' --- 取消所有类型的缩进 ---
                    ' 磅值缩进归零
                    .LeftIndent = 0
                    .RightIndent = 0
                    .FirstLineIndent = 0
                    
                    ' 字符单位缩进归零 (针对中文版 Word 很重要)
                    .CharacterUnitLeftIndent = 0
                    .CharacterUnitRightIndent = 0
                    .CharacterUnitFirstLineIndent = 0
                    
                    ' --- 设置居中 ---
                    .Alignment = wdAlignParagraphCenter
                End With
                
                count = count + 1
            End If
        End If
    Next fld
    
    Application.ScreenUpdating = True
    MsgBox "处理完成！共调整了 " & count & " 处图表题注。"
End Sub

