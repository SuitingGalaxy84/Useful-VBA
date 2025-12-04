Attribute VB_Name = "ForceAllHeadingsToHeiti_ByScan"
Sub ForceAllHeadingsToHeiti_ByScan()
    Dim doc As Document
    Dim p As Paragraph
    Dim count As Integer
    
    Set doc = ActiveDocument
    count = 0
    
    ' Turn off screen updating: accelerate execution time
    Application.ScreenUpdating = False
    
    ' enumerate every paragraph
    For Each p In doc.Paragraphs
        
        ' OutlineLevel
        ' wdOutlineLevelBodyText represent "body text" (level 10)
        ' if not level 10 => headings
        If p.Format.OutlineLevel <> wdOutlineLevelBodyText Then
            
            ' --- reset font by force ---
            With p.Range.Font
                ' 1. FarEast (Asian) to Heiti
                .NameFarEast = "黑体"
                
                ' 2. Ascii to SimHei
                .NameAscii = "SimHei"
                
                ' 3. Color to Black (Auto)
                .ColorIndex = wdAuto
            End With
            
            count = count + 1
        End If
        
        ' 特殊情况补充：文档的主标题 (Title 样式) 有时大纲级别是正文，但它也是标题
        ' 如果需要包含主标题，取消下面这段注释：
        ' If p.Style = "标题" Or p.Style = "Title" Then
        '     p.Range.Font.NameFarEast = "黑体"
        '     p.Range.Font.NameAscii = "SimHei"
        ' End If
        
    Next p
    
    Application.ScreenUpdating = True
    
    MsgBox "扫描完成！已强制修改 " & count & " 个标题段落为黑体。"
End Sub