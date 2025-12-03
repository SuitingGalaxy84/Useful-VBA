Attribute VB_Name = "Module7"
Sub RemoveAllIndentsInTables_Safe()
    Dim doc As Document
    Dim tbl As Table
    Dim count As Integer
    
    Set doc = ActiveDocument
    count = 0
    
    ' 关闭屏幕更新，防止页面闪烁，大幅提高速度
    Application.ScreenUpdating = False
    
    ' 遍历文档中的每一个表格
    For Each tbl In doc.Tables
        
        ' 使用 With 结构直接操作整个表格的段落格式
        ' 这样做比遍历每一行每一格要快得多
        With tbl.Range.ParagraphFormat
            
            ' --- 1. 清除常规缩进 (磅值) ---
            .LeftIndent = 0         ' 左缩进
            .RightIndent = 0        ' 右缩进
            .FirstLineIndent = 0    ' 首行缩进/悬挂缩进
            
            ' --- 2. 清除字符级缩进 (针对中文版 Word) ---
            ' 在中文排版中，"缩进2字符" 实际上是储存在这里
            ' 必须设为0，否则 LeftIndent = 0 可能无效
            .CharacterUnitLeftIndent = 0
            .CharacterUnitRightIndent = 0
            .CharacterUnitFirstLineIndent = 0
            
            ' --- 3. 确保没有“镜像缩进”等奇葩设置 ---
            .MirrorIndents = False
            
        End With
        
        count = count + 1
    Next tbl
    
    Application.ScreenUpdating = True
    
    MsgBox "处理完成！" & vbCrLf & "已移除 " & count & " 个表格内的所有缩进设置。"
End Sub
