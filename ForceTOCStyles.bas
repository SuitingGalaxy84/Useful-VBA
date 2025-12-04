Attribute VB_Name = "ForceTOCStyles"
Sub ForceTOCStyles()
    Dim doc As Document
    Dim toc As TableOfContents
    Dim para As Paragraph
    Dim styleName As String
    
    Set doc = ActiveDocument
    
    If doc.TablesOfContents.count = 0 Then
        MsgBox "未找到目录！"
        Exit Sub
    End If
    
    Set toc = doc.TablesOfContents(1)
    
    Application.ScreenUpdating = False
    
    ' 1. 先更新目录，保证内容最新 (此时字体可能会变乱)
    toc.Update
    
    ' 2. 遍历目录里的每一行
    For Each para In toc.Range.Paragraphs
        ' 获取这一行当前应用的样式名 (例如 "TOC 2" 或 "目录 2")
        styleName = para.Style.NameLocal
        
        ' 3. 关键步骤：重新应用样式 + 清除直接格式
        ' 只要样式名里包含 "TOC" 或 "目录"，我们就重置它
        If InStr(styleName, "TOC 2") > 0 Or InStr(styleName, "目录") > 0 Then
            
            ' 重新将该段落的样式设为它自己 (这一步是为了激活样式覆盖)
            ' para.Style = doc.Styles(styleName)
            
            ' ！！！核心大招！！！
            ' 对该段落执行 "Ctrl + Space" (清除字符格式)
            ' 这会把从正文标题继承来的 "黑体" 属性抹掉，只保留 TOC 样式定义的 "仿宋"
            para.Range.Font.NameFarEast = "FangSong"
            para.Range.Font.NameAscii = "Times New Roman"
            
        End If
    Next para
    
    Application.ScreenUpdating = True
    MsgBox "已强制重置目录格式！现在应该显示为 TOC 样式定义的字体了。"
End Sub
