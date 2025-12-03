Attribute VB_Name = "Module6"
Sub ReplaceSEQIdentifiers()
    Dim doc As Document
    Dim fld As Field
    Dim codeText As String
    Dim originalCode As String
    Dim countFig As Integer
    Dim countTab As Integer
    
    Set doc = ActiveDocument
    countFig = 0
    countTab = 0
    
    ' 关闭屏幕更新以防止闪烁并提高速度
    Application.ScreenUpdating = False
    
    ' 遍历文档中的每一个域
    For Each fld In doc.Fields
        ' 仅处理序列域 (SEQ)
        If fld.Type = wdFieldSequence Then
            
            ' 获取原始域代码文本
            originalCode = fld.code.Text
            codeText = originalCode
            
            ' --- 处理 "SEQ 图" -> "SEQ Figure" ---
            ' 检查代码中是否包含 "SEQ 图" (注意 SEQ 和中文之间的空格)
            If InStr(codeText, "SEQ 图") > 0 Then
                ' 执行替换
                codeText = Replace(codeText, "SEQ 图", "SEQ Figure")
                countFig = countFig + 1
            End If
            
            ' --- 处理 "SEQ 表格" -> "SEQ Table" ---
            ' 检查代码中是否包含 "SEQ 表格"
            If InStr(codeText, "SEQ 表格") > 0 Then
                codeText = Replace(codeText, "SEQ 表格", "SEQ Table")
                countTab = countTab + 1
            End If
            
            ' ?? 补充情况：Word默认插入的中文标签通常是 "SEQ 表" 而不是 "SEQ 表格"
            ' 如果你的文档里用的是 "表"，请取消下面几行的注释：
            ' If InStr(codeText, "SEQ 表 ") > 0 Or Right(Trim(codeText), 3) = "SEQ 表" Then
            '     codeText = Replace(codeText, "SEQ 表", "SEQ Table")
            '     countTab = countTab + 1
            ' End If
            
            ' --- 应用更改 ---
            ' 只有当代码发生变化时才写入和更新
            If codeText <> originalCode Then
                fld.code.Text = codeText
                fld.Update ' 更新域以显示新编号（虽然数字可能不变，但序列名变了）
            End If
            
        End If
    Next fld
    
    ' 更新全文档的所有域，确保编号连贯
    doc.Fields.Update
    
    Application.ScreenUpdating = True
    
    MsgBox "处理完成！" & vbCrLf & _
           "替换 SEQ 图 -> Figure: " & countFig & " 处" & vbCrLf & _
           "替换 SEQ 表格 -> Table: " & countTab & ""
End Sub
