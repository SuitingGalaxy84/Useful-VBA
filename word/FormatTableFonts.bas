Attribute VB_Name = "FormatTableFonts"
Sub FormatTableFonts_MergeSafe()
    Dim doc As Document
    Dim tbl As Table
    Dim cel As Cell
    Dim count As Integer
    
    Set doc = ActiveDocument
    count = 0
    
    ' 关闭屏幕更新
    Application.ScreenUpdating = False
    
    For Each tbl In doc.Tables
        ' ==========================================
        ' 第一步：无视行列，将整个表格内容设为“其他”格式
        ' Asian: 宋体, Ascii: Times New Roman, Size: 10.5 (五号)
        ' ==========================================
        With tbl.Range.Font
            .NameFarEast = "宋体"
            .NameAscii = "Times New Roman"
            .NameOther = "Times New Roman"
            .Size = 10.5
            .Bold = False
            .Italic = False
        End With
        
        ' ==========================================
        ' 第二步：安全地处理表头 (逐个检查单元格)
        ' ==========================================
        ' 即使 .Rows 集合不可用，.Cells 集合通常是可用的
        For Each cel In tbl.Range.Cells
            ' 使用错误捕获，防止极少数极其复杂的损坏表格导致脚本中断
            On Error Resume Next
            
            ' 判断逻辑：如果单元格的行号为 1，则它是表头
            If cel.RowIndex = 1 Then
                With cel.Range.Font
                    .NameFarEast = "宋体"
                    .NameAscii = "Times New Roman"
                    .NameOther = "Times New Roman"
                    .Size = 10.5       ' 小四
                    .Bold = True    ' Regular
                End With
            ElseIf cel.RowIndex > 1 Then
                ' 优化：因为 Cell 集合是按顺序排列的
                ' 一旦发现行号大于 1，说明表头已经处理完了，直接跳出内层循环，处理下一个表格
                Exit For
            End If
            
            On Error GoTo 0
        Next cel
        
        count = count + 1
    Next tbl
    
    Application.ScreenUpdating = True
    MsgBox "处理完成！已格式化 " & count & " 个表格（包含合并单元格）。"
End Sub
