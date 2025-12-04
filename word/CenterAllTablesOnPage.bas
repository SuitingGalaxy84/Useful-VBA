Attribute VB_Name = "CenterAllTablesOnPage"
Sub CenterAllTablesOnPage()
    Dim doc As Document
    Dim tbl As Table
    Dim count As Integer
    
    Set doc = ActiveDocument
    count = 0
    
    Application.ScreenUpdating = False
    
    For Each tbl In doc.Tables
        ' 1. 先清除表格自身的左缩进 (防止表格虽然居中但因为缩进而偏右)
        tbl.Rows.LeftIndent = 0
        
        ' 2. 将表格行的对齐方式设为居中
        tbl.Rows.Alignment = wdAlignRowCenter
        
        count = count + 1
    Next tbl
    
    Application.ScreenUpdating = True
    MsgBox "处理完成！已将 " & count & " 个表格调整为页面居中。"
End Sub
