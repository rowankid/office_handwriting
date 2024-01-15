Attribute VB_Name = "handwriting"
Sub handwriting()
'
' handwriting Macro
'
'

    Dim font_count As Long
    font_count = 11
    
    Dim font_config() As FontConfig
    ReDim font_config(font_count)
    Dim total_probability As Double
    
    
    
    For i = 0 To (font_count - 1)
        Set font_config(i) = New FontConfig
    Next i
    
    
    font_config(0).InitializeWithValues "世界那么大", 18, -2, 20
    font_config(1).InitializeWithValues "美玉体", 16, 0, 25
    font_config(2).InitializeWithValues "方正静蕾简体", 14, 2, 15
    font_config(3).InitializeWithValues "文鼎大钢笔行楷", 14, 2, 20
    font_config(4).InitializeWithValues "汉仪井柏然体简", 17, -1, 15
    font_config(5).InitializeWithValues "伯乐童年体", 15, 0, 0
    font_config(6).InitializeWithValues "伯乐字库竹笋体", 15, 0, 15
    font_config(7).InitializeWithValues "华康翩翩体W3P", 15, 1.2, 8
    font_config(8).InitializeWithValues "BoLeYaYati", 16, 0, 0
    font_config(9).InitializeWithValues "汉仪PP体简", 15, 1.2, 0
    font_config(10).InitializeWithValues "伯乐俏皮体", 15, 0, 0
    
    
    total_probability = 0
    For i = 0 To (font_count - 1)
        total_probability = total_probability + font_config(i).probability
    Next i
    
    ' 初始化随机数生成器
    VBA.Randomize

    ' 初始化字体比例变量
    Dim last_font_ratio As Double
    Dim font_ratio As Double
    Dim font_size As Double
    last_font_ratio = 0.2

    ' 检查是否有选中的文本
    If Not Selection.Range.Text = vbNullString Then
        ' 遍历选中文本中的每个字符
        For Each R_Character In Selection.Range.Characters
            Dim random As Integer
            random = Int(VBA.Rnd * total_probability)

            ' 选择字体
            Dim current_font As FontConfig
            Dim current_count As Double
            current_count = 0
            For i = 0 To (font_count - 1)
                current_count = current_count + font_config(i).probability
                If random < current_count Then
                    Set current_font = font_config(i)
                    Exit For
                End If
            Next i

            ' 计算并应用字体样式
            font_ratio = last_font_ratio + (0.1 * VBA.Rnd - 0.05)
            If font_ratio > 0.25 Then font_ratio = 0.25
            If font_ratio < 0.15 Then font_ratio = 0.15
            last_font_ratio = font_ratio
            font_size = current_font.size * (1 + last_font_ratio)

            R_Character.Font.name = current_font.name
            R_Character.Font.size = font_size
            R_Character.Font.Position = -(VBA.Rnd * 0.2 + 0.1) * (font_size - 15)
            R_Character.Font.Spacing = current_font.expanded + VBA.Rnd * 2 - 2
        Next R_Character
    End If

    ' 更新屏幕显示
    Application.ScreenUpdating = True

End Sub
