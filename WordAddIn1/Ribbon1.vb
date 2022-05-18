Imports Microsoft.Office.Tools.Ribbon
Imports word = Microsoft.Office.Interop.Word


Public Class Ribbon1
    Dim wdapp As word.Application
    Dim wddoc As word.Document
    Private Sub Ribbon1_Load(ByVal sender As System.Object, ByVal e As RibbonUIEventArgs) Handles MyBase.Load

    End Sub


    Private Sub Button1_Click(sender As Object, e As RibbonControlEventArgs) Handles Button1.Click
        'MsgBox("第一个按钮")
        Dim wdapp As word.Application = Globals.ThisAddIn.Application '定义word程序
        Dim kuang As word.Shape
        kuang = wdapp.ActiveDocument.Shapes.AddTextbox(1, 79.38, 104.9, 436.6, 56.7)
        Dim L1, L2 As word.Shape
        L1 = wdapp.ActiveDocument.Shapes.AddLine(100, 100, 538, 100)
        Dim shang, xia, zuo, you As Double
        Dim kuan As Long = 595.35                    '纸的宽度
        Dim gao As Long = 841.95                     '纸的高度

        shang = wdapp.ActiveDocument.PageSetup.TopMargin        '获取上边距
        xia = wdapp.ActiveDocument.PageSetup.BottomMargin       '获取下边距
        zuo = wdapp.ActiveDocument.PageSetup.LeftMargin           '获取左边距
        you = wdapp.ActiveDocument.PageSetup.RightMargin           '获取左边距   单位是磅

        With kuang
            .Name = "text1"
            '设置外框不可见
            .Line.Visible = Microsoft.Office.Core.MsoTriState.msoFalse
            '相对页面固定
            .RelativeHorizontalPosition = Microsoft.Office.Interop.Word.WdRelativeHorizontalPosition.wdRelativeHorizontalPositionPage
            .RelativeVerticalPosition = Microsoft.Office.Interop.Word.WdRelativeVerticalPosition.wdRelativeVerticalPositionPage
            .Height = wdapp.Application.MillimetersToPoints(20)   '20mm高
            .Width = wdapp.ActiveDocument.PageSetup.PageWidth - zuo - you     '154mm宽
            .Top = shang       '上边距
            .Left = zuo      '左边距
            .TextFrame.TextRange.Text = EditBox2.Text       'caoxian jiaoyu he tiyu ju
        End With
        kuang.Select()
        With wdapp.Selection
            .Font.Name = "方正小标宋简体"   '字体
            .Font.Size = 36    '初号字
            .Font.ColorIndex = Microsoft.Office.Interop.Word.WdColorIndex.wdRed   '红色
            .Font.Bold = True
            .ParagraphFormat.Alignment = Microsoft.Office.Interop.Word.WdParagraphAlignment.wdAlignParagraphDistribute  '段落字符被分布排列，以填满整个段落宽度
            .ShapeRange.WrapFormat.Type = Microsoft.Office.Interop.Word.WdWrapType.wdWrapNone  '形状环绕方式：衬于文字下方
            .ShapeRange.Align(Microsoft.Office.Core.MsoAlignCmd.msoAlignCenters, True) '中间对齐  true表示相对于页面，false表示相对于形状

        End With



        With L1
            .Name = "x1"
            .Line.Weight = 4 '粗细 4 磅
            .Line.ForeColor.RGB = RGB(255, 0, 0)
            .Line.DashStyle = Microsoft.Office.Core.MsoLineDashStyle.msoLineSolid '长实线
            .Line.Style = Microsoft.Office.Core.MsoLineStyle.msoLineThickThin  '上粗下细
            .RelativeHorizontalPosition = Microsoft.Office.Interop.Word.WdRelativeHorizontalPosition.wdRelativeHorizontalPositionPage
            .RelativeVerticalPosition = Microsoft.Office.Interop.Word.WdRelativeVerticalPosition.wdRelativeVerticalPositionPage
            .Top = shang + wdapp.Application.MillimetersToPoints(20)
            .Width = wdapp.ActiveDocument.PageSetup.PageWidth - zuo - you

        End With
        L1.Select()
        With wdapp.Selection
            .ShapeRange.Align(Microsoft.Office.Core.MsoAlignCmd.msoAlignCenters, True) '想对于页面  中间对齐

        End With

        'wdapp.ActiveDocument.Shapes.SelectAll()
        'wdapp.Selection.ShapeRange.Group()





        L2 = wdapp.ActiveDocument.Shapes.AddLine(100, 200, 538, 200)

        With L2
            .Name = "x2"
            .Line.Weight = 4 '粗细 4 磅
            .Line.ForeColor.RGB = RGB(255, 0, 0)
            .Line.DashStyle = Microsoft.Office.Core.MsoLineDashStyle.msoLineSolid '长实线
            .Line.Style = Microsoft.Office.Core.MsoLineStyle.msoLineThinThick  '上粗下细
            '.RelativeHorizontalPosition = Microsoft.Office.Interop.Word.WdRelativeHorizontalPosition.wdRelativeHorizontalPositionPage
            .RelativeVerticalPosition = Microsoft.Office.Interop.Word.WdRelativeVerticalPosition.wdRelativeVerticalPositionBottomMarginArea '相对于下边距
            .Top = wdapp.Application.MillimetersToPoints(0)  '相对于下边距
            .Width = wdapp.ActiveDocument.PageSetup.PageWidth - zuo - you
        End With
        L2.Select()
        With wdapp.Selection
            .ShapeRange.Align(Microsoft.Office.Core.MsoAlignCmd.msoAlignCenters, True) '想对于页面  中间对齐

        End With






    End Sub

    Private Sub SplitButton1_Click(sender As Object, e As RibbonControlEventArgs) Handles SplitButton1.Click
        Dim t As Single
        Dim wdapp As word.Application = Globals.ThisAddIn.Application
        wdapp.ScreenUpdating = False
        t = Timer
        Dim s As word.Range
        s = wdapp.ActiveDocument.Content
        s.Find.Execute("^13[   ^t" & ChrW(160) & "^11^13]{1,}", , , 2, , , , , , "^p", 2)
        s = Nothing
        wdapp.ScreenUpdating = True
        MsgBox("去掉所有空白行用时  " & Timer - t & "秒")

    End Sub



    Private Sub Button3_Click(sender As Object, e As RibbonControlEventArgs)
        Dim wdapp As word.Application = Globals.ThisAddIn.Application '定义word程序
        Dim a, b As Double
        a = wdapp.ActiveDocument.PageSetup.LeftMargin
        b = wdapp.Application.PointsToCentimeters(a)
        MsgBox(ComboBox1.Text)

        MsgBox(a)

    End Sub

    Private Sub Button4_Click(sender As Object, e As RibbonControlEventArgs) Handles Button4.Click
        Dim wdapp As word.Application = Globals.ThisAddIn.Application
        With wdapp.ActiveDocument.PageSetup
            .LeftMargin = wdapp.Application.CentimetersToPoints(EditBox1.Text)
            .RightMargin = wdapp.Application.CentimetersToPoints(EditBox4.Text)
            .TopMargin = wdapp.Application.CentimetersToPoints(EditBox5.Text)
            .BottomMargin = wdapp.Application.CentimetersToPoints(EditBox6.Text)
            .Gutter = wdapp.Application.CentimetersToPoints(0) '装订线0cm
            .HeaderDistance = wdapp.Application.CentimetersToPoints(1.5) '页眉1.5cm
            .FooterDistance = wdapp.Application.CentimetersToPoints(1.75) '页脚1.75cm
            .PageWidth = wdapp.Application.CentimetersToPoints(21) '纸张宽21cm
            .PageHeight = wdapp.Application.CentimetersToPoints(29.7) '纸张高29.7cm
            .SectionStart = Microsoft.Office.Interop.Word.WdSectionStart.wdSectionNewPage '节的起始位置：新建页
            .OddAndEvenPagesHeaderFooter = False '不勾选“奇偶页不同”
            .DifferentFirstPageHeaderFooter = False '不勾选“首页不同”
            .VerticalAlignment = Microsoft.Office.Interop.Word.WdVerticalAlignment.wdAlignVerticalTop '页面垂直对齐方式为“顶端对齐”
            .SuppressEndnotes = False '不隐藏尾注
            .MirrorMargins = False '不设置首页的内外边距
            .BookFoldRevPrinting = False '不设置手动双面打印
            .BookFoldPrintingSheets = 1 '默认打印份数为1
            .GutterPos = Microsoft.Office.Interop.Word.WdGutterStyle.wdGutterPosLeft '装订线位于左侧
            .LayoutMode = Microsoft.Office.Interop.Word.WdLayoutMode.wdLayoutModeLineGrid '版式模式为“只指定行网格”
            .Orientation = Microsoft.Office.Interop.Word.WdOrientation.wdOrientPortrait '页面方向为纵向
        End With

        Dim suojin As Single


        If ComboBox1.Text = "首行缩进" Then
            suojin = 2
        ElseIf ComboBox1.Text = "悬挂缩进" Then
            suojin = -2
        Else
            suojin = 0
        End If



        With wdapp.ActiveDocument.Paragraphs
            .LineSpacingRule = Microsoft.Office.Interop.Word.WdLineSpacing.wdLineSpaceExactly   '固定行离  28磅
            .LineSpacing = 28
            .Alignment = Microsoft.Office.Interop.Word.WdParagraphAlignment.wdAlignParagraphJustify    '两端对齐
            .CharacterUnitFirstLineIndent = suojin
            .LeftIndent = wdapp.Application.CentimetersToPoints(0) '左缩进0cm
            .RightIndent = wdapp.Application.CentimetersToPoints(0) '右缩进0cm
            .SpaceBefore = 0 '段前间距0cm
            .SpaceBeforeAuto = False '段前间距不设为“自动”
            .SpaceAfter = 0 '段后间距0cm
            .SpaceAfterAuto = False '段后间距不设为“自动”
            .WidowControl = False '不勾选“孤行控制”
            .KeepWithNext = False '不勾选“与下段同页”
            .KeepTogether = False '不勾选“段中不分页”
            .PageBreakBefore = False '不勾选“段前同页”
            .NoLineNumber = False '不勾选“取消行号”
            .Hyphenation = True '不勾选“允许西文在单词中间换行”
            .FirstLineIndent = wdapp.Application.CentimetersToPoints(0) '首行缩进0cm
            .OutlineLevel = Microsoft.Office.Interop.Word.WdOutlineLevel.wdOutlineLevelBodyText  '大纲级别为“正文文本”
            .CharacterUnitLeftIndent = 0 '段落左缩进0cm
            .CharacterUnitRightIndent = 0 '段落右缩进0cm
            .LineUnitBefore = 0 '段前间距为0
            .LineUnitAfter = 0 '段后间距为0
            .AutoAdjustRightIndent = False '自动调整段落的右缩进   false bu dai gou
            .DisableLineHeightGrid = True '不勾选“如果定义了文档网格，则对齐网格”，即指定段落中的字符与行网格对齐  true bu dai gou
            .FarEastLineBreakControl = True '将东亚语言文字的换行规则应用于指定的段落
            .WordWrap = True '在指定段落或文本框的西文单词中间断字换行
            .HangingPunctuation = True '指定段落中的标点将可以溢出边界
            .HalfWidthPunctuationOnTopOfLine = False
            .AddSpaceBetweenFarEastAndAlpha = True '自动在指定段落的中文文字和拉丁文字之间添加空格。
            .AddSpaceBetweenFarEastAndDigit = True '自动在指定段落中的中文文字与数字之间添加空格
            .BaseLineAlignment = Microsoft.Office.Interop.Word.WdBaselineAlignment.wdBaselineAlignAuto '自动调整基线字体对齐方式

        End With

        With wdapp.ActiveDocument.Content
            With .Font
                .NameFarEast = "宋体"
                .NameAscii = "times new roman"
                .Size = 16
            End With
            With .Paragraphs.First
                .Range.Font.Size = 16
                .Alignment = Microsoft.Office.Interop.Word.WdParagraphAlignment.wdAlignParagraphJustify
            End With
        End With

        With wdapp.ActiveDocument.Paragraphs
            .LineSpacingRule = Microsoft.Office.Interop.Word.WdLineSpacing.wdLineSpaceExactly
            .LineSpacing = 28
        End With


    End Sub

    Private Sub Button2_Click(sender As Object, e As RibbonControlEventArgs) Handles Button2.Click
        Dim t As Single
        Dim wdapp As word.Application = Globals.ThisAddIn.Application
        wdapp.ScreenUpdating = False
        t = Timer
        Dim s As word.Selection
        's = wdapp.ActiveDocument.Content
        s = wdapp.Selection
        s.Find.Execute("^13[   ^t" & ChrW(160) & "^11^13]{1,}", , , 2, , , , , , "^p", 2)
        s = Nothing
        wdapp.ScreenUpdating = True
        MsgBox("去掉选中空白行用时  " & Timer - t & "秒")
    End Sub

    Private Sub Button5_Click(sender As Object, e As RibbonControlEventArgs) Handles Button5.Click
        Dim wdapp As word.Application = Globals.ThisAddIn.Application
        With wdapp.ActiveDocument.PageSetup
            .LeftMargin = wdapp.Application.CentimetersToPoints(EditBox1.Text)
            .RightMargin = wdapp.Application.CentimetersToPoints(EditBox4.Text)
            .TopMargin = wdapp.Application.CentimetersToPoints(EditBox5.Text)
            .BottomMargin = wdapp.Application.CentimetersToPoints(EditBox6.Text)
            .Gutter = wdapp.Application.CentimetersToPoints(0) '装订线0cm
            .HeaderDistance = wdapp.Application.CentimetersToPoints(1.5) '页眉1.5cm
            .FooterDistance = wdapp.Application.CentimetersToPoints(1.75) '页脚1.75cm
            .PageWidth = wdapp.Application.CentimetersToPoints(21) '纸张宽21cm
            .PageHeight = wdapp.Application.CentimetersToPoints(29.7) '纸张高29.7cm
            .SectionStart = Microsoft.Office.Interop.Word.WdSectionStart.wdSectionNewPage '节的起始位置：新建页
            .OddAndEvenPagesHeaderFooter = False '不勾选“奇偶页不同”
            .DifferentFirstPageHeaderFooter = False '不勾选“首页不同”
            .VerticalAlignment = Microsoft.Office.Interop.Word.WdVerticalAlignment.wdAlignVerticalTop '页面垂直对齐方式为“顶端对齐”
            .SuppressEndnotes = False '不隐藏尾注
            .MirrorMargins = False '不设置首页的内外边距
            .BookFoldRevPrinting = False '不设置手动双面打印
            .BookFoldPrintingSheets = 1 '默认打印份数为1
            .GutterPos = Microsoft.Office.Interop.Word.WdGutterStyle.wdGutterPosLeft '装订线位于左侧
            .LayoutMode = Microsoft.Office.Interop.Word.WdLayoutMode.wdLayoutModeLineGrid '版式模式为“只指定行网格”
            .Orientation = Microsoft.Office.Interop.Word.WdOrientation.wdOrientPortrait '页面方向为纵向
        End With

        Dim suojin As Single


        If ComboBox1.Text = "首行缩进" Then
            suojin = 2
        ElseIf ComboBox1.Text = "悬挂缩进" Then
            suojin = -2
        Else
            suojin = 0
        End If



        With wdapp.ActiveDocument.Paragraphs
            .LineSpacingRule = Microsoft.Office.Interop.Word.WdLineSpacing.wdLineSpaceExactly   '固定行离  28磅
            .LineSpacing = 28
            .Alignment = Microsoft.Office.Interop.Word.WdParagraphAlignment.wdAlignParagraphJustify    '两端对齐
            .CharacterUnitFirstLineIndent = suojin
            .LeftIndent = wdapp.Application.CentimetersToPoints(0) '左缩进0cm
            .RightIndent = wdapp.Application.CentimetersToPoints(0) '右缩进0cm
            .SpaceBefore = 0 '段前间距0cm
            .SpaceBeforeAuto = False '段前间距不设为“自动”
            .SpaceAfter = 0 '段后间距0cm
            .SpaceAfterAuto = False '段后间距不设为“自动”
            .WidowControl = False '不勾选“孤行控制”
            .KeepWithNext = False '不勾选“与下段同页”
            .KeepTogether = False '不勾选“段中不分页”
            .PageBreakBefore = False '不勾选“段前同页”
            .NoLineNumber = False '不勾选“取消行号”
            .Hyphenation = True '不勾选“允许西文在单词中间换行”
            .FirstLineIndent = wdapp.Application.CentimetersToPoints(0) '首行缩进0cm
            .OutlineLevel = Microsoft.Office.Interop.Word.WdOutlineLevel.wdOutlineLevelBodyText  '大纲级别为“正文文本”
            .CharacterUnitLeftIndent = 0 '段落左缩进0cm
            .CharacterUnitRightIndent = 0 '段落右缩进0cm
            .LineUnitBefore = 0 '段前间距为0
            .LineUnitAfter = 0 '段后间距为0
            .AutoAdjustRightIndent = False '自动调整段落的右缩进   false bu dai gou
            .DisableLineHeightGrid = True '不勾选“如果定义了文档网格，则对齐网格”，即指定段落中的字符与行网格对齐  true bu dai gou
            .FarEastLineBreakControl = True '将东亚语言文字的换行规则应用于指定的段落
            .WordWrap = True '在指定段落或文本框的西文单词中间断字换行
            .HangingPunctuation = True '指定段落中的标点将可以溢出边界
            .HalfWidthPunctuationOnTopOfLine = False
            .AddSpaceBetweenFarEastAndAlpha = True '自动在指定段落的中文文字和拉丁文字之间添加空格。
            .AddSpaceBetweenFarEastAndDigit = True '自动在指定段落中的中文文字与数字之间添加空格
            .BaseLineAlignment = Microsoft.Office.Interop.Word.WdBaselineAlignment.wdBaselineAlignAuto '自动调整基线字体对齐方式

        End With

        With wdapp.ActiveDocument.Content
            With .Font
                .NameFarEast = "宋体"
                .NameAscii = "times new roman"
                .Size = 16
            End With
            With .Paragraphs.First
                .Range.Font.Size = 16
                .Alignment = Microsoft.Office.Interop.Word.WdParagraphAlignment.wdAlignParagraphJustify
            End With
        End With

        With wdapp.ActiveDocument.Paragraphs
            .LineSpacingRule = Microsoft.Office.Interop.Word.WdLineSpacing.wdLineSpaceExactly
            .LineSpacing = 28
        End With
    End Sub

    Private Sub Button3_Click_1(sender As Object, e As RibbonControlEventArgs) Handles Button3.Click
        MsgBox("此功能暂未开发", Title:="Sorry")
    End Sub
End Class
