Partial Class Ribbon1
    Inherits Microsoft.Office.Tools.Ribbon.RibbonBase

    <System.Diagnostics.DebuggerNonUserCode()>
    Public Sub New(ByVal container As System.ComponentModel.IContainer)
        MyClass.New()

        'Windows.Forms 类撰写设计器支持所必需的
        If (container IsNot Nothing) Then
            container.Add(Me)
        End If

    End Sub

    <System.Diagnostics.DebuggerNonUserCode()>
    Public Sub New()
        MyBase.New(Globals.Factory.GetRibbonFactory())

        '组件设计器需要此调用。
        InitializeComponent()

    End Sub

    '组件重写释放以清理组件列表。
    <System.Diagnostics.DebuggerNonUserCode()>
    Protected Overrides Sub Dispose(ByVal disposing As Boolean)
        Try
            If disposing AndAlso components IsNot Nothing Then
                components.Dispose()
            End If
        Finally
            MyBase.Dispose(disposing)
        End Try
    End Sub

    '组件设计器所必需的
    Private components As System.ComponentModel.IContainer

    '注意: 以下过程是组件设计器所必需的
    '可使用组件设计器修改它。
    '不要使用代码编辑器修改它。
    <System.Diagnostics.DebuggerStepThrough()>
    Private Sub InitializeComponent()
        Dim Separator4 As Microsoft.Office.Tools.Ribbon.RibbonSeparator
        Dim RibbonDropDownItemImpl1 As Microsoft.Office.Tools.Ribbon.RibbonDropDownItem = Me.Factory.CreateRibbonDropDownItem
        Dim RibbonDropDownItemImpl2 As Microsoft.Office.Tools.Ribbon.RibbonDropDownItem = Me.Factory.CreateRibbonDropDownItem
        Dim RibbonDropDownItemImpl3 As Microsoft.Office.Tools.Ribbon.RibbonDropDownItem = Me.Factory.CreateRibbonDropDownItem
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(Ribbon1))
        Me.Tab1 = Me.Factory.CreateRibbonTab
        Me.Group1 = Me.Factory.CreateRibbonGroup
        Me.EditBox2 = Me.Factory.CreateRibbonEditBox
        Me.Separator2 = Me.Factory.CreateRibbonSeparator
        Me.Group3 = Me.Factory.CreateRibbonGroup
        Me.ButtonGroup1 = Me.Factory.CreateRibbonButtonGroup
        Me.EditBox1 = Me.Factory.CreateRibbonEditBox
        Me.EditBox4 = Me.Factory.CreateRibbonEditBox
        Me.ButtonGroup2 = Me.Factory.CreateRibbonButtonGroup
        Me.Label2 = Me.Factory.CreateRibbonLabel
        Me.Label1 = Me.Factory.CreateRibbonLabel
        Me.Separator1 = Me.Factory.CreateRibbonSeparator
        Me.ButtonGroup3 = Me.Factory.CreateRibbonButtonGroup
        Me.EditBox5 = Me.Factory.CreateRibbonEditBox
        Me.EditBox6 = Me.Factory.CreateRibbonEditBox
        Me.Separator3 = Me.Factory.CreateRibbonSeparator
        Me.ButtonGroup4 = Me.Factory.CreateRibbonButtonGroup
        Me.Label4 = Me.Factory.CreateRibbonLabel
        Me.Label3 = Me.Factory.CreateRibbonLabel
        Me.Separator6 = Me.Factory.CreateRibbonSeparator
        Me.ComboBox1 = Me.Factory.CreateRibbonComboBox
        Me.Separator5 = Me.Factory.CreateRibbonSeparator
        Me.Group2 = Me.Factory.CreateRibbonGroup
        Me.FontDialog1 = New System.Windows.Forms.FontDialog()
        Me.Button1 = Me.Factory.CreateRibbonButton
        Me.Button4 = Me.Factory.CreateRibbonButton
        Me.SplitButton1 = Me.Factory.CreateRibbonSplitButton
        Me.Button2 = Me.Factory.CreateRibbonButton
        Me.Tab2 = Me.Factory.CreateRibbonTab
        Me.Group5 = Me.Factory.CreateRibbonGroup
        Me.Button5 = Me.Factory.CreateRibbonButton
        Me.Button3 = Me.Factory.CreateRibbonButton
        Separator4 = Me.Factory.CreateRibbonSeparator
        Me.Tab1.SuspendLayout()
        Me.Group1.SuspendLayout()
        Me.Group3.SuspendLayout()
        Me.Group2.SuspendLayout()
        Me.Tab2.SuspendLayout()
        Me.Group5.SuspendLayout()
        Me.SuspendLayout()
        '
        'Separator4
        '
        Separator4.Name = "Separator4"
        '
        'Tab1
        '
        Me.Tab1.ControlId.ControlIdType = Microsoft.Office.Tools.Ribbon.RibbonControlIdType.Office
        Me.Tab1.Groups.Add(Me.Group1)
        Me.Tab1.Groups.Add(Me.Group3)
        Me.Tab1.Groups.Add(Me.Group2)
        Me.Tab1.Label = "GL插件"
        Me.Tab1.Name = "Tab1"
        '
        'Group1
        '
        Me.Group1.Items.Add(Me.EditBox2)
        Me.Group1.Items.Add(Me.Button1)
        Me.Group1.Items.Add(Me.Separator2)
        Me.Group1.Items.Add(Me.Button3)
        Me.Group1.Label = "红头"
        Me.Group1.Name = "Group1"
        '
        'EditBox2
        '
        Me.EditBox2.Label = "文字"
        Me.EditBox2.MaxLength = 11
        Me.EditBox2.Name = "EditBox2"
        Me.EditBox2.SizeString = "00000000000000000000"
        Me.EditBox2.Text = "曹县教育和体育局"
        '
        'Separator2
        '
        Me.Separator2.Name = "Separator2"
        '
        'Group3
        '
        Me.Group3.Items.Add(Me.ButtonGroup1)
        Me.Group3.Items.Add(Me.EditBox1)
        Me.Group3.Items.Add(Me.EditBox4)
        Me.Group3.Items.Add(Me.ButtonGroup2)
        Me.Group3.Items.Add(Separator4)
        Me.Group3.Items.Add(Me.Label2)
        Me.Group3.Items.Add(Me.Label1)
        Me.Group3.Items.Add(Me.Separator1)
        Me.Group3.Items.Add(Me.ButtonGroup3)
        Me.Group3.Items.Add(Me.EditBox5)
        Me.Group3.Items.Add(Me.EditBox6)
        Me.Group3.Items.Add(Me.Separator3)
        Me.Group3.Items.Add(Me.ButtonGroup4)
        Me.Group3.Items.Add(Me.Label4)
        Me.Group3.Items.Add(Me.Label3)
        Me.Group3.Items.Add(Me.Separator6)
        Me.Group3.Items.Add(Me.ComboBox1)
        Me.Group3.Items.Add(Me.Separator5)
        Me.Group3.Items.Add(Me.Button4)
        Me.Group3.Label = "格式"
        Me.Group3.Name = "Group3"
        '
        'ButtonGroup1
        '
        Me.ButtonGroup1.Name = "ButtonGroup1"
        '
        'EditBox1
        '
        Me.EditBox1.Label = "左"
        Me.EditBox1.Name = "EditBox1"
        Me.EditBox1.SizeString = "0000"
        Me.EditBox1.Text = "2.8"
        '
        'EditBox4
        '
        Me.EditBox4.Label = "右"
        Me.EditBox4.Name = "EditBox4"
        Me.EditBox4.SizeString = "0000"
        Me.EditBox4.Text = "2.8"
        '
        'ButtonGroup2
        '
        Me.ButtonGroup2.Name = "ButtonGroup2"
        '
        'Label2
        '
        Me.Label2.Label = "厘米"
        Me.Label2.Name = "Label2"
        '
        'Label1
        '
        Me.Label1.Label = "厘米"
        Me.Label1.Name = "Label1"
        '
        'Separator1
        '
        Me.Separator1.Name = "Separator1"
        '
        'ButtonGroup3
        '
        Me.ButtonGroup3.Name = "ButtonGroup3"
        '
        'EditBox5
        '
        Me.EditBox5.Label = "上"
        Me.EditBox5.Name = "EditBox5"
        Me.EditBox5.SizeString = "0000"
        Me.EditBox5.Text = "3.7"
        '
        'EditBox6
        '
        Me.EditBox6.Label = "下"
        Me.EditBox6.Name = "EditBox6"
        Me.EditBox6.SizeString = "0000"
        Me.EditBox6.Text = "3.2"
        '
        'Separator3
        '
        Me.Separator3.Name = "Separator3"
        '
        'ButtonGroup4
        '
        Me.ButtonGroup4.Name = "ButtonGroup4"
        '
        'Label4
        '
        Me.Label4.Label = "厘米"
        Me.Label4.Name = "Label4"
        '
        'Label3
        '
        Me.Label3.Label = "厘米"
        Me.Label3.Name = "Label3"
        '
        'Separator6
        '
        Me.Separator6.Name = "Separator6"
        '
        'ComboBox1
        '
        RibbonDropDownItemImpl1.Label = "(无)"
        RibbonDropDownItemImpl2.Label = "首行缩进"
        RibbonDropDownItemImpl3.Label = "悬挂缩进"
        Me.ComboBox1.Items.Add(RibbonDropDownItemImpl1)
        Me.ComboBox1.Items.Add(RibbonDropDownItemImpl2)
        Me.ComboBox1.Items.Add(RibbonDropDownItemImpl3)
        Me.ComboBox1.Label = "ComboBox1"
        Me.ComboBox1.Name = "ComboBox1"
        Me.ComboBox1.ShowLabel = False
        Me.ComboBox1.SizeString = "aaaaaaaa"
        Me.ComboBox1.Text = Nothing
        '
        'Separator5
        '
        Me.Separator5.Name = "Separator5"
        '
        'Group2
        '
        Me.Group2.Items.Add(Me.SplitButton1)
        Me.Group2.Label = "空白行"
        Me.Group2.Name = "Group2"
        '
        'Button1
        '
        Me.Button1.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge
        Me.Button1.Image = CType(resources.GetObject("Button1.Image"), System.Drawing.Image)
        Me.Button1.ImageName = "aaaaa"
        Me.Button1.Label = "生成小红头"
        Me.Button1.Name = "Button1"
        Me.Button1.ShowImage = True
        '
        'Button4
        '
        Me.Button4.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge
        Me.Button4.Image = CType(resources.GetObject("Button4.Image"), System.Drawing.Image)
        Me.Button4.Label = "设置格式"
        Me.Button4.Name = "Button4"
        Me.Button4.ShowImage = True
        '
        'SplitButton1
        '
        Me.SplitButton1.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge
        Me.SplitButton1.Image = CType(resources.GetObject("SplitButton1.Image"), System.Drawing.Image)
        Me.SplitButton1.Items.Add(Me.Button2)
        Me.SplitButton1.Label = "删除全部空白行"
        Me.SplitButton1.Name = "SplitButton1"
        Me.SplitButton1.SuperTip = "删除全文的空白行，请谨慎使用哦"
        '
        'Button2
        '
        Me.Button2.Image = CType(resources.GetObject("Button2.Image"), System.Drawing.Image)
        Me.Button2.Label = "删除选中空白行"
        Me.Button2.Name = "Button2"
        Me.Button2.ShowImage = True
        Me.Button2.SuperTip = "删除选中的空白行"
        '
        'Tab2
        '
        Me.Tab2.ControlId.ControlIdType = Microsoft.Office.Tools.Ribbon.RibbonControlIdType.Office
        Me.Tab2.ControlId.OfficeId = "TabHome"
        Me.Tab2.Groups.Add(Me.Group5)
        Me.Tab2.Label = "TabHome"
        Me.Tab2.Name = "Tab2"
        '
        'Group5
        '
        Me.Group5.Items.Add(Me.Button5)
        Me.Group5.Name = "Group5"
        '
        'Button5
        '
        Me.Button5.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge
        Me.Button5.Image = CType(resources.GetObject("Button5.Image"), System.Drawing.Image)
        Me.Button5.Label = "一键设置格式"
        Me.Button5.Name = "Button5"
        Me.Button5.ShowImage = True
        '
        'Button3
        '
        Me.Button3.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge
        Me.Button3.Image = CType(resources.GetObject("Button3.Image"), System.Drawing.Image)
        Me.Button3.ImageName = "aaaaa"
        Me.Button3.Label = "生成大红头"
        Me.Button3.Name = "Button3"
        Me.Button3.ShowImage = True
        '
        'Ribbon1
        '
        Me.Name = "Ribbon1"
        Me.RibbonType = "Microsoft.Word.Document"
        Me.Tabs.Add(Me.Tab1)
        Me.Tabs.Add(Me.Tab2)
        Me.Tab1.ResumeLayout(False)
        Me.Tab1.PerformLayout()
        Me.Group1.ResumeLayout(False)
        Me.Group1.PerformLayout()
        Me.Group3.ResumeLayout(False)
        Me.Group3.PerformLayout()
        Me.Group2.ResumeLayout(False)
        Me.Group2.PerformLayout()
        Me.Tab2.ResumeLayout(False)
        Me.Tab2.PerformLayout()
        Me.Group5.ResumeLayout(False)
        Me.Group5.PerformLayout()
        Me.ResumeLayout(False)

    End Sub

    Friend WithEvents Tab1 As Microsoft.Office.Tools.Ribbon.RibbonTab
    Friend WithEvents Group1 As Microsoft.Office.Tools.Ribbon.RibbonGroup
    Friend WithEvents Button1 As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents Group2 As Microsoft.Office.Tools.Ribbon.RibbonGroup
    Friend WithEvents SplitButton1 As Microsoft.Office.Tools.Ribbon.RibbonSplitButton
    Friend WithEvents Button2 As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents EditBox2 As Microsoft.Office.Tools.Ribbon.RibbonEditBox
    Friend WithEvents Separator2 As Microsoft.Office.Tools.Ribbon.RibbonSeparator
    Friend WithEvents Group3 As Microsoft.Office.Tools.Ribbon.RibbonGroup
    Friend WithEvents ButtonGroup1 As Microsoft.Office.Tools.Ribbon.RibbonButtonGroup
    Friend WithEvents EditBox1 As Microsoft.Office.Tools.Ribbon.RibbonEditBox
    Friend WithEvents EditBox4 As Microsoft.Office.Tools.Ribbon.RibbonEditBox
    Friend WithEvents ButtonGroup2 As Microsoft.Office.Tools.Ribbon.RibbonButtonGroup
    Friend WithEvents Label2 As Microsoft.Office.Tools.Ribbon.RibbonLabel
    Friend WithEvents Label1 As Microsoft.Office.Tools.Ribbon.RibbonLabel
    Friend WithEvents Separator1 As Microsoft.Office.Tools.Ribbon.RibbonSeparator
    Friend WithEvents ButtonGroup3 As Microsoft.Office.Tools.Ribbon.RibbonButtonGroup
    Friend WithEvents EditBox5 As Microsoft.Office.Tools.Ribbon.RibbonEditBox
    Friend WithEvents EditBox6 As Microsoft.Office.Tools.Ribbon.RibbonEditBox
    Friend WithEvents Separator3 As Microsoft.Office.Tools.Ribbon.RibbonSeparator
    Friend WithEvents ButtonGroup4 As Microsoft.Office.Tools.Ribbon.RibbonButtonGroup
    Friend WithEvents Label4 As Microsoft.Office.Tools.Ribbon.RibbonLabel
    Friend WithEvents Label3 As Microsoft.Office.Tools.Ribbon.RibbonLabel
    Friend WithEvents Button4 As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents Separator5 As Microsoft.Office.Tools.Ribbon.RibbonSeparator
    Friend WithEvents FontDialog1 As Windows.Forms.FontDialog
    Friend WithEvents Separator6 As Microsoft.Office.Tools.Ribbon.RibbonSeparator
    Friend WithEvents ComboBox1 As Microsoft.Office.Tools.Ribbon.RibbonComboBox
    Friend WithEvents Tab2 As Microsoft.Office.Tools.Ribbon.RibbonTab
    Friend WithEvents Group5 As Microsoft.Office.Tools.Ribbon.RibbonGroup
    Friend WithEvents Button5 As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents Button3 As Microsoft.Office.Tools.Ribbon.RibbonButton
End Class

Partial Class ThisRibbonCollection

    <System.Diagnostics.DebuggerNonUserCode()> _
    Friend ReadOnly Property Ribbon1() As Ribbon1
        Get
            Return Me.GetRibbon(Of Ribbon1)()
        End Get
    End Property
End Class
