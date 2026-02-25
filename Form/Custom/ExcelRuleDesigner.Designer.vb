<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()>
Partial Class ExcelRuleDesigner
  Inherits System.Windows.Forms.UserControl

  'UserControl overrides dispose to clean up the component list.
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

  'Required by the Windows Form Designer
  Private components As System.ComponentModel.IContainer

  'NOTE: The following procedure is required by the Windows Form Designer
  'It can be modified using the Windows Form Designer.  
  'Do not modify it using the code editor.
  <System.Diagnostics.DebuggerStepThrough()>
  Private Sub InitializeComponent()
    Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(ExcelRuleDesigner))
    Me.tabPane = New System.Windows.Forms.TabControl()
    Me.tabRules = New System.Windows.Forms.TabPage()
    Me.pnlRule = New System.Windows.Forms.Panel()
    Me.tlpRule = New System.Windows.Forms.TableLayoutPanel()
    Me.txtRuleFilterExpression = New System.Windows.Forms.RichTextBox()
    Me.lvRuleFilters = New System.Windows.Forms.ListView()
        Me.RuleFilter = CType(New System.Windows.Forms.ColumnHeader(), System.Windows.Forms.ColumnHeader)
        Me.lvRuleValues = New System.Windows.Forms.ListView()
        Me.RuleField = CType(New System.Windows.Forms.ColumnHeader(), System.Windows.Forms.ColumnHeader)
        Me.lblRuleDragFields = New System.Windows.Forms.Label()
        Me.tvRuleFields = New System.Windows.Forms.TreeView()
        Me.lblRuleChooseFields = New System.Windows.Forms.Label()
        Me.txtRuleName = New System.Windows.Forms.TextBox()
        Me.lblRuleHeader = New System.Windows.Forms.Label()
        Me.cmbRuleNames = New System.Windows.Forms.ComboBox()
        Me.lblRuleRuleName = New System.Windows.Forms.Label()
        Me.lblRuleSelectRule = New System.Windows.Forms.Label()
        Me.flpLabelValues = New System.Windows.Forms.FlowLayoutPanel()
        Me.pbValues = New System.Windows.Forms.PictureBox()
        Me.lblValues = New System.Windows.Forms.Label()
        Me.flpLabelFilters = New System.Windows.Forms.FlowLayoutPanel()
        Me.pbFilter = New System.Windows.Forms.PictureBox()
        Me.lblFilters = New System.Windows.Forms.Label()
        Me.flpRuleButtons = New System.Windows.Forms.FlowLayoutPanel()
        Me.btnRuleSave = New System.Windows.Forms.Button()
        Me.btnRuleDelete = New System.Windows.Forms.Button()
        Me.pnlRuleOptions = New System.Windows.Forms.Panel()
        Me.tlpRuleOptions = New System.Windows.Forms.TableLayoutPanel()
        Me.optRange = New System.Windows.Forms.RadioButton()
        Me.optList = New System.Windows.Forms.RadioButton()
        Me.optSingle = New System.Windows.Forms.RadioButton()
        Me.tabApply = New System.Windows.Forms.TabPage()
        Me.pnlApply = New System.Windows.Forms.Panel()
        Me.tlpApply = New System.Windows.Forms.TableLayoutPanel()
        Me.cmbApplyNames = New System.Windows.Forms.ComboBox()
        Me.txtApplyFilterExpression = New System.Windows.Forms.RichTextBox()
        Me.lvApplyFilters = New System.Windows.Forms.ListView()
        Me.ApplyFilter = CType(New System.Windows.Forms.ColumnHeader(), System.Windows.Forms.ColumnHeader)
        Me.lblApplyFilters = New System.Windows.Forms.Label()
        Me.cmbApplyListSelectType = New System.Windows.Forms.ComboBox()
        Me.lblApplyListSelectType = New System.Windows.Forms.Label()
        Me.cmbApplyRules = New System.Windows.Forms.ComboBox()
        Me.lblApplySelectRule = New System.Windows.Forms.Label()
        Me.txtApplyName = New System.Windows.Forms.TextBox()
        Me.lblApplyName = New System.Windows.Forms.Label()
        Me.lblApplySelectApply = New System.Windows.Forms.Label()
        Me.lblApplyHeader = New System.Windows.Forms.Label()
        Me.flpApplyButtons = New System.Windows.Forms.FlowLayoutPanel()
        Me.btnApplySave = New System.Windows.Forms.Button()
        Me.btnApplyDelete = New System.Windows.Forms.Button()
        Me.tabPane.SuspendLayout()
        Me.tabRules.SuspendLayout()
        Me.pnlRule.SuspendLayout()
        Me.tlpRule.SuspendLayout()
        Me.flpLabelValues.SuspendLayout()
        CType(Me.pbValues, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.flpLabelFilters.SuspendLayout()
        CType(Me.pbFilter, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.flpRuleButtons.SuspendLayout()
        Me.pnlRuleOptions.SuspendLayout()
        Me.tlpRuleOptions.SuspendLayout()
        Me.tabApply.SuspendLayout()
        Me.pnlApply.SuspendLayout()
        Me.tlpApply.SuspendLayout()
        Me.flpApplyButtons.SuspendLayout()
        Me.SuspendLayout()
        '
        'tabPane
        '
        Me.tabPane.Alignment = System.Windows.Forms.TabAlignment.Bottom
        Me.tabPane.Controls.Add(Me.tabRules)
        Me.tabPane.Controls.Add(Me.tabApply)
        Me.tabPane.Dock = System.Windows.Forms.DockStyle.Fill
        Me.tabPane.Location = New System.Drawing.Point(0, 0)
        Me.tabPane.Margin = New System.Windows.Forms.Padding(4)
        Me.tabPane.Name = "tabPane"
        Me.tabPane.SelectedIndex = 0
        Me.tabPane.Size = New System.Drawing.Size(300, 788)
        Me.tabPane.TabIndex = 0
        '
        'tabRules
        '
        Me.tabRules.Controls.Add(Me.pnlRule)
        Me.tabRules.Location = New System.Drawing.Point(4, 4)
        Me.tabRules.Margin = New System.Windows.Forms.Padding(4)
        Me.tabRules.Name = "tabRules"
        Me.tabRules.Padding = New System.Windows.Forms.Padding(4)
        Me.tabRules.Size = New System.Drawing.Size(292, 758)
        Me.tabRules.TabIndex = 0
        Me.tabRules.Text = "Rules"
        Me.tabRules.UseVisualStyleBackColor = True
        '
        'pnlRule
        '
        Me.pnlRule.AutoScroll = True
        Me.pnlRule.AutoSizeMode = System.Windows.Forms.AutoSizeMode.GrowAndShrink
        Me.pnlRule.Controls.Add(Me.tlpRule)
        Me.pnlRule.Dock = System.Windows.Forms.DockStyle.Fill
        Me.pnlRule.Location = New System.Drawing.Point(4, 4)
        Me.pnlRule.Margin = New System.Windows.Forms.Padding(0)
        Me.pnlRule.Name = "pnlRule"
        Me.pnlRule.Padding = New System.Windows.Forms.Padding(4)
        Me.pnlRule.Size = New System.Drawing.Size(284, 750)
        Me.pnlRule.TabIndex = 0
        '
        'tlpRule
        '
        Me.tlpRule.AutoSizeMode = System.Windows.Forms.AutoSizeMode.GrowAndShrink
        Me.tlpRule.ColumnCount = 1
        Me.tlpRule.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 100.0!))
        Me.tlpRule.Controls.Add(Me.txtRuleFilterExpression, 0, 14)
        Me.tlpRule.Controls.Add(Me.lvRuleFilters, 0, 13)
        Me.tlpRule.Controls.Add(Me.lvRuleValues, 0, 9)
        Me.tlpRule.Controls.Add(Me.lblRuleDragFields, 0, 7)
        Me.tlpRule.Controls.Add(Me.tvRuleFields, 0, 6)
        Me.tlpRule.Controls.Add(Me.lblRuleChooseFields, 0, 5)
        Me.tlpRule.Controls.Add(Me.txtRuleName, 0, 4)
        Me.tlpRule.Controls.Add(Me.lblRuleHeader, 0, 0)
        Me.tlpRule.Controls.Add(Me.cmbRuleNames, 0, 2)
        Me.tlpRule.Controls.Add(Me.lblRuleRuleName, 0, 3)
        Me.tlpRule.Controls.Add(Me.lblRuleSelectRule, 0, 1)
        Me.tlpRule.Controls.Add(Me.flpLabelValues, 0, 8)
        Me.tlpRule.Controls.Add(Me.flpLabelFilters, 0, 12)
        Me.tlpRule.Controls.Add(Me.flpRuleButtons, 0, 15)
        Me.tlpRule.Controls.Add(Me.pnlRuleOptions, 0, 10)
        Me.tlpRule.Dock = System.Windows.Forms.DockStyle.Top
        Me.tlpRule.Location = New System.Drawing.Point(4, 4)
        Me.tlpRule.Margin = New System.Windows.Forms.Padding(0)
        Me.tlpRule.Name = "tlpRule"
        Me.tlpRule.RowCount = 16
        Me.tlpRule.RowStyles.Add(New System.Windows.Forms.RowStyle())
        Me.tlpRule.RowStyles.Add(New System.Windows.Forms.RowStyle())
        Me.tlpRule.RowStyles.Add(New System.Windows.Forms.RowStyle())
        Me.tlpRule.RowStyles.Add(New System.Windows.Forms.RowStyle())
        Me.tlpRule.RowStyles.Add(New System.Windows.Forms.RowStyle())
        Me.tlpRule.RowStyles.Add(New System.Windows.Forms.RowStyle())
        Me.tlpRule.RowStyles.Add(New System.Windows.Forms.RowStyle())
        Me.tlpRule.RowStyles.Add(New System.Windows.Forms.RowStyle())
        Me.tlpRule.RowStyles.Add(New System.Windows.Forms.RowStyle())
        Me.tlpRule.RowStyles.Add(New System.Windows.Forms.RowStyle())
        Me.tlpRule.RowStyles.Add(New System.Windows.Forms.RowStyle())
        Me.tlpRule.RowStyles.Add(New System.Windows.Forms.RowStyle())
        Me.tlpRule.RowStyles.Add(New System.Windows.Forms.RowStyle())
        Me.tlpRule.RowStyles.Add(New System.Windows.Forms.RowStyle())
        Me.tlpRule.RowStyles.Add(New System.Windows.Forms.RowStyle())
        Me.tlpRule.RowStyles.Add(New System.Windows.Forms.RowStyle())
        Me.tlpRule.Size = New System.Drawing.Size(276, 741)
        Me.tlpRule.TabIndex = 18
        '
        'txtRuleFilterExpression
        '
        Me.txtRuleFilterExpression.BackColor = System.Drawing.SystemColors.Control
        Me.txtRuleFilterExpression.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.txtRuleFilterExpression.Dock = System.Windows.Forms.DockStyle.Fill
        Me.txtRuleFilterExpression.Location = New System.Drawing.Point(3, 584)
        Me.txtRuleFilterExpression.MinimumSize = New System.Drawing.Size(4, 60)
        Me.txtRuleFilterExpression.Name = "txtRuleFilterExpression"
        Me.txtRuleFilterExpression.ReadOnly = True
        Me.txtRuleFilterExpression.ScrollBars = System.Windows.Forms.RichTextBoxScrollBars.ForcedVertical
        Me.txtRuleFilterExpression.Size = New System.Drawing.Size(270, 80)
        Me.txtRuleFilterExpression.TabIndex = 16
        Me.txtRuleFilterExpression.Text = ""
        '
        'lvRuleFilters
        '
        Me.lvRuleFilters.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.lvRuleFilters.Columns.AddRange(New System.Windows.Forms.ColumnHeader() {Me.RuleFilter})
        Me.lvRuleFilters.FullRowSelect = True
        Me.lvRuleFilters.HeaderStyle = System.Windows.Forms.ColumnHeaderStyle.None
        Me.lvRuleFilters.HideSelection = False
        Me.lvRuleFilters.Location = New System.Drawing.Point(3, 498)
        Me.lvRuleFilters.MinimumSize = New System.Drawing.Size(4, 80)
        Me.lvRuleFilters.MultiSelect = False
        Me.lvRuleFilters.Name = "lvRuleFilters"
        Me.lvRuleFilters.Size = New System.Drawing.Size(270, 80)
        Me.lvRuleFilters.TabIndex = 15
        Me.lvRuleFilters.UseCompatibleStateImageBehavior = False
        Me.lvRuleFilters.View = System.Windows.Forms.View.Details
        '
        'lvRuleValues
        '
        Me.lvRuleValues.Columns.AddRange(New System.Windows.Forms.ColumnHeader() {Me.RuleField})
        Me.lvRuleValues.Dock = System.Windows.Forms.DockStyle.Fill
        Me.lvRuleValues.FullRowSelect = True
        Me.lvRuleValues.HeaderStyle = System.Windows.Forms.ColumnHeaderStyle.None
        Me.lvRuleValues.HideSelection = False
        Me.lvRuleValues.Location = New System.Drawing.Point(3, 316)
        Me.lvRuleValues.MinimumSize = New System.Drawing.Size(4, 80)
        Me.lvRuleValues.MultiSelect = False
        Me.lvRuleValues.Name = "lvRuleValues"
        Me.lvRuleValues.Size = New System.Drawing.Size(270, 80)
        Me.lvRuleValues.TabIndex = 12
        Me.lvRuleValues.UseCompatibleStateImageBehavior = False
        Me.lvRuleValues.View = System.Windows.Forms.View.Details
        '
        'lblRuleDragFields
        '
        Me.lblRuleDragFields.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.lblRuleDragFields.AutoSize = True
        Me.lblRuleDragFields.Location = New System.Drawing.Point(3, 268)
        Me.lblRuleDragFields.Name = "lblRuleDragFields"
        Me.lblRuleDragFields.Size = New System.Drawing.Size(164, 17)
        Me.lblRuleDragFields.TabIndex = 10
        Me.lblRuleDragFields.Text = "Drag fields between areas:"
        '
        'tvRuleFields
        '
        Me.tvRuleFields.Dock = System.Windows.Forms.DockStyle.Fill
        Me.tvRuleFields.HideSelection = False
        Me.tvRuleFields.Location = New System.Drawing.Point(4, 144)
        Me.tvRuleFields.Margin = New System.Windows.Forms.Padding(4)
        Me.tvRuleFields.MinimumSize = New System.Drawing.Size(4, 100)
        Me.tvRuleFields.Name = "tvRuleFields"
        Me.tvRuleFields.Size = New System.Drawing.Size(268, 120)
        Me.tvRuleFields.TabIndex = 9
        '
        'lblRuleChooseFields
        '
        Me.lblRuleChooseFields.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.lblRuleChooseFields.AutoSize = True
        Me.lblRuleChooseFields.Location = New System.Drawing.Point(3, 123)
        Me.lblRuleChooseFields.Name = "lblRuleChooseFields"
        Me.lblRuleChooseFields.Size = New System.Drawing.Size(197, 17)
        Me.lblRuleChooseFields.TabIndex = 8
        Me.lblRuleChooseFields.Text = "Choose fields to add to the rule:"
        '
        'txtRuleName
        '
        Me.txtRuleName.Anchor = CType((System.Windows.Forms.AnchorStyles.Left Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.txtRuleName.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.txtRuleName.Location = New System.Drawing.Point(4, 94)
        Me.txtRuleName.Margin = New System.Windows.Forms.Padding(4)
        Me.txtRuleName.Name = "txtRuleName"
        Me.txtRuleName.Size = New System.Drawing.Size(268, 25)
        Me.txtRuleName.TabIndex = 7
        '
        'lblRuleHeader
        '
        Me.lblRuleHeader.AutoSize = True
        Me.lblRuleHeader.Font = New System.Drawing.Font("Segoe UI", 14.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblRuleHeader.Location = New System.Drawing.Point(3, 0)
        Me.lblRuleHeader.Name = "lblRuleHeader"
        Me.lblRuleHeader.Size = New System.Drawing.Size(109, 25)
        Me.lblRuleHeader.TabIndex = 3
        Me.lblRuleHeader.Text = "Define Rule"
        '
        'cmbRuleNames
        '
        Me.cmbRuleNames.Anchor = CType((System.Windows.Forms.AnchorStyles.Left Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.cmbRuleNames.FormattingEnabled = True
        Me.cmbRuleNames.Location = New System.Drawing.Point(3, 45)
        Me.cmbRuleNames.Name = "cmbRuleNames"
        Me.cmbRuleNames.Size = New System.Drawing.Size(270, 25)
        Me.cmbRuleNames.TabIndex = 5
        '
        'lblRuleRuleName
        '
        Me.lblRuleRuleName.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.lblRuleRuleName.AutoSize = True
        Me.lblRuleRuleName.Location = New System.Drawing.Point(3, 73)
        Me.lblRuleRuleName.Name = "lblRuleRuleName"
        Me.lblRuleRuleName.Size = New System.Drawing.Size(75, 17)
        Me.lblRuleRuleName.TabIndex = 6
        Me.lblRuleRuleName.Text = "Rule Name:"
        '
        'lblRuleSelectRule
        '
        Me.lblRuleSelectRule.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.lblRuleSelectRule.AutoSize = True
        Me.lblRuleSelectRule.Location = New System.Drawing.Point(3, 25)
        Me.lblRuleSelectRule.Name = "lblRuleSelectRule"
        Me.lblRuleSelectRule.Size = New System.Drawing.Size(74, 17)
        Me.lblRuleSelectRule.TabIndex = 4
        Me.lblRuleSelectRule.Text = "Select Rule:"
        '
        'flpLabelValues
        '
        Me.flpLabelValues.AutoSize = True
        Me.flpLabelValues.AutoSizeMode = System.Windows.Forms.AutoSizeMode.GrowAndShrink
        Me.flpLabelValues.Controls.Add(Me.pbValues)
        Me.flpLabelValues.Controls.Add(Me.lblValues)
        Me.flpLabelValues.Dock = System.Windows.Forms.DockStyle.Fill
        Me.flpLabelValues.Location = New System.Drawing.Point(3, 288)
        Me.flpLabelValues.Name = "flpLabelValues"
        Me.flpLabelValues.Size = New System.Drawing.Size(270, 22)
        Me.flpLabelValues.TabIndex = 18
        '
        'pbValues
        '
        Me.pbValues.Anchor = System.Windows.Forms.AnchorStyles.Left
        Me.pbValues.ErrorImage = Nothing
        Me.pbValues.Image = CType(resources.GetObject("pbValues.Image"), System.Drawing.Image)
        Me.pbValues.InitialImage = CType(resources.GetObject("pbValues.InitialImage"), System.Drawing.Image)
        Me.pbValues.Location = New System.Drawing.Point(3, 3)
        Me.pbValues.Name = "pbValues"
        Me.pbValues.Size = New System.Drawing.Size(16, 16)
        Me.pbValues.SizeMode = System.Windows.Forms.PictureBoxSizeMode.AutoSize
        Me.pbValues.TabIndex = 0
        Me.pbValues.TabStop = False
        '
        'lblValues
        '
        Me.lblValues.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.lblValues.AutoSize = True
        Me.lblValues.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.lblValues.Location = New System.Drawing.Point(25, 5)
        Me.lblValues.Name = "lblValues"
        Me.lblValues.Size = New System.Drawing.Size(45, 17)
        Me.lblValues.TabIndex = 12
        Me.lblValues.Text = "Values"
        Me.lblValues.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'flpLabelFilters
        '
        Me.flpLabelFilters.AutoSize = True
        Me.flpLabelFilters.AutoSizeMode = System.Windows.Forms.AutoSizeMode.GrowAndShrink
        Me.flpLabelFilters.Controls.Add(Me.pbFilter)
        Me.flpLabelFilters.Controls.Add(Me.lblFilters)
        Me.flpLabelFilters.Dock = System.Windows.Forms.DockStyle.Fill
        Me.flpLabelFilters.Location = New System.Drawing.Point(3, 470)
        Me.flpLabelFilters.Name = "flpLabelFilters"
        Me.flpLabelFilters.Size = New System.Drawing.Size(270, 22)
        Me.flpLabelFilters.TabIndex = 19
        '
        'pbFilter
        '
        Me.pbFilter.Anchor = System.Windows.Forms.AnchorStyles.Left
        Me.pbFilter.Image = CType(resources.GetObject("pbFilter.Image"), System.Drawing.Image)
        Me.pbFilter.Location = New System.Drawing.Point(3, 3)
        Me.pbFilter.Name = "pbFilter"
        Me.pbFilter.Size = New System.Drawing.Size(16, 16)
        Me.pbFilter.SizeMode = System.Windows.Forms.PictureBoxSizeMode.AutoSize
        Me.pbFilter.TabIndex = 0
        Me.pbFilter.TabStop = False
        '
        'lblFilters
        '
        Me.lblFilters.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.lblFilters.AutoSize = True
        Me.lblFilters.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.lblFilters.Location = New System.Drawing.Point(25, 5)
        Me.lblFilters.Name = "lblFilters"
        Me.lblFilters.Size = New System.Drawing.Size(42, 17)
        Me.lblFilters.TabIndex = 15
        Me.lblFilters.Text = "Filters"
        Me.lblFilters.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'flpRuleButtons
        '
        Me.flpRuleButtons.AutoSize = True
        Me.flpRuleButtons.AutoSizeMode = System.Windows.Forms.AutoSizeMode.GrowAndShrink
        Me.flpRuleButtons.Controls.Add(Me.btnRuleSave)
        Me.flpRuleButtons.Controls.Add(Me.btnRuleDelete)
        Me.flpRuleButtons.Dock = System.Windows.Forms.DockStyle.Right
        Me.flpRuleButtons.Location = New System.Drawing.Point(208, 667)
        Me.flpRuleButtons.Margin = New System.Windows.Forms.Padding(0)
        Me.flpRuleButtons.Name = "flpRuleButtons"
        Me.flpRuleButtons.Padding = New System.Windows.Forms.Padding(0, 0, 0, 4)
        Me.flpRuleButtons.Size = New System.Drawing.Size(68, 74)
        Me.flpRuleButtons.TabIndex = 17
        Me.flpRuleButtons.WrapContents = False
        '
        'btnRuleSave
        '
        Me.btnRuleSave.AutoSize = True
        Me.btnRuleSave.Image = CType(resources.GetObject("btnRuleSave.Image"), System.Drawing.Image)
        Me.btnRuleSave.Location = New System.Drawing.Point(1, 1)
        Me.btnRuleSave.Margin = New System.Windows.Forms.Padding(1)
        Me.btnRuleSave.Name = "btnRuleSave"
        Me.btnRuleSave.Size = New System.Drawing.Size(32, 32)
        Me.btnRuleSave.TabIndex = 16
        Me.btnRuleSave.UseVisualStyleBackColor = True
        '
        'btnRuleDelete
        '
        Me.btnRuleDelete.AutoSize = True
        Me.btnRuleDelete.Image = CType(resources.GetObject("btnRuleDelete.Image"), System.Drawing.Image)
        Me.btnRuleDelete.Location = New System.Drawing.Point(35, 1)
        Me.btnRuleDelete.Margin = New System.Windows.Forms.Padding(1)
        Me.btnRuleDelete.Name = "btnRuleDelete"
        Me.btnRuleDelete.Size = New System.Drawing.Size(32, 32)
        Me.btnRuleDelete.TabIndex = 17
        Me.btnRuleDelete.UseVisualStyleBackColor = True
        '
        'pnlRuleOptions
        '
        Me.pnlRuleOptions.AutoSize = True
        Me.pnlRuleOptions.AutoSizeMode = System.Windows.Forms.AutoSizeMode.GrowAndShrink
        Me.pnlRuleOptions.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.pnlRuleOptions.Controls.Add(Me.tlpRuleOptions)
        Me.pnlRuleOptions.Location = New System.Drawing.Point(3, 402)
        Me.pnlRuleOptions.Name = "pnlRuleOptions"
        Me.pnlRuleOptions.Padding = New System.Windows.Forms.Padding(3)
        Me.pnlRuleOptions.Size = New System.Drawing.Size(259, 62)
        Me.pnlRuleOptions.TabIndex = 20
        '
        'tlpRuleOptions
        '
        Me.tlpRuleOptions.AutoSize = True
        Me.tlpRuleOptions.AutoSizeMode = System.Windows.Forms.AutoSizeMode.GrowAndShrink
        Me.tlpRuleOptions.ColumnCount = 2
        Me.tlpRuleOptions.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 50.0!))
        Me.tlpRuleOptions.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 50.0!))
        Me.tlpRuleOptions.Controls.Add(Me.optRange, 0, 1)
        Me.tlpRuleOptions.Controls.Add(Me.optList, 1, 0)
        Me.tlpRuleOptions.Controls.Add(Me.optSingle, 0, 0)
        Me.tlpRuleOptions.Location = New System.Drawing.Point(4, 3)
        Me.tlpRuleOptions.Margin = New System.Windows.Forms.Padding(0)
        Me.tlpRuleOptions.Name = "tlpRuleOptions"
        Me.tlpRuleOptions.RowCount = 2
        Me.tlpRuleOptions.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 50.0!))
        Me.tlpRuleOptions.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 50.0!))
        Me.tlpRuleOptions.Size = New System.Drawing.Size(250, 54)
        Me.tlpRuleOptions.TabIndex = 4
        '
        'optRange
        '
        Me.optRange.AutoSize = True
        Me.optRange.Location = New System.Drawing.Point(3, 30)
        Me.optRange.Name = "optRange"
        Me.optRange.Size = New System.Drawing.Size(119, 21)
        Me.optRange.TabIndex = 3
        Me.optRange.TabStop = True
        Me.optRange.Text = "Range of values"
        Me.optRange.UseVisualStyleBackColor = True
        '
        'optList
        '
        Me.optList.AutoSize = True
        Me.optList.Location = New System.Drawing.Point(128, 3)
        Me.optList.Name = "optList"
        Me.optList.Size = New System.Drawing.Size(101, 21)
        Me.optList.TabIndex = 2
        Me.optList.TabStop = True
        Me.optList.Text = "List of values"
        Me.optList.UseVisualStyleBackColor = True
        '
        'optSingle
        '
        Me.optSingle.AutoSize = True
        Me.optSingle.Location = New System.Drawing.Point(3, 3)
        Me.optSingle.Name = "optSingle"
        Me.optSingle.Size = New System.Drawing.Size(95, 21)
        Me.optSingle.TabIndex = 1
        Me.optSingle.TabStop = True
        Me.optSingle.Text = "Single value"
        Me.optSingle.UseVisualStyleBackColor = True
        '
        'tabApply
        '
        Me.tabApply.Controls.Add(Me.pnlApply)
        Me.tabApply.Location = New System.Drawing.Point(4, 4)
        Me.tabApply.Margin = New System.Windows.Forms.Padding(4)
        Me.tabApply.Name = "tabApply"
        Me.tabApply.Padding = New System.Windows.Forms.Padding(4)
        Me.tabApply.Size = New System.Drawing.Size(292, 758)
        Me.tabApply.TabIndex = 1
        Me.tabApply.Text = "Apply"
        Me.tabApply.UseVisualStyleBackColor = True
        '
        'pnlApply
        '
        Me.pnlApply.AutoScroll = True
        Me.pnlApply.AutoSizeMode = System.Windows.Forms.AutoSizeMode.GrowAndShrink
        Me.pnlApply.Controls.Add(Me.tlpApply)
        Me.pnlApply.Dock = System.Windows.Forms.DockStyle.Fill
        Me.pnlApply.Location = New System.Drawing.Point(4, 4)
        Me.pnlApply.Name = "pnlApply"
        Me.pnlApply.Size = New System.Drawing.Size(284, 750)
        Me.pnlApply.TabIndex = 24
        '
        'tlpApply
        '
        Me.tlpApply.AutoSizeMode = System.Windows.Forms.AutoSizeMode.GrowAndShrink
        Me.tlpApply.ColumnCount = 1
        Me.tlpApply.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 100.0!))
        Me.tlpApply.Controls.Add(Me.cmbApplyNames, 0, 2)
        Me.tlpApply.Controls.Add(Me.txtApplyFilterExpression, 0, 11)
        Me.tlpApply.Controls.Add(Me.lvApplyFilters, 0, 10)
        Me.tlpApply.Controls.Add(Me.lblApplyFilters, 0, 9)
        Me.tlpApply.Controls.Add(Me.cmbApplyListSelectType, 0, 8)
        Me.tlpApply.Controls.Add(Me.lblApplyListSelectType, 0, 7)
        Me.tlpApply.Controls.Add(Me.cmbApplyRules, 0, 6)
        Me.tlpApply.Controls.Add(Me.lblApplySelectRule, 0, 5)
        Me.tlpApply.Controls.Add(Me.txtApplyName, 0, 4)
        Me.tlpApply.Controls.Add(Me.lblApplyName, 0, 3)
        Me.tlpApply.Controls.Add(Me.lblApplySelectApply, 0, 1)
        Me.tlpApply.Controls.Add(Me.lblApplyHeader, 0, 0)
        Me.tlpApply.Controls.Add(Me.flpApplyButtons, 0, 12)
        Me.tlpApply.Dock = System.Windows.Forms.DockStyle.Top
        Me.tlpApply.Location = New System.Drawing.Point(0, 0)
        Me.tlpApply.Name = "tlpApply"
        Me.tlpApply.RowCount = 13
        Me.tlpApply.RowStyles.Add(New System.Windows.Forms.RowStyle())
        Me.tlpApply.RowStyles.Add(New System.Windows.Forms.RowStyle())
        Me.tlpApply.RowStyles.Add(New System.Windows.Forms.RowStyle())
        Me.tlpApply.RowStyles.Add(New System.Windows.Forms.RowStyle())
        Me.tlpApply.RowStyles.Add(New System.Windows.Forms.RowStyle())
        Me.tlpApply.RowStyles.Add(New System.Windows.Forms.RowStyle())
        Me.tlpApply.RowStyles.Add(New System.Windows.Forms.RowStyle())
        Me.tlpApply.RowStyles.Add(New System.Windows.Forms.RowStyle())
        Me.tlpApply.RowStyles.Add(New System.Windows.Forms.RowStyle())
        Me.tlpApply.RowStyles.Add(New System.Windows.Forms.RowStyle())
        Me.tlpApply.RowStyles.Add(New System.Windows.Forms.RowStyle())
        Me.tlpApply.RowStyles.Add(New System.Windows.Forms.RowStyle())
        Me.tlpApply.RowStyles.Add(New System.Windows.Forms.RowStyle())
        Me.tlpApply.Size = New System.Drawing.Size(284, 496)
        Me.tlpApply.TabIndex = 0
        '
        'cmbApplyNames
        '
        Me.cmbApplyNames.Anchor = CType((System.Windows.Forms.AnchorStyles.Left Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.cmbApplyNames.FormattingEnabled = True
        Me.cmbApplyNames.Location = New System.Drawing.Point(3, 45)
        Me.cmbApplyNames.Name = "cmbApplyNames"
        Me.cmbApplyNames.Size = New System.Drawing.Size(278, 25)
        Me.cmbApplyNames.TabIndex = 22
        '
        'txtApplyFilterExpression
        '
        Me.txtApplyFilterExpression.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.txtApplyFilterExpression.BackColor = System.Drawing.SystemColors.Control
        Me.txtApplyFilterExpression.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.txtApplyFilterExpression.Location = New System.Drawing.Point(3, 334)
        Me.txtApplyFilterExpression.Name = "txtApplyFilterExpression"
        Me.txtApplyFilterExpression.ReadOnly = True
        Me.txtApplyFilterExpression.ScrollBars = System.Windows.Forms.RichTextBoxScrollBars.ForcedVertical
        Me.txtApplyFilterExpression.Size = New System.Drawing.Size(278, 80)
        Me.txtApplyFilterExpression.TabIndex = 31
        Me.txtApplyFilterExpression.Text = ""
        '
        'lvApplyFilters
        '
        Me.lvApplyFilters.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.lvApplyFilters.Columns.AddRange(New System.Windows.Forms.ColumnHeader() {Me.ApplyFilter})
        Me.lvApplyFilters.FullRowSelect = True
        Me.lvApplyFilters.HeaderStyle = System.Windows.Forms.ColumnHeaderStyle.None
        Me.lvApplyFilters.HideSelection = False
        Me.lvApplyFilters.Location = New System.Drawing.Point(3, 248)
        Me.lvApplyFilters.MultiSelect = False
        Me.lvApplyFilters.Name = "lvApplyFilters"
        Me.lvApplyFilters.Size = New System.Drawing.Size(278, 80)
        Me.lvApplyFilters.TabIndex = 30
        Me.lvApplyFilters.UseCompatibleStateImageBehavior = False
        Me.lvApplyFilters.View = System.Windows.Forms.View.Details
        '
        'lblApplyFilters
        '
        Me.lblApplyFilters.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.lblApplyFilters.Image = Global.ResourceManagement.My.Resources.Resources._16filter
        Me.lblApplyFilters.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.lblApplyFilters.Location = New System.Drawing.Point(3, 223)
        Me.lblApplyFilters.Name = "lblApplyFilters"
        Me.lblApplyFilters.Size = New System.Drawing.Size(150, 22)
        Me.lblApplyFilters.TabIndex = 29
        Me.lblApplyFilters.Text = "Filters Parameters:"
        Me.lblApplyFilters.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'cmbApplyListSelectType
        '
        Me.cmbApplyListSelectType.Anchor = CType((System.Windows.Forms.AnchorStyles.Left Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.cmbApplyListSelectType.FormattingEnabled = True
        Me.cmbApplyListSelectType.Location = New System.Drawing.Point(4, 194)
        Me.cmbApplyListSelectType.Margin = New System.Windows.Forms.Padding(4)
        Me.cmbApplyListSelectType.Name = "cmbApplyListSelectType"
        Me.cmbApplyListSelectType.Size = New System.Drawing.Size(276, 25)
        Me.cmbApplyListSelectType.TabIndex = 28
        '
        'lblApplyListSelectType
        '
        Me.lblApplyListSelectType.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.lblApplyListSelectType.AutoSize = True
        Me.lblApplyListSelectType.Location = New System.Drawing.Point(4, 173)
        Me.lblApplyListSelectType.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.lblApplyListSelectType.Name = "lblApplyListSelectType"
        Me.lblApplyListSelectType.Size = New System.Drawing.Size(99, 17)
        Me.lblApplyListSelectType.TabIndex = 27
        Me.lblApplyListSelectType.Text = "List Select Type:"
        '
        'cmbApplyRules
        '
        Me.cmbApplyRules.Anchor = CType((System.Windows.Forms.AnchorStyles.Left Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.cmbApplyRules.FormattingEnabled = True
        Me.cmbApplyRules.Location = New System.Drawing.Point(4, 144)
        Me.cmbApplyRules.Margin = New System.Windows.Forms.Padding(4)
        Me.cmbApplyRules.Name = "cmbApplyRules"
        Me.cmbApplyRules.Size = New System.Drawing.Size(276, 25)
        Me.cmbApplyRules.TabIndex = 26
        '
        'lblApplySelectRule
        '
        Me.lblApplySelectRule.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.lblApplySelectRule.AutoSize = True
        Me.lblApplySelectRule.Location = New System.Drawing.Point(3, 123)
        Me.lblApplySelectRule.Name = "lblApplySelectRule"
        Me.lblApplySelectRule.Size = New System.Drawing.Size(74, 17)
        Me.lblApplySelectRule.TabIndex = 25
        Me.lblApplySelectRule.Text = "Select Rule:"
        '
        'txtApplyName
        '
        Me.txtApplyName.Anchor = CType((System.Windows.Forms.AnchorStyles.Left Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.txtApplyName.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.txtApplyName.Location = New System.Drawing.Point(4, 94)
        Me.txtApplyName.Margin = New System.Windows.Forms.Padding(4)
        Me.txtApplyName.Name = "txtApplyName"
        Me.txtApplyName.Size = New System.Drawing.Size(276, 25)
        Me.txtApplyName.TabIndex = 24
        '
        'lblApplyName
        '
        Me.lblApplyName.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.lblApplyName.AutoSize = True
        Me.lblApplyName.Location = New System.Drawing.Point(3, 73)
        Me.lblApplyName.Name = "lblApplyName"
        Me.lblApplyName.Size = New System.Drawing.Size(83, 17)
        Me.lblApplyName.TabIndex = 23
        Me.lblApplyName.Text = "Apply Name:"
        '
        'lblApplySelectApply
        '
        Me.lblApplySelectApply.AutoSize = True
        Me.lblApplySelectApply.Location = New System.Drawing.Point(3, 25)
        Me.lblApplySelectApply.Name = "lblApplySelectApply"
        Me.lblApplySelectApply.Size = New System.Drawing.Size(82, 17)
        Me.lblApplySelectApply.TabIndex = 21
        Me.lblApplySelectApply.Text = "Select Apply:"
        '
        'lblApplyHeader
        '
        Me.lblApplyHeader.AutoSize = True
        Me.lblApplyHeader.Font = New System.Drawing.Font("Segoe UI", 14.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblApplyHeader.Location = New System.Drawing.Point(3, 0)
        Me.lblApplyHeader.Name = "lblApplyHeader"
        Me.lblApplyHeader.Size = New System.Drawing.Size(135, 25)
        Me.lblApplyHeader.TabIndex = 9
        Me.lblApplyHeader.Text = "Apply List Rule"
        '
        'flpApplyButtons
        '
        Me.flpApplyButtons.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.flpApplyButtons.AutoSize = True
        Me.flpApplyButtons.Controls.Add(Me.btnApplySave)
        Me.flpApplyButtons.Controls.Add(Me.btnApplyDelete)
        Me.flpApplyButtons.Location = New System.Drawing.Point(213, 420)
        Me.flpApplyButtons.Name = "flpApplyButtons"
        Me.flpApplyButtons.Size = New System.Drawing.Size(68, 34)
        Me.flpApplyButtons.TabIndex = 32
        '
        'btnApplySave
        '
        Me.btnApplySave.Image = Global.ResourceManagement.My.Resources.Resources._24save
        Me.btnApplySave.Location = New System.Drawing.Point(1, 1)
        Me.btnApplySave.Margin = New System.Windows.Forms.Padding(1)
        Me.btnApplySave.Name = "btnApplySave"
        Me.btnApplySave.Size = New System.Drawing.Size(32, 32)
        Me.btnApplySave.TabIndex = 20
        Me.btnApplySave.UseVisualStyleBackColor = True
        '
        'btnApplyDelete
        '
        Me.btnApplyDelete.Image = Global.ResourceManagement.My.Resources.Resources._24delete
        Me.btnApplyDelete.Location = New System.Drawing.Point(35, 1)
        Me.btnApplyDelete.Margin = New System.Windows.Forms.Padding(1)
        Me.btnApplyDelete.Name = "btnApplyDelete"
        Me.btnApplyDelete.Size = New System.Drawing.Size(32, 32)
        Me.btnApplyDelete.TabIndex = 21
        Me.btnApplyDelete.UseVisualStyleBackColor = True
        '
        'ExcelRuleDesigner
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(7.0!, 17.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.Controls.Add(Me.tabPane)
        Me.Font = New System.Drawing.Font("Segoe UI", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Margin = New System.Windows.Forms.Padding(4)
        Me.Name = "ExcelRuleDesigner"
        Me.Size = New System.Drawing.Size(300, 788)
        Me.tabPane.ResumeLayout(False)
        Me.tabRules.ResumeLayout(False)
        Me.pnlRule.ResumeLayout(False)
        Me.tlpRule.ResumeLayout(False)
        Me.tlpRule.PerformLayout()
        Me.flpLabelValues.ResumeLayout(False)
        Me.flpLabelValues.PerformLayout()
        CType(Me.pbValues, System.ComponentModel.ISupportInitialize).EndInit()
        Me.flpLabelFilters.ResumeLayout(False)
        Me.flpLabelFilters.PerformLayout()
        CType(Me.pbFilter, System.ComponentModel.ISupportInitialize).EndInit()
        Me.flpRuleButtons.ResumeLayout(False)
        Me.flpRuleButtons.PerformLayout()
        Me.pnlRuleOptions.ResumeLayout(False)
        Me.pnlRuleOptions.PerformLayout()
        Me.tlpRuleOptions.ResumeLayout(False)
        Me.tlpRuleOptions.PerformLayout()
        Me.tabApply.ResumeLayout(False)
        Me.pnlApply.ResumeLayout(False)
        Me.tlpApply.ResumeLayout(False)
        Me.tlpApply.PerformLayout()
        Me.flpApplyButtons.ResumeLayout(False)
        Me.ResumeLayout(False)

    End Sub

    Friend WithEvents tabPane As Windows.Forms.TabControl
  Friend WithEvents tabRules As Windows.Forms.TabPage
  Friend WithEvents tabApply As Windows.Forms.TabPage
  Friend WithEvents pnlRule As Windows.Forms.Panel
  Friend WithEvents tlpRule As Windows.Forms.TableLayoutPanel
  Friend WithEvents txtRuleFilterExpression As Windows.Forms.RichTextBox
  Friend WithEvents lvRuleFilters As Windows.Forms.ListView
  Friend WithEvents RuleFilter As Windows.Forms.ColumnHeader
    Friend WithEvents lvRuleValues As Windows.Forms.ListView
    Friend WithEvents RuleField As Windows.Forms.ColumnHeader
    Friend WithEvents lblRuleDragFields As Windows.Forms.Label
    Friend WithEvents tvRuleFields As Windows.Forms.TreeView
    Friend WithEvents lblRuleChooseFields As Windows.Forms.Label
    Friend WithEvents txtRuleName As Windows.Forms.TextBox
    Friend WithEvents lblRuleHeader As Windows.Forms.Label
    Friend WithEvents cmbRuleNames As Windows.Forms.ComboBox
    Friend WithEvents lblRuleRuleName As Windows.Forms.Label
    Friend WithEvents lblRuleSelectRule As Windows.Forms.Label
    Friend WithEvents flpLabelValues As Windows.Forms.FlowLayoutPanel
    Friend WithEvents pbValues As Windows.Forms.PictureBox
    Friend WithEvents lblValues As Windows.Forms.Label
    Friend WithEvents flpLabelFilters As Windows.Forms.FlowLayoutPanel
    Friend WithEvents pbFilter As Windows.Forms.PictureBox
    Friend WithEvents lblFilters As Windows.Forms.Label
    Friend WithEvents pnlApply As Windows.Forms.Panel
    Friend WithEvents tlpApply As Windows.Forms.TableLayoutPanel
    Friend WithEvents cmbApplyRules As Windows.Forms.ComboBox
    Friend WithEvents lblApplySelectRule As Windows.Forms.Label
    Friend WithEvents txtApplyName As Windows.Forms.TextBox
    Friend WithEvents lblApplyName As Windows.Forms.Label
    Friend WithEvents cmbApplyNames As Windows.Forms.ComboBox
    Friend WithEvents lblApplySelectApply As Windows.Forms.Label
    Friend WithEvents lblApplyHeader As Windows.Forms.Label
    Friend WithEvents txtApplyFilterExpression As Windows.Forms.RichTextBox
    Friend WithEvents lvApplyFilters As Windows.Forms.ListView
    Friend WithEvents ApplyFilter As Windows.Forms.ColumnHeader
    Friend WithEvents lblApplyFilters As Windows.Forms.Label
    Friend WithEvents cmbApplyListSelectType As Windows.Forms.ComboBox
    Friend WithEvents lblApplyListSelectType As Windows.Forms.Label
    Friend WithEvents flpApplyButtons As Windows.Forms.FlowLayoutPanel
    Friend WithEvents btnApplySave As Windows.Forms.Button
    Friend WithEvents btnApplyDelete As Windows.Forms.Button
    Friend WithEvents flpRuleButtons As Windows.Forms.FlowLayoutPanel
    Friend WithEvents btnRuleSave As Windows.Forms.Button
    Friend WithEvents btnRuleDelete As Windows.Forms.Button
    Friend WithEvents pnlRuleOptions As Windows.Forms.Panel
    Friend WithEvents tlpRuleOptions As Windows.Forms.TableLayoutPanel
    Friend WithEvents optRange As Windows.Forms.RadioButton
    Friend WithEvents optList As Windows.Forms.RadioButton
    Friend WithEvents optSingle As Windows.Forms.RadioButton
End Class
