<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()>
Partial Class ResourceManager
  Inherits System.Windows.Forms.Form

  'Form overrides dispose to clean up the component list.
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
    Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(ResourceManager))
    Me.lstResourceAvailability = New System.Windows.Forms.ListBox()
    Me.GroupBox1 = New System.Windows.Forms.GroupBox()
    Me.GroupBox2 = New System.Windows.Forms.GroupBox()
    Me.optFullName = New System.Windows.Forms.RadioButton()
    Me.optPreferredName = New System.Windows.Forms.RadioButton()
    Me.chkShowInactive = New System.Windows.Forms.CheckBox()
    Me.PictureBox1 = New System.Windows.Forms.PictureBox()
    Me.txtFilter = New System.Windows.Forms.TextBox()
    Me.btnNewResource = New System.Windows.Forms.Button()
    Me.btnEditResource = New System.Windows.Forms.Button()
    Me.btnNewAvailability = New System.Windows.Forms.Button()
    Me.btnEditAvailability = New System.Windows.Forms.Button()
    Me.btnClose = New System.Windows.Forms.Button()
    Me.lblDefaultMode = New System.Windows.Forms.Label()
    Me.lstResources = New ResourceManagement.CustomControls.ExtendedListView()
    Me.GroupBox1.SuspendLayout()
        Me.GroupBox2.SuspendLayout()
        CType(Me.PictureBox1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'lstResourceAvailability
        '
        Me.lstResourceAvailability.FormattingEnabled = True
        Me.lstResourceAvailability.ItemHeight = 17
        Me.lstResourceAvailability.Location = New System.Drawing.Point(12, 334)
        Me.lstResourceAvailability.Margin = New System.Windows.Forms.Padding(4)
        Me.lstResourceAvailability.Name = "lstResourceAvailability"
        Me.lstResourceAvailability.Size = New System.Drawing.Size(461, 157)
        Me.lstResourceAvailability.TabIndex = 1
        '
        'GroupBox1
        '
        Me.GroupBox1.Controls.Add(Me.GroupBox2)
        Me.GroupBox1.Controls.Add(Me.chkShowInactive)
        Me.GroupBox1.Controls.Add(Me.PictureBox1)
        Me.GroupBox1.Controls.Add(Me.txtFilter)
        Me.GroupBox1.Location = New System.Drawing.Point(12, 12)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Size = New System.Drawing.Size(459, 104)
        Me.GroupBox1.TabIndex = 3
        Me.GroupBox1.TabStop = False
        '
        'GroupBox2
        '
        Me.GroupBox2.Controls.Add(Me.optFullName)
        Me.GroupBox2.Controls.Add(Me.optPreferredName)
        Me.GroupBox2.Location = New System.Drawing.Point(35, 24)
        Me.GroupBox2.Name = "GroupBox2"
        Me.GroupBox2.Size = New System.Drawing.Size(220, 41)
        Me.GroupBox2.TabIndex = 6
        Me.GroupBox2.TabStop = False
        '
        'optFullName
        '
        Me.optFullName.AutoSize = True
        Me.optFullName.Location = New System.Drawing.Point(132, 15)
        Me.optFullName.Name = "optFullName"
        Me.optFullName.Size = New System.Drawing.Size(84, 21)
        Me.optFullName.TabIndex = 4
        Me.optFullName.TabStop = True
        Me.optFullName.Text = "Full Name"
        Me.optFullName.UseVisualStyleBackColor = True
        '
        'optPreferredName
        '
        Me.optPreferredName.AutoSize = True
        Me.optPreferredName.Location = New System.Drawing.Point(6, 15)
        Me.optPreferredName.Name = "optPreferredName"
        Me.optPreferredName.Size = New System.Drawing.Size(120, 21)
        Me.optPreferredName.TabIndex = 3
        Me.optPreferredName.TabStop = True
        Me.optPreferredName.Text = "Preferred Name"
        Me.optPreferredName.UseVisualStyleBackColor = True
        '
        'chkShowInactive
        '
        Me.chkShowInactive.AutoSize = True
        Me.chkShowInactive.Location = New System.Drawing.Point(261, 39)
        Me.chkShowInactive.Name = "chkShowInactive"
        Me.chkShowInactive.Size = New System.Drawing.Size(105, 21)
        Me.chkShowInactive.TabIndex = 5
        Me.chkShowInactive.Text = "Show Inactive"
        Me.chkShowInactive.UseVisualStyleBackColor = True
        '
        'PictureBox1
        '
        Me.PictureBox1.Image = Global.ResourceManagement.My.Resources.Resources._16filter
        Me.PictureBox1.Location = New System.Drawing.Point(6, 15)
        Me.PictureBox1.Name = "PictureBox1"
        Me.PictureBox1.Size = New System.Drawing.Size(18, 19)
        Me.PictureBox1.SizeMode = System.Windows.Forms.PictureBoxSizeMode.CenterImage
        Me.PictureBox1.TabIndex = 4
        Me.PictureBox1.TabStop = False
        '
        'txtFilter
        '
        Me.txtFilter.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.txtFilter.Location = New System.Drawing.Point(35, 71)
        Me.txtFilter.Name = "txtFilter"
        Me.txtFilter.Size = New System.Drawing.Size(212, 25)
        Me.txtFilter.TabIndex = 3
        '
        'btnNewResource
        '
        Me.btnNewResource.Image = Global.ResourceManagement.My.Resources.Resources._24new
        Me.btnNewResource.Location = New System.Drawing.Point(400, 294)
        Me.btnNewResource.Margin = New System.Windows.Forms.Padding(4)
        Me.btnNewResource.Name = "btnNewResource"
        Me.btnNewResource.Size = New System.Drawing.Size(32, 32)
        Me.btnNewResource.TabIndex = 5
        Me.btnNewResource.UseVisualStyleBackColor = True
        '
        'btnEditResource
        '
        Me.btnEditResource.Image = Global.ResourceManagement.My.Resources.Resources._24edit
        Me.btnEditResource.Location = New System.Drawing.Point(440, 294)
        Me.btnEditResource.Margin = New System.Windows.Forms.Padding(4)
        Me.btnEditResource.Name = "btnEditResource"
        Me.btnEditResource.Size = New System.Drawing.Size(32, 32)
        Me.btnEditResource.TabIndex = 6
        Me.btnEditResource.UseVisualStyleBackColor = True
        '
        'btnNewAvailability
        '
        Me.btnNewAvailability.Image = Global.ResourceManagement.My.Resources.Resources._24new
        Me.btnNewAvailability.Location = New System.Drawing.Point(401, 499)
        Me.btnNewAvailability.Margin = New System.Windows.Forms.Padding(4)
        Me.btnNewAvailability.Name = "btnNewAvailability"
        Me.btnNewAvailability.Size = New System.Drawing.Size(32, 32)
        Me.btnNewAvailability.TabIndex = 7
        Me.btnNewAvailability.UseVisualStyleBackColor = True
        '
        'btnEditAvailability
        '
        Me.btnEditAvailability.Image = Global.ResourceManagement.My.Resources.Resources._24edit
        Me.btnEditAvailability.Location = New System.Drawing.Point(441, 499)
        Me.btnEditAvailability.Margin = New System.Windows.Forms.Padding(4)
        Me.btnEditAvailability.Name = "btnEditAvailability"
        Me.btnEditAvailability.Size = New System.Drawing.Size(32, 32)
        Me.btnEditAvailability.TabIndex = 8
        Me.btnEditAvailability.UseVisualStyleBackColor = True
        '
        'btnClose
        '
        Me.btnClose.Image = Global.ResourceManagement.My.Resources.Resources._24close
        Me.btnClose.Location = New System.Drawing.Point(440, 539)
        Me.btnClose.Margin = New System.Windows.Forms.Padding(4)
        Me.btnClose.Name = "btnClose"
        Me.btnClose.Size = New System.Drawing.Size(32, 32)
        Me.btnClose.TabIndex = 11
        Me.btnClose.UseVisualStyleBackColor = True
        '
        'lblDefaultMode
        '
        Me.lblDefaultMode.AutoSize = True
        Me.lblDefaultMode.Location = New System.Drawing.Point(12, 589)
        Me.lblDefaultMode.Name = "lblDefaultMode"
        Me.lblDefaultMode.Size = New System.Drawing.Size(0, 17)
        Me.lblDefaultMode.TabIndex = 12
        '
        'lstResources
        '
        Me.lstResources.ColumnPercents = Nothing
        Me.lstResources.FullRowSelect = True
        Me.lstResources.HideSelection = False
        Me.lstResources.Location = New System.Drawing.Point(12, 125)
        Me.lstResources.Name = "lstResources"
        Me.lstResources.Size = New System.Drawing.Size(459, 161)
        Me.lstResources.TabIndex = 13
        Me.lstResources.UseCompatibleStateImageBehavior = False
        '
        'ResourceManager
        '
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.None
        Me.ClientSize = New System.Drawing.Size(483, 581)
        Me.Controls.Add(Me.lstResources)
        Me.Controls.Add(Me.lblDefaultMode)
        Me.Controls.Add(Me.btnClose)
        Me.Controls.Add(Me.btnNewAvailability)
        Me.Controls.Add(Me.btnEditAvailability)
        Me.Controls.Add(Me.btnNewResource)
        Me.Controls.Add(Me.btnEditResource)
        Me.Controls.Add(Me.GroupBox1)
        Me.Controls.Add(Me.lstResourceAvailability)
        Me.Font = New System.Drawing.Font("Segoe UI", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.Margin = New System.Windows.Forms.Padding(4)
        Me.Name = "ResourceManager"
        Me.SizeGripStyle = System.Windows.Forms.SizeGripStyle.Hide
        Me.Text = "Resource Manager"
        Me.GroupBox1.ResumeLayout(False)
        Me.GroupBox1.PerformLayout()
        Me.GroupBox2.ResumeLayout(False)
        Me.GroupBox2.PerformLayout()
        CType(Me.PictureBox1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents lstResourceAvailability As Windows.Forms.ListBox
    Friend WithEvents GroupBox1 As Windows.Forms.GroupBox
    Friend WithEvents txtFilter As Windows.Forms.TextBox
    Friend WithEvents PictureBox1 As Windows.Forms.PictureBox
    Friend WithEvents btnNewResource As Windows.Forms.Button
    Friend WithEvents btnEditResource As Windows.Forms.Button
    Friend WithEvents btnNewAvailability As Windows.Forms.Button
    Friend WithEvents btnEditAvailability As Windows.Forms.Button
    Friend WithEvents GroupBox2 As Windows.Forms.GroupBox
    Friend WithEvents optFullName As Windows.Forms.RadioButton
    Friend WithEvents optPreferredName As Windows.Forms.RadioButton
    Friend WithEvents chkShowInactive As Windows.Forms.CheckBox
    Friend WithEvents btnClose As Windows.Forms.Button
    Friend WithEvents lblDefaultMode As Windows.Forms.Label
  Friend WithEvents lstResources As ResourceManagement.CustomControls.ExtendedListView
End Class
