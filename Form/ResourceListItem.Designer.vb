<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class ResourceListItem
  Inherits System.Windows.Forms.Form

  'Form overrides dispose to clean up the component list.
  <System.Diagnostics.DebuggerNonUserCode()> _
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
  <System.Diagnostics.DebuggerStepThrough()> _
  Private Sub InitializeComponent()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(ResourceListItem))
        Me.cmbListItemTypes = New System.Windows.Forms.ComboBox()
        Me.txtResourceListItemName = New System.Windows.Forms.TextBox()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.lstResourceListItems = New System.Windows.Forms.ListView()
        Me.btnClose = New System.Windows.Forms.Button()
        Me.btnUpdate = New System.Windows.Forms.Button()
        Me.btnAddNew = New System.Windows.Forms.Button()
        Me.btnDelete = New System.Windows.Forms.Button()
        Me.cmbValueTypes = New System.Windows.Forms.ComboBox()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.SuspendLayout()
        '
        'cmbListItemTypes
        '
        Me.cmbListItemTypes.FormattingEnabled = True
        Me.cmbListItemTypes.Location = New System.Drawing.Point(351, 34)
        Me.cmbListItemTypes.Margin = New System.Windows.Forms.Padding(4)
        Me.cmbListItemTypes.Name = "cmbListItemTypes"
        Me.cmbListItemTypes.Size = New System.Drawing.Size(207, 25)
        Me.cmbListItemTypes.TabIndex = 0
        '
        'txtResourceListItemName
        '
        Me.txtResourceListItemName.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.txtResourceListItemName.Location = New System.Drawing.Point(13, 34)
        Me.txtResourceListItemName.Margin = New System.Windows.Forms.Padding(5)
        Me.txtResourceListItemName.Name = "txtResourceListItemName"
        Me.txtResourceListItemName.Size = New System.Drawing.Size(207, 25)
        Me.txtResourceListItemName.TabIndex = 9
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(13, 12)
        Me.Label1.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(148, 17)
        Me.Label1.TabIndex = 13
        Me.Label1.Text = "Resouce List Item Name"
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(348, 13)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(87, 17)
        Me.Label2.TabIndex = 14
        Me.Label2.Text = "List Item Type"
        '
        'lstResourceListItems
        '
        Me.lstResourceListItems.HideSelection = False
        Me.lstResourceListItems.Location = New System.Drawing.Point(13, 66)
        Me.lstResourceListItems.Name = "lstResourceListItems"
        Me.lstResourceListItems.Size = New System.Drawing.Size(545, 123)
        Me.lstResourceListItems.TabIndex = 15
        Me.lstResourceListItems.UseCompatibleStateImageBehavior = False
        '
        'btnClose
        '
        Me.btnClose.Image = Global.ResourceManagement.My.Resources.Resources._24close
        Me.btnClose.Location = New System.Drawing.Point(565, 200)
        Me.btnClose.Margin = New System.Windows.Forms.Padding(5)
        Me.btnClose.Name = "btnClose"
        Me.btnClose.Size = New System.Drawing.Size(32, 32)
        Me.btnClose.TabIndex = 19
        Me.btnClose.UseVisualStyleBackColor = True
        '
        'btnUpdate
        '
        Me.btnUpdate.Image = Global.ResourceManagement.My.Resources.Resources._24updatecommit
        Me.btnUpdate.Location = New System.Drawing.Point(565, 109)
        Me.btnUpdate.Margin = New System.Windows.Forms.Padding(5)
        Me.btnUpdate.Name = "btnUpdate"
        Me.btnUpdate.Size = New System.Drawing.Size(32, 32)
        Me.btnUpdate.TabIndex = 17
        Me.btnUpdate.UseVisualStyleBackColor = True
        '
        'btnAddNew
        '
        Me.btnAddNew.Image = Global.ResourceManagement.My.Resources.Resources._24new
        Me.btnAddNew.Location = New System.Drawing.Point(565, 67)
        Me.btnAddNew.Margin = New System.Windows.Forms.Padding(5)
        Me.btnAddNew.Name = "btnAddNew"
        Me.btnAddNew.Size = New System.Drawing.Size(32, 32)
        Me.btnAddNew.TabIndex = 16
        Me.btnAddNew.UseVisualStyleBackColor = True
        '
        'btnDelete
        '
        Me.btnDelete.Image = Global.ResourceManagement.My.Resources.Resources._24delete
        Me.btnDelete.Location = New System.Drawing.Point(565, 151)
        Me.btnDelete.Margin = New System.Windows.Forms.Padding(5)
        Me.btnDelete.Name = "btnDelete"
        Me.btnDelete.Size = New System.Drawing.Size(32, 32)
        Me.btnDelete.TabIndex = 18
        Me.btnDelete.UseVisualStyleBackColor = True
        '
        'cmbValueTypes
        '
        Me.cmbValueTypes.FormattingEnabled = True
        Me.cmbValueTypes.Location = New System.Drawing.Point(229, 34)
        Me.cmbValueTypes.Margin = New System.Windows.Forms.Padding(4)
        Me.cmbValueTypes.Name = "cmbValueTypes"
        Me.cmbValueTypes.Size = New System.Drawing.Size(114, 25)
        Me.cmbValueTypes.TabIndex = 20
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Location = New System.Drawing.Point(226, 13)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(70, 17)
        Me.Label3.TabIndex = 21
        Me.Label3.Text = "Value Type"
        '
        'ResourceListItem
        '
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.None
        Me.ClientSize = New System.Drawing.Size(605, 242)
        Me.Controls.Add(Me.cmbValueTypes)
        Me.Controls.Add(Me.btnClose)
        Me.Controls.Add(Me.btnUpdate)
        Me.Controls.Add(Me.btnAddNew)
        Me.Controls.Add(Me.btnDelete)
        Me.Controls.Add(Me.lstResourceListItems)
        Me.Controls.Add(Me.txtResourceListItemName)
        Me.Controls.Add(Me.cmbListItemTypes)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.Label3)
        Me.Font = New System.Drawing.Font("Segoe UI", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.Margin = New System.Windows.Forms.Padding(4)
        Me.Name = "ResourceListItem"
        Me.SizeGripStyle = System.Windows.Forms.SizeGripStyle.Hide
        Me.Text = "Resource Name / Value"
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

    Friend WithEvents cmbListItemTypes As Windows.Forms.ComboBox
    Friend WithEvents txtResourceListItemName As Windows.Forms.TextBox
    Friend WithEvents Label1 As Windows.Forms.Label
    Friend WithEvents Label2 As Windows.Forms.Label
    Friend WithEvents lstResourceListItems As Windows.Forms.ListView
    Friend WithEvents btnClose As Windows.Forms.Button
    Friend WithEvents btnUpdate As Windows.Forms.Button
    Friend WithEvents btnAddNew As Windows.Forms.Button
    Friend WithEvents btnDelete As Windows.Forms.Button
    Friend WithEvents cmbValueTypes As Windows.Forms.ComboBox
    Friend WithEvents Label3 As Windows.Forms.Label
End Class
