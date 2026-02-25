<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()>
Partial Class ListItemTypeAndItem
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
    Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(ListItemTypeAndItem))
    Me.btnClose = New System.Windows.Forms.Button()
    Me.btnUpdateListItem = New System.Windows.Forms.Button()
    Me.btnAddNewListItem = New System.Windows.Forms.Button()
    Me.btnDeleteListItem = New System.Windows.Forms.Button()
    Me.lstListItems = New System.Windows.Forms.ListBox()
    Me.txtListItemName = New System.Windows.Forms.TextBox()
    Me.lstListItemTypes = New System.Windows.Forms.ListBox()
    Me.txtListItemTypeName = New System.Windows.Forms.TextBox()
    Me.btnUpdateListItemType = New System.Windows.Forms.Button()
    Me.btnAddNewListItemType = New System.Windows.Forms.Button()
    Me.btnDeleteListItemType = New System.Windows.Forms.Button()
    Me.Label1 = New System.Windows.Forms.Label()
    Me.Label2 = New System.Windows.Forms.Label()
        Me.SuspendLayout()
        '
        'btnClose
        '
        Me.btnClose.Image = Global.ResourceManagement.My.Resources.Resources._24close
        Me.btnClose.Location = New System.Drawing.Point(531, 204)
        Me.btnClose.Margin = New System.Windows.Forms.Padding(5)
        Me.btnClose.Name = "btnClose"
        Me.btnClose.Size = New System.Drawing.Size(32, 32)
        Me.btnClose.TabIndex = 6
        Me.btnClose.UseVisualStyleBackColor = True
        '
        'btnUpdateListItem
        '
        Me.btnUpdateListItem.Image = Global.ResourceManagement.My.Resources.Resources._24updatecommit
        Me.btnUpdateListItem.Location = New System.Drawing.Point(531, 113)
        Me.btnUpdateListItem.Margin = New System.Windows.Forms.Padding(5)
        Me.btnUpdateListItem.Name = "btnUpdateListItem"
        Me.btnUpdateListItem.Size = New System.Drawing.Size(32, 32)
        Me.btnUpdateListItem.TabIndex = 4
        Me.btnUpdateListItem.UseVisualStyleBackColor = True
        '
        'btnAddNewListItem
        '
        Me.btnAddNewListItem.Image = Global.ResourceManagement.My.Resources.Resources._24new
        Me.btnAddNewListItem.Location = New System.Drawing.Point(531, 71)
        Me.btnAddNewListItem.Margin = New System.Windows.Forms.Padding(5)
        Me.btnAddNewListItem.Name = "btnAddNewListItem"
        Me.btnAddNewListItem.Size = New System.Drawing.Size(32, 32)
        Me.btnAddNewListItem.TabIndex = 3
        Me.btnAddNewListItem.UseVisualStyleBackColor = True
        '
        'btnDeleteListItem
        '
        Me.btnDeleteListItem.Image = Global.ResourceManagement.My.Resources.Resources._24delete
        Me.btnDeleteListItem.Location = New System.Drawing.Point(531, 155)
        Me.btnDeleteListItem.Margin = New System.Windows.Forms.Padding(5)
        Me.btnDeleteListItem.Name = "btnDeleteListItem"
        Me.btnDeleteListItem.Size = New System.Drawing.Size(32, 32)
        Me.btnDeleteListItem.TabIndex = 5
        Me.btnDeleteListItem.UseVisualStyleBackColor = True
        '
        'lstListItems
        '
        Me.lstListItems.FormattingEnabled = True
        Me.lstListItems.ItemHeight = 17
        Me.lstListItems.Location = New System.Drawing.Point(293, 71)
        Me.lstListItems.Margin = New System.Windows.Forms.Padding(4)
        Me.lstListItems.Name = "lstListItems"
        Me.lstListItems.Size = New System.Drawing.Size(229, 123)
        Me.lstListItems.TabIndex = 2
        '
        'txtListItemName
        '
        Me.txtListItemName.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.txtListItemName.Location = New System.Drawing.Point(293, 38)
        Me.txtListItemName.Margin = New System.Windows.Forms.Padding(4)
        Me.txtListItemName.Name = "txtListItemName"
        Me.txtListItemName.Size = New System.Drawing.Size(229, 25)
        Me.txtListItemName.TabIndex = 1
        '
        'lstListItemTypes
        '
        Me.lstListItemTypes.FormattingEnabled = True
        Me.lstListItemTypes.ItemHeight = 17
        Me.lstListItemTypes.Location = New System.Drawing.Point(14, 71)
        Me.lstListItemTypes.Margin = New System.Windows.Forms.Padding(4)
        Me.lstListItemTypes.Name = "lstListItemTypes"
        Me.lstListItemTypes.Size = New System.Drawing.Size(229, 123)
        Me.lstListItemTypes.TabIndex = 7
        '
        'txtListItemTypeName
        '
        Me.txtListItemTypeName.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.txtListItemTypeName.Location = New System.Drawing.Point(13, 38)
        Me.txtListItemTypeName.Margin = New System.Windows.Forms.Padding(4)
        Me.txtListItemTypeName.Name = "txtListItemTypeName"
        Me.txtListItemTypeName.Size = New System.Drawing.Size(229, 25)
        Me.txtListItemTypeName.TabIndex = 8
        '
        'btnUpdateListItemType
        '
        Me.btnUpdateListItemType.Image = Global.ResourceManagement.My.Resources.Resources._24updatecommit
        Me.btnUpdateListItemType.Location = New System.Drawing.Point(252, 113)
        Me.btnUpdateListItemType.Margin = New System.Windows.Forms.Padding(5)
        Me.btnUpdateListItemType.Name = "btnUpdateListItemType"
        Me.btnUpdateListItemType.Size = New System.Drawing.Size(32, 32)
        Me.btnUpdateListItemType.TabIndex = 10
        Me.btnUpdateListItemType.UseVisualStyleBackColor = True
        '
        'btnAddNewListItemType
        '
        Me.btnAddNewListItemType.Image = Global.ResourceManagement.My.Resources.Resources._24new
        Me.btnAddNewListItemType.Location = New System.Drawing.Point(252, 71)
        Me.btnAddNewListItemType.Margin = New System.Windows.Forms.Padding(5)
        Me.btnAddNewListItemType.Name = "btnAddNewListItemType"
        Me.btnAddNewListItemType.Size = New System.Drawing.Size(32, 32)
        Me.btnAddNewListItemType.TabIndex = 9
        Me.btnAddNewListItemType.UseVisualStyleBackColor = True
        '
        'btnDeleteListItemType
        '
        Me.btnDeleteListItemType.Image = Global.ResourceManagement.My.Resources.Resources._24delete
        Me.btnDeleteListItemType.Location = New System.Drawing.Point(252, 155)
        Me.btnDeleteListItemType.Margin = New System.Windows.Forms.Padding(5)
        Me.btnDeleteListItemType.Name = "btnDeleteListItemType"
        Me.btnDeleteListItemType.Size = New System.Drawing.Size(32, 32)
        Me.btnDeleteListItemType.TabIndex = 11
        Me.btnDeleteListItemType.UseVisualStyleBackColor = True
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(73, 17)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(87, 17)
        Me.Label1.TabIndex = 12
        Me.Label1.Text = "List Item Type"
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(375, 17)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(56, 17)
        Me.Label2.TabIndex = 13
        Me.Label2.Text = "List Item"
        '
        'ListItemTypeAndItem
        '
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.None
        Me.ClientSize = New System.Drawing.Size(571, 243)
        Me.Controls.Add(Me.txtListItemTypeName)
        Me.Controls.Add(Me.txtListItemName)
        Me.Controls.Add(Me.btnUpdateListItemType)
        Me.Controls.Add(Me.btnAddNewListItemType)
        Me.Controls.Add(Me.btnDeleteListItemType)
        Me.Controls.Add(Me.lstListItemTypes)
        Me.Controls.Add(Me.btnClose)
        Me.Controls.Add(Me.btnUpdateListItem)
        Me.Controls.Add(Me.btnAddNewListItem)
        Me.Controls.Add(Me.btnDeleteListItem)
        Me.Controls.Add(Me.lstListItems)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.Label2)
        Me.Font = New System.Drawing.Font("Segoe UI", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.Margin = New System.Windows.Forms.Padding(4)
        Me.Name = "ListItemTypeAndItem"
        Me.SizeGripStyle = System.Windows.Forms.SizeGripStyle.Hide
        Me.Text = "List Type and Item"
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

    Friend WithEvents btnClose As Windows.Forms.Button
  Friend WithEvents btnUpdateListItem As Windows.Forms.Button
  Friend WithEvents btnAddNewListItem As Windows.Forms.Button
  Friend WithEvents btnDeleteListItem As Windows.Forms.Button
  Friend WithEvents lstListItems As Windows.Forms.ListBox
  Friend WithEvents txtListItemName As Windows.Forms.TextBox
  Friend WithEvents lstListItemTypes As Windows.Forms.ListBox
  Friend WithEvents txtListItemTypeName As Windows.Forms.TextBox
  Friend WithEvents btnUpdateListItemType As Windows.Forms.Button
  Friend WithEvents btnAddNewListItemType As Windows.Forms.Button
  Friend WithEvents btnDeleteListItemType As Windows.Forms.Button
  Friend WithEvents Label1 As Windows.Forms.Label
  Friend WithEvents Label2 As Windows.Forms.Label
End Class
