<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()>
Partial Class Closure
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
    Me.components = New System.ComponentModel.Container()
    Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(Closure))
    Me.btnClose = New System.Windows.Forms.Button()
    Me.btnUpdate = New System.Windows.Forms.Button()
    Me.btnAddNew = New System.Windows.Forms.Button()
    Me.btnDelete = New System.Windows.Forms.Button()
    Me.Label1 = New System.Windows.Forms.Label()
    Me.Label2 = New System.Windows.Forms.Label()
    Me.Label3 = New System.Windows.Forms.Label()
    Me.txtClosureName = New System.Windows.Forms.TextBox()
    Me.dtpStartDate = New System.Windows.Forms.DateTimePicker()
    Me.dtpEndDate = New System.Windows.Forms.DateTimePicker()
    Me.dgvClosures = New System.Windows.Forms.DataGridView()
    Me.ToolTip1 = New System.Windows.Forms.ToolTip(Me.components)
        CType(Me.dgvClosures, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'btnClose
        '
        Me.btnClose.Image = Global.ResourceManagement.My.Resources.Resources._24close
        Me.btnClose.Location = New System.Drawing.Point(451, 207)
        Me.btnClose.Margin = New System.Windows.Forms.Padding(4)
        Me.btnClose.Name = "btnClose"
        Me.btnClose.Size = New System.Drawing.Size(32, 32)
        Me.btnClose.TabIndex = 10
        Me.btnClose.UseVisualStyleBackColor = True
        '
        'btnUpdate
        '
        Me.btnUpdate.Image = Global.ResourceManagement.My.Resources.Resources._24updatecommit
        Me.btnUpdate.Location = New System.Drawing.Point(451, 100)
        Me.btnUpdate.Margin = New System.Windows.Forms.Padding(4)
        Me.btnUpdate.Name = "btnUpdate"
        Me.btnUpdate.Size = New System.Drawing.Size(32, 32)
        Me.btnUpdate.TabIndex = 8
        Me.btnUpdate.UseVisualStyleBackColor = True
        '
        'btnAddNew
        '
        Me.btnAddNew.Image = Global.ResourceManagement.My.Resources.Resources._24new
        Me.btnAddNew.Location = New System.Drawing.Point(451, 60)
        Me.btnAddNew.Margin = New System.Windows.Forms.Padding(4)
        Me.btnAddNew.Name = "btnAddNew"
        Me.btnAddNew.Size = New System.Drawing.Size(32, 32)
        Me.btnAddNew.TabIndex = 7
        Me.btnAddNew.UseVisualStyleBackColor = True
        '
        'btnDelete
        '
        Me.btnDelete.Image = Global.ResourceManagement.My.Resources.Resources._24delete
        Me.btnDelete.Location = New System.Drawing.Point(451, 140)
        Me.btnDelete.Margin = New System.Windows.Forms.Padding(4)
        Me.btnDelete.Name = "btnDelete"
        Me.btnDelete.Size = New System.Drawing.Size(32, 32)
        Me.btnDelete.TabIndex = 9
        Me.btnDelete.UseVisualStyleBackColor = True
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(87, 9)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(43, 17)
        Me.Label1.TabIndex = 0
        Me.Label1.Text = "Name"
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(235, 9)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(66, 17)
        Me.Label2.TabIndex = 2
        Me.Label2.Text = "Start Date"
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Location = New System.Drawing.Point(356, 9)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(61, 17)
        Me.Label3.TabIndex = 4
        Me.Label3.Text = "End Date"
        '
        'txtClosureName
        '
        Me.txtClosureName.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.txtClosureName.Location = New System.Drawing.Point(12, 29)
        Me.txtClosureName.Name = "txtClosureName"
        Me.txtClosureName.Size = New System.Drawing.Size(191, 25)
        Me.txtClosureName.TabIndex = 1
        '
        'dtpStartDate
        '
        Me.dtpStartDate.Format = System.Windows.Forms.DateTimePickerFormat.[Short]
        Me.dtpStartDate.Location = New System.Drawing.Point(210, 29)
        Me.dtpStartDate.Name = "dtpStartDate"
        Me.dtpStartDate.ShowCheckBox = True
        Me.dtpStartDate.Size = New System.Drawing.Size(114, 25)
        Me.dtpStartDate.TabIndex = 3
        '
        'dtpEndDate
        '
        Me.dtpEndDate.Format = System.Windows.Forms.DateTimePickerFormat.[Short]
        Me.dtpEndDate.Location = New System.Drawing.Point(330, 29)
        Me.dtpEndDate.Name = "dtpEndDate"
        Me.dtpEndDate.ShowCheckBox = True
        Me.dtpEndDate.Size = New System.Drawing.Size(114, 25)
        Me.dtpEndDate.TabIndex = 5
        '
        'dgvClosures
        '
        Me.dgvClosures.AllowUserToAddRows = False
        Me.dgvClosures.AllowUserToDeleteRows = False
        Me.dgvClosures.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.dgvClosures.ColumnHeadersVisible = False
        Me.dgvClosures.Location = New System.Drawing.Point(12, 60)
        Me.dgvClosures.Name = "dgvClosures"
        Me.dgvClosures.ReadOnly = True
        Me.dgvClosures.RowHeadersVisible = False
        Me.dgvClosures.SelectionMode = System.Windows.Forms.DataGridViewSelectionMode.FullRowSelect
        Me.dgvClosures.Size = New System.Drawing.Size(432, 142)
        Me.dgvClosures.TabIndex = 6
        '
        'Closure
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(7.0!, 17.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(495, 251)
        Me.Controls.Add(Me.dgvClosures)
        Me.Controls.Add(Me.dtpEndDate)
        Me.Controls.Add(Me.dtpStartDate)
        Me.Controls.Add(Me.txtClosureName)
        Me.Controls.Add(Me.Label3)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.btnClose)
        Me.Controls.Add(Me.btnUpdate)
        Me.Controls.Add(Me.btnAddNew)
        Me.Controls.Add(Me.btnDelete)
        Me.Font = New System.Drawing.Font("Segoe UI", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.Margin = New System.Windows.Forms.Padding(4)
        Me.Name = "Closure"
        Me.SizeGripStyle = System.Windows.Forms.SizeGripStyle.Hide
        Me.Text = "Closure"
        CType(Me.dgvClosures, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents btnClose As Windows.Forms.Button
    Friend WithEvents btnUpdate As Windows.Forms.Button
    Friend WithEvents btnAddNew As Windows.Forms.Button
    Friend WithEvents btnDelete As Windows.Forms.Button
    Friend WithEvents Label1 As Windows.Forms.Label
    Friend WithEvents Label2 As Windows.Forms.Label
    Friend WithEvents Label3 As Windows.Forms.Label
    Friend WithEvents txtClosureName As Windows.Forms.TextBox
    Friend WithEvents dtpStartDate As Windows.Forms.DateTimePicker
    Friend WithEvents dtpEndDate As Windows.Forms.DateTimePicker
  Friend WithEvents dgvClosures As Windows.Forms.DataGridView
  Friend WithEvents ToolTip1 As Windows.Forms.ToolTip
End Class
