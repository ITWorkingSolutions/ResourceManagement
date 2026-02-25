<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()>
Partial Class DatabaseManager
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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(DatabaseManager))
        Me.lstDatabasePaths = New System.Windows.Forms.ListBox()
        Me.btnDelete = New System.Windows.Forms.Button()
        Me.btnAddNew = New System.Windows.Forms.Button()
        Me.btnActivate = New System.Windows.Forms.Button()
        Me.btnUpdate = New System.Windows.Forms.Button()
        Me.btnClose = New System.Windows.Forms.Button()
        Me.ToolTip1 = New System.Windows.Forms.ToolTip(Me.components)
        Me.btnAddExisting = New System.Windows.Forms.Button()
        Me.SuspendLayout()
        '
        'lstDatabasePaths
        '
        Me.lstDatabasePaths.FormattingEnabled = True
        Me.lstDatabasePaths.ItemHeight = 17
        Me.lstDatabasePaths.Location = New System.Drawing.Point(14, 16)
        Me.lstDatabasePaths.Margin = New System.Windows.Forms.Padding(4)
        Me.lstDatabasePaths.Name = "lstDatabasePaths"
        Me.lstDatabasePaths.Size = New System.Drawing.Size(386, 242)
        Me.lstDatabasePaths.TabIndex = 0
        '
        'btnDelete
        '
        Me.btnDelete.Image = Global.ResourceManagement.My.Resources.Resources._24delete
        Me.btnDelete.Location = New System.Drawing.Point(408, 165)
        Me.btnDelete.Margin = New System.Windows.Forms.Padding(4)
        Me.btnDelete.Name = "btnDelete"
        Me.btnDelete.Size = New System.Drawing.Size(37, 42)
        Me.btnDelete.TabIndex = 1
        Me.btnDelete.UseVisualStyleBackColor = True
        '
        'btnAddNew
        '
        Me.btnAddNew.Image = Global.ResourceManagement.My.Resources.Resources._24new
        Me.btnAddNew.Location = New System.Drawing.Point(408, 15)
        Me.btnAddNew.Margin = New System.Windows.Forms.Padding(4)
        Me.btnAddNew.Name = "btnAddNew"
        Me.btnAddNew.Size = New System.Drawing.Size(37, 42)
        Me.btnAddNew.TabIndex = 2
        Me.btnAddNew.UseVisualStyleBackColor = True
        '
        'btnActivate
        '
        Me.btnActivate.Image = Global.ResourceManagement.My.Resources.Resources._24set
        Me.btnActivate.Location = New System.Drawing.Point(408, 215)
        Me.btnActivate.Margin = New System.Windows.Forms.Padding(4)
        Me.btnActivate.Name = "btnActivate"
        Me.btnActivate.Size = New System.Drawing.Size(37, 42)
        Me.btnActivate.TabIndex = 3
        Me.btnActivate.UseVisualStyleBackColor = True
        '
        'btnUpdate
        '
        Me.btnUpdate.Image = Global.ResourceManagement.My.Resources.Resources._24updatecommit
        Me.btnUpdate.Location = New System.Drawing.Point(408, 115)
        Me.btnUpdate.Margin = New System.Windows.Forms.Padding(4)
        Me.btnUpdate.Name = "btnUpdate"
        Me.btnUpdate.Size = New System.Drawing.Size(37, 42)
        Me.btnUpdate.TabIndex = 4
        Me.btnUpdate.UseVisualStyleBackColor = True
        '
        'btnClose
        '
        Me.btnClose.Image = Global.ResourceManagement.My.Resources.Resources._24close
        Me.btnClose.Location = New System.Drawing.Point(408, 265)
        Me.btnClose.Margin = New System.Windows.Forms.Padding(4)
        Me.btnClose.Name = "btnClose"
        Me.btnClose.Size = New System.Drawing.Size(37, 42)
        Me.btnClose.TabIndex = 5
        Me.btnClose.UseVisualStyleBackColor = True
        '
        'btnAddExisting
        '
        Me.btnAddExisting.Image = Global.ResourceManagement.My.Resources.Resources._24openfolder
        Me.btnAddExisting.Location = New System.Drawing.Point(408, 65)
        Me.btnAddExisting.Margin = New System.Windows.Forms.Padding(4)
        Me.btnAddExisting.Name = "btnAddExisting"
        Me.btnAddExisting.Size = New System.Drawing.Size(37, 42)
        Me.btnAddExisting.TabIndex = 6
        Me.btnAddExisting.UseVisualStyleBackColor = True
        '
        'DatabaseManager
        '
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.None
        Me.ClientSize = New System.Drawing.Size(460, 322)
        Me.Controls.Add(Me.btnAddExisting)
        Me.Controls.Add(Me.btnClose)
        Me.Controls.Add(Me.btnUpdate)
        Me.Controls.Add(Me.btnActivate)
        Me.Controls.Add(Me.btnAddNew)
        Me.Controls.Add(Me.btnDelete)
        Me.Controls.Add(Me.lstDatabasePaths)
        Me.Font = New System.Drawing.Font("Segoe UI", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.Margin = New System.Windows.Forms.Padding(4)
        Me.Name = "DatabaseManager"
        Me.SizeGripStyle = System.Windows.Forms.SizeGripStyle.Hide
        Me.Text = "Database Manager"
        Me.ResumeLayout(False)

    End Sub

    Friend WithEvents lstDatabasePaths As Windows.Forms.ListBox
    Friend WithEvents btnDelete As Windows.Forms.Button
    Friend WithEvents btnAddNew As Windows.Forms.Button
    Friend WithEvents btnActivate As Windows.Forms.Button
    Friend WithEvents btnUpdate As Windows.Forms.Button
    Friend WithEvents btnClose As Windows.Forms.Button
    Friend WithEvents ToolTip1 As Windows.Forms.ToolTip
    Friend WithEvents btnAddExisting As Windows.Forms.Button
End Class
