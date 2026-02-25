<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class SelectDefaultMode
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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(SelectDefaultMode))
        Me.lblMessage = New System.Windows.Forms.Label()
        Me.btnAvailable = New System.Windows.Forms.Button()
        Me.btnUnavailable = New System.Windows.Forms.Button()
        Me.btnCancel = New System.Windows.Forms.Button()
        Me.SuspendLayout()
        '
        'lblMessage
        '
        Me.lblMessage.Location = New System.Drawing.Point(13, 9)
        Me.lblMessage.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.lblMessage.Name = "lblMessage"
        Me.lblMessage.Size = New System.Drawing.Size(435, 55)
        Me.lblMessage.TabIndex = 0
        Me.lblMessage.Text = "Please select the default mode for resources. If no record is created for the ind" &
    "ividual resource explicitly stating 'Availability' or 'Unavailability' this will" &
    " be the default for the resource."
        '
        'btnAvailable
        '
        Me.btnAvailable.DialogResult = System.Windows.Forms.DialogResult.OK
        Me.btnAvailable.Location = New System.Drawing.Point(181, 65)
        Me.btnAvailable.Name = "btnAvailable"
        Me.btnAvailable.Size = New System.Drawing.Size(84, 27)
        Me.btnAvailable.TabIndex = 1
        Me.btnAvailable.Text = "Available"
        Me.btnAvailable.UseVisualStyleBackColor = True
        '
        'btnUnavailable
        '
        Me.btnUnavailable.DialogResult = System.Windows.Forms.DialogResult.OK
        Me.btnUnavailable.Location = New System.Drawing.Point(271, 65)
        Me.btnUnavailable.Name = "btnUnavailable"
        Me.btnUnavailable.Size = New System.Drawing.Size(84, 27)
        Me.btnUnavailable.TabIndex = 2
        Me.btnUnavailable.Text = "Unavailable"
        Me.btnUnavailable.UseVisualStyleBackColor = True
        '
        'btnCancel
        '
        Me.btnCancel.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.btnCancel.Location = New System.Drawing.Point(361, 65)
        Me.btnCancel.Name = "btnCancel"
        Me.btnCancel.Size = New System.Drawing.Size(84, 27)
        Me.btnCancel.TabIndex = 3
        Me.btnCancel.Text = "Cancel"
        Me.btnCancel.UseVisualStyleBackColor = True
        '
        'SelectDefaultMode
        '
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.None
        Me.ClientSize = New System.Drawing.Size(457, 104)
        Me.Controls.Add(Me.btnCancel)
        Me.Controls.Add(Me.btnUnavailable)
        Me.Controls.Add(Me.btnAvailable)
        Me.Controls.Add(Me.lblMessage)
        Me.Font = New System.Drawing.Font("Segoe UI", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.Margin = New System.Windows.Forms.Padding(4)
        Me.Name = "SelectDefaultMode"
        Me.SizeGripStyle = System.Windows.Forms.SizeGripStyle.Hide
        Me.Text = "Select database default mode"
        Me.ResumeLayout(False)

    End Sub

    Friend WithEvents lblMessage As Windows.Forms.Label
    Friend WithEvents btnAvailable As Windows.Forms.Button
    Friend WithEvents btnUnavailable As Windows.Forms.Button
    Friend WithEvents btnCancel As Windows.Forms.Button
End Class
