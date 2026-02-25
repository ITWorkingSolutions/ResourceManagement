<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()>
Partial Class Resource
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
    Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(Resource))
    Me.txtPreferredName = New System.Windows.Forms.TextBox()
    Me.Label1 = New System.Windows.Forms.Label()
    Me.Label3 = New System.Windows.Forms.Label()
    Me.Label2 = New System.Windows.Forms.Label()
    Me.dtpStartDate = New System.Windows.Forms.DateTimePicker()
    Me.dtpEndDate = New System.Windows.Forms.DateTimePicker()
    Me.btnClose = New System.Windows.Forms.Button()
    Me.btnSave = New System.Windows.Forms.Button()
    Me.btnDelete = New System.Windows.Forms.Button()
    Me.cmbSalutation = New System.Windows.Forms.ComboBox()
    Me.lblSalutation = New System.Windows.Forms.Label()
    Me.txtFirstName = New System.Windows.Forms.TextBox()
    Me.txtLastName = New System.Windows.Forms.TextBox()
    Me.Label7 = New System.Windows.Forms.Label()
    Me.Label8 = New System.Windows.Forms.Label()
    Me.txtEmail = New System.Windows.Forms.TextBox()
    Me.Label9 = New System.Windows.Forms.Label()
    Me.Label10 = New System.Windows.Forms.Label()
    Me.PictureBox1 = New System.Windows.Forms.PictureBox()
    Me.PictureBox2 = New System.Windows.Forms.PictureBox()
    Me.GroupBox1 = New System.Windows.Forms.GroupBox()
    Me.txtPhone = New System.Windows.Forms.TextBox()
    Me.GroupBox2 = New System.Windows.Forms.GroupBox()
    Me.cmbGender = New System.Windows.Forms.ComboBox()
    Me.lblGender = New System.Windows.Forms.Label()
    Me.GroupBox3 = New System.Windows.Forms.GroupBox()
    Me.GroupBox6 = New System.Windows.Forms.GroupBox()
    Me.txtNotes = New System.Windows.Forms.TextBox()
    Me.Label12 = New System.Windows.Forms.Label()
    Me.GroupBox4 = New System.Windows.Forms.GroupBox()
        Me.Label4 = New System.Windows.Forms.Label()
        Me.dgvResourceNameValue = New System.Windows.Forms.DataGridView()
        Me.GroupBox7 = New System.Windows.Forms.GroupBox()
        Me.btnCancelListItems = New System.Windows.Forms.Button()
        Me.btnSaveListItems = New System.Windows.Forms.Button()
        Me.Label5 = New System.Windows.Forms.Label()
        Me.lstListItems = New System.Windows.Forms.ListView()
        Me.cmbResourceListItemNames = New System.Windows.Forms.ComboBox()
        CType(Me.PictureBox1, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.PictureBox2, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.GroupBox1.SuspendLayout()
        Me.GroupBox2.SuspendLayout()
        Me.GroupBox3.SuspendLayout()
        Me.GroupBox6.SuspendLayout()
        Me.GroupBox4.SuspendLayout()
        CType(Me.dgvResourceNameValue, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.GroupBox7.SuspendLayout()
        Me.SuspendLayout()
        '
        'txtPreferredName
        '
        Me.txtPreferredName.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.txtPreferredName.Location = New System.Drawing.Point(446, 33)
        Me.txtPreferredName.Margin = New System.Windows.Forms.Padding(4)
        Me.txtPreferredName.Name = "txtPreferredName"
        Me.txtPreferredName.Size = New System.Drawing.Size(198, 25)
        Me.txtPreferredName.TabIndex = 3
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.ImageAlign = System.Drawing.ContentAlignment.MiddleRight
        Me.Label1.Location = New System.Drawing.Point(334, 36)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(105, 17)
        Me.Label1.TabIndex = 2
        Me.Label1.Text = "Preferred Name:"
        Me.Label1.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.ImageAlign = System.Drawing.ContentAlignment.MiddleRight
        Me.Label3.Location = New System.Drawing.Point(42, 66)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(64, 17)
        Me.Label3.TabIndex = 2
        Me.Label3.Text = "End Date:"
        Me.Label3.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.ImageAlign = System.Drawing.ContentAlignment.MiddleRight
        Me.Label2.Location = New System.Drawing.Point(37, 35)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(69, 17)
        Me.Label2.TabIndex = 0
        Me.Label2.Text = "Start Date:"
        Me.Label2.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'dtpStartDate
        '
        Me.dtpStartDate.Checked = False
        Me.dtpStartDate.Format = System.Windows.Forms.DateTimePickerFormat.[Short]
        Me.dtpStartDate.Location = New System.Drawing.Point(112, 29)
        Me.dtpStartDate.Name = "dtpStartDate"
        Me.dtpStartDate.ShowCheckBox = True
        Me.dtpStartDate.Size = New System.Drawing.Size(108, 25)
        Me.dtpStartDate.TabIndex = 1
        '
        'dtpEndDate
        '
        Me.dtpEndDate.Checked = False
        Me.dtpEndDate.Format = System.Windows.Forms.DateTimePickerFormat.[Short]
        Me.dtpEndDate.Location = New System.Drawing.Point(112, 60)
        Me.dtpEndDate.Name = "dtpEndDate"
        Me.dtpEndDate.ShowCheckBox = True
        Me.dtpEndDate.Size = New System.Drawing.Size(108, 25)
        Me.dtpEndDate.TabIndex = 3
        '
        'btnClose
        '
        Me.btnClose.Image = Global.ResourceManagement.My.Resources.Resources._24close
        Me.btnClose.Location = New System.Drawing.Point(643, 688)
        Me.btnClose.Margin = New System.Windows.Forms.Padding(4)
        Me.btnClose.Name = "btnClose"
        Me.btnClose.Size = New System.Drawing.Size(32, 32)
        Me.btnClose.TabIndex = 8
        Me.btnClose.UseVisualStyleBackColor = True
        '
        'btnSave
        '
        Me.btnSave.Image = Global.ResourceManagement.My.Resources.Resources._24save
        Me.btnSave.Location = New System.Drawing.Point(563, 688)
        Me.btnSave.Margin = New System.Windows.Forms.Padding(4)
        Me.btnSave.Name = "btnSave"
        Me.btnSave.Size = New System.Drawing.Size(32, 32)
        Me.btnSave.TabIndex = 6
        Me.btnSave.UseVisualStyleBackColor = True
        '
        'btnDelete
        '
        Me.btnDelete.Image = Global.ResourceManagement.My.Resources.Resources._24delete
        Me.btnDelete.Location = New System.Drawing.Point(603, 688)
        Me.btnDelete.Margin = New System.Windows.Forms.Padding(4)
        Me.btnDelete.Name = "btnDelete"
        Me.btnDelete.Size = New System.Drawing.Size(32, 32)
        Me.btnDelete.TabIndex = 7
        Me.btnDelete.UseVisualStyleBackColor = True
        '
        'cmbSalutation
        '
        Me.cmbSalutation.FormattingEnabled = True
        Me.cmbSalutation.Location = New System.Drawing.Point(118, 32)
        Me.cmbSalutation.Name = "cmbSalutation"
        Me.cmbSalutation.Size = New System.Drawing.Size(102, 25)
        Me.cmbSalutation.TabIndex = 1
        '
        'lblSalutation
        '
        Me.lblSalutation.Location = New System.Drawing.Point(8, 35)
        Me.lblSalutation.Name = "lblSalutation"
        Me.lblSalutation.Size = New System.Drawing.Size(104, 17)
        Me.lblSalutation.TabIndex = 0
        Me.lblSalutation.Text = "Salutation:"
        Me.lblSalutation.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'txtFirstName
        '
        Me.txtFirstName.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.txtFirstName.Location = New System.Drawing.Point(118, 67)
        Me.txtFirstName.Margin = New System.Windows.Forms.Padding(4)
        Me.txtFirstName.Name = "txtFirstName"
        Me.txtFirstName.Size = New System.Drawing.Size(198, 25)
        Me.txtFirstName.TabIndex = 5
        '
        'txtLastName
        '
        Me.txtLastName.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.txtLastName.Location = New System.Drawing.Point(446, 67)
        Me.txtLastName.Margin = New System.Windows.Forms.Padding(4)
        Me.txtLastName.Name = "txtLastName"
        Me.txtLastName.Size = New System.Drawing.Size(198, 25)
        Me.txtLastName.TabIndex = 7
        '
        'Label7
        '
        Me.Label7.AutoSize = True
        Me.Label7.Location = New System.Drawing.Point(37, 69)
        Me.Label7.Name = "Label7"
        Me.Label7.Size = New System.Drawing.Size(74, 17)
        Me.Label7.TabIndex = 4
        Me.Label7.Text = "First Name:"
        Me.Label7.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'Label8
        '
        Me.Label8.AutoSize = True
        Me.Label8.Location = New System.Drawing.Point(366, 67)
        Me.Label8.Name = "Label8"
        Me.Label8.Size = New System.Drawing.Size(73, 17)
        Me.Label8.TabIndex = 6
        Me.Label8.Text = "Last Name:"
        Me.Label8.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'txtEmail
        '
        Me.txtEmail.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.txtEmail.Location = New System.Drawing.Point(112, 30)
        Me.txtEmail.Margin = New System.Windows.Forms.Padding(4)
        Me.txtEmail.Name = "txtEmail"
        Me.txtEmail.Size = New System.Drawing.Size(198, 25)
        Me.txtEmail.TabIndex = 1
        '
        'Label9
        '
        Me.Label9.Image = Global.ResourceManagement.My.Resources.Resources._16email
        Me.Label9.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.Label9.Location = New System.Drawing.Point(44, 31)
        Me.Label9.Name = "Label9"
        Me.Label9.Size = New System.Drawing.Size(61, 19)
        Me.Label9.TabIndex = 0
        Me.Label9.Text = "Email:"
        Me.Label9.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'Label10
        '
        Me.Label10.Image = Global.ResourceManagement.My.Resources.Resources._16telephone
        Me.Label10.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.Label10.Location = New System.Drawing.Point(39, 64)
        Me.Label10.Name = "Label10"
        Me.Label10.Size = New System.Drawing.Size(66, 19)
        Me.Label10.TabIndex = 2
        Me.Label10.Text = "Phone:"
        Me.Label10.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'PictureBox1
        '
        Me.PictureBox1.Image = CType(resources.GetObject("PictureBox1.Image"), System.Drawing.Image)
        Me.PictureBox1.Location = New System.Drawing.Point(6, 14)
        Me.PictureBox1.Name = "PictureBox1"
        Me.PictureBox1.Size = New System.Drawing.Size(24, 20)
        Me.PictureBox1.SizeMode = System.Windows.Forms.PictureBoxSizeMode.CenterImage
        Me.PictureBox1.TabIndex = 39
        Me.PictureBox1.TabStop = False
        '
        'PictureBox2
        '
        Me.PictureBox2.Image = Global.ResourceManagement.My.Resources.Resources._16lifecycle
        Me.PictureBox2.Location = New System.Drawing.Point(6, 15)
        Me.PictureBox2.Name = "PictureBox2"
        Me.PictureBox2.Size = New System.Drawing.Size(24, 20)
        Me.PictureBox2.SizeMode = System.Windows.Forms.PictureBoxSizeMode.CenterImage
        Me.PictureBox2.TabIndex = 40
        Me.PictureBox2.TabStop = False
        '
        'GroupBox1
        '
        Me.GroupBox1.Controls.Add(Me.txtPhone)
        Me.GroupBox1.Controls.Add(Me.txtEmail)
        Me.GroupBox1.Controls.Add(Me.Label9)
        Me.GroupBox1.Controls.Add(Me.Label10)
        Me.GroupBox1.Location = New System.Drawing.Point(12, 243)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Size = New System.Drawing.Size(325, 139)
        Me.GroupBox1.TabIndex = 1
        Me.GroupBox1.TabStop = False
        '
        'txtPhone
        '
        Me.txtPhone.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.txtPhone.Location = New System.Drawing.Point(112, 63)
        Me.txtPhone.Margin = New System.Windows.Forms.Padding(4)
        Me.txtPhone.Name = "txtPhone"
        Me.txtPhone.Size = New System.Drawing.Size(198, 25)
        Me.txtPhone.TabIndex = 3
        '
        'GroupBox2
        '
        Me.GroupBox2.Controls.Add(Me.cmbGender)
        Me.GroupBox2.Controls.Add(Me.lblGender)
        Me.GroupBox2.Controls.Add(Me.PictureBox1)
        Me.GroupBox2.Controls.Add(Me.cmbSalutation)
        Me.GroupBox2.Controls.Add(Me.lblSalutation)
        Me.GroupBox2.Controls.Add(Me.txtPreferredName)
        Me.GroupBox2.Controls.Add(Me.Label1)
        Me.GroupBox2.Controls.Add(Me.Label7)
        Me.GroupBox2.Controls.Add(Me.Label8)
        Me.GroupBox2.Controls.Add(Me.txtFirstName)
        Me.GroupBox2.Controls.Add(Me.txtLastName)
        Me.GroupBox2.Location = New System.Drawing.Point(12, 6)
        Me.GroupBox2.Name = "GroupBox2"
        Me.GroupBox2.Size = New System.Drawing.Size(662, 135)
        Me.GroupBox2.TabIndex = 0
        Me.GroupBox2.TabStop = False
        '
        'cmbGender
        '
        Me.cmbGender.FormattingEnabled = True
        Me.cmbGender.Location = New System.Drawing.Point(118, 99)
        Me.cmbGender.Name = "cmbGender"
        Me.cmbGender.Size = New System.Drawing.Size(102, 25)
        Me.cmbGender.TabIndex = 9
        '
        'lblGender
        '
        Me.lblGender.Location = New System.Drawing.Point(8, 102)
        Me.lblGender.Name = "lblGender"
        Me.lblGender.Size = New System.Drawing.Size(104, 17)
        Me.lblGender.TabIndex = 8
        Me.lblGender.Text = "Gender:"
        Me.lblGender.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'GroupBox3
        '
        Me.GroupBox3.Controls.Add(Me.PictureBox2)
        Me.GroupBox3.Controls.Add(Me.dtpStartDate)
        Me.GroupBox3.Controls.Add(Me.Label2)
        Me.GroupBox3.Controls.Add(Me.Label3)
        Me.GroupBox3.Controls.Add(Me.dtpEndDate)
        Me.GroupBox3.Location = New System.Drawing.Point(12, 141)
        Me.GroupBox3.Name = "GroupBox3"
        Me.GroupBox3.Size = New System.Drawing.Size(325, 101)
        Me.GroupBox3.TabIndex = 2
        Me.GroupBox3.TabStop = False
        '
        'GroupBox6
        '
        Me.GroupBox6.Controls.Add(Me.txtNotes)
        Me.GroupBox6.Controls.Add(Me.Label12)
        Me.GroupBox6.Location = New System.Drawing.Point(13, 538)
        Me.GroupBox6.Name = "GroupBox6"
        Me.GroupBox6.Size = New System.Drawing.Size(659, 143)
        Me.GroupBox6.TabIndex = 5
        Me.GroupBox6.TabStop = False
        '
        'txtNotes
        '
        Me.txtNotes.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.txtNotes.Location = New System.Drawing.Point(110, 24)
        Me.txtNotes.Multiline = True
        Me.txtNotes.Name = "txtNotes"
        Me.txtNotes.Size = New System.Drawing.Size(532, 103)
        Me.txtNotes.TabIndex = 1
        '
        'Label12
        '
        Me.Label12.Image = Global.ResourceManagement.My.Resources.Resources._16pen
        Me.Label12.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.Label12.Location = New System.Drawing.Point(36, 25)
        Me.Label12.Name = "Label12"
        Me.Label12.Size = New System.Drawing.Size(68, 19)
        Me.Label12.TabIndex = 0
        Me.Label12.Text = "Notes:"
        Me.Label12.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'GroupBox4
        '
        Me.GroupBox4.Controls.Add(Me.Label4)
        Me.GroupBox4.Controls.Add(Me.dgvResourceNameValue)
        Me.GroupBox4.Location = New System.Drawing.Point(12, 382)
        Me.GroupBox4.Name = "GroupBox4"
        Me.GroupBox4.Size = New System.Drawing.Size(658, 156)
        Me.GroupBox4.TabIndex = 9
        Me.GroupBox4.TabStop = False
        '
        'Label4
        '
        Me.Label4.Image = Global.ResourceManagement.My.Resources.Resources._16tag
        Me.Label4.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.Label4.Location = New System.Drawing.Point(6, 23)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(98, 19)
        Me.Label4.TabIndex = 45
        Me.Label4.Text = "Single Value:"
        Me.Label4.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'dgvResourceNameValue
        '
        Me.dgvResourceNameValue.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.dgvResourceNameValue.Location = New System.Drawing.Point(110, 23)
        Me.dgvResourceNameValue.Name = "dgvResourceNameValue"
        Me.dgvResourceNameValue.Size = New System.Drawing.Size(531, 119)
        Me.dgvResourceNameValue.TabIndex = 44
        '
        'GroupBox7
        '
        Me.GroupBox7.Controls.Add(Me.btnCancelListItems)
        Me.GroupBox7.Controls.Add(Me.btnSaveListItems)
        Me.GroupBox7.Controls.Add(Me.Label5)
        Me.GroupBox7.Controls.Add(Me.lstListItems)
        Me.GroupBox7.Controls.Add(Me.cmbResourceListItemNames)
        Me.GroupBox7.Location = New System.Drawing.Point(349, 142)
        Me.GroupBox7.Name = "GroupBox7"
        Me.GroupBox7.Size = New System.Drawing.Size(325, 240)
        Me.GroupBox7.TabIndex = 10
        Me.GroupBox7.TabStop = False
        '
        'btnCancelListItems
        '
        Me.btnCancelListItems.Image = Global.ResourceManagement.My.Resources.Resources._24cancel
        Me.btnCancelListItems.Location = New System.Drawing.Point(275, 196)
        Me.btnCancelListItems.Margin = New System.Windows.Forms.Padding(4)
        Me.btnCancelListItems.Name = "btnCancelListItems"
        Me.btnCancelListItems.Size = New System.Drawing.Size(32, 32)
        Me.btnCancelListItems.TabIndex = 48
        Me.btnCancelListItems.UseVisualStyleBackColor = True
        '
        'btnSaveListItems
        '
        Me.btnSaveListItems.Image = Global.ResourceManagement.My.Resources.Resources._24update
        Me.btnSaveListItems.Location = New System.Drawing.Point(235, 196)
        Me.btnSaveListItems.Margin = New System.Windows.Forms.Padding(4)
        Me.btnSaveListItems.Name = "btnSaveListItems"
        Me.btnSaveListItems.Size = New System.Drawing.Size(32, 32)
        Me.btnSaveListItems.TabIndex = 47
        Me.btnSaveListItems.UseVisualStyleBackColor = True
        '
        'Label5
        '
        Me.Label5.Image = Global.ResourceManagement.My.Resources.Resources._16tag
        Me.Label5.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.Label5.Location = New System.Drawing.Point(6, 26)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(98, 19)
        Me.Label5.TabIndex = 46
        Me.Label5.Text = "Multi Value:"
        Me.Label5.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'lstListItems
        '
        Me.lstListItems.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.lstListItems.CheckBoxes = True
        Me.lstListItems.HideSelection = False
        Me.lstListItems.Location = New System.Drawing.Point(19, 55)
        Me.lstListItems.Name = "lstListItems"
        Me.lstListItems.Size = New System.Drawing.Size(288, 134)
        Me.lstListItems.TabIndex = 45
        Me.lstListItems.UseCompatibleStateImageBehavior = False
        '
        'cmbResourceListItemNames
        '
        Me.cmbResourceListItemNames.FormattingEnabled = True
        Me.cmbResourceListItemNames.Location = New System.Drawing.Point(108, 24)
        Me.cmbResourceListItemNames.Name = "cmbResourceListItemNames"
        Me.cmbResourceListItemNames.Size = New System.Drawing.Size(198, 25)
        Me.cmbResourceListItemNames.TabIndex = 43
        '
        'Resource
        '
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.None
        Me.ClientSize = New System.Drawing.Size(688, 730)
        Me.Controls.Add(Me.GroupBox7)
        Me.Controls.Add(Me.GroupBox4)
        Me.Controls.Add(Me.GroupBox6)
        Me.Controls.Add(Me.GroupBox3)
        Me.Controls.Add(Me.GroupBox2)
        Me.Controls.Add(Me.GroupBox1)
        Me.Controls.Add(Me.btnDelete)
        Me.Controls.Add(Me.btnSave)
        Me.Controls.Add(Me.btnClose)
        Me.Font = New System.Drawing.Font("Segoe UI", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.Margin = New System.Windows.Forms.Padding(4)
        Me.Name = "Resource"
        Me.SizeGripStyle = System.Windows.Forms.SizeGripStyle.Hide
        Me.Text = "Resource"
        CType(Me.PictureBox1, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.PictureBox2, System.ComponentModel.ISupportInitialize).EndInit()
        Me.GroupBox1.ResumeLayout(False)
        Me.GroupBox1.PerformLayout()
        Me.GroupBox2.ResumeLayout(False)
        Me.GroupBox2.PerformLayout()
        Me.GroupBox3.ResumeLayout(False)
        Me.GroupBox3.PerformLayout()
        Me.GroupBox6.ResumeLayout(False)
        Me.GroupBox6.PerformLayout()
        Me.GroupBox4.ResumeLayout(False)
        CType(Me.dgvResourceNameValue, System.ComponentModel.ISupportInitialize).EndInit()
        Me.GroupBox7.ResumeLayout(False)
        Me.ResumeLayout(False)

    End Sub

    Friend WithEvents txtPreferredName As Windows.Forms.TextBox
  Friend WithEvents Label1 As Windows.Forms.Label
  Friend WithEvents Label3 As Windows.Forms.Label
  Friend WithEvents Label2 As Windows.Forms.Label
  Friend WithEvents dtpStartDate As Windows.Forms.DateTimePicker
  Friend WithEvents dtpEndDate As Windows.Forms.DateTimePicker
  Friend WithEvents btnClose As Windows.Forms.Button
  Friend WithEvents btnSave As Windows.Forms.Button
  Friend WithEvents btnDelete As Windows.Forms.Button
  Friend WithEvents cmbSalutation As Windows.Forms.ComboBox
  Friend WithEvents lblSalutation As Windows.Forms.Label
  Friend WithEvents txtFirstName As Windows.Forms.TextBox
  Friend WithEvents txtLastName As Windows.Forms.TextBox
  Friend WithEvents Label7 As Windows.Forms.Label
  Friend WithEvents Label8 As Windows.Forms.Label
  Friend WithEvents txtEmail As Windows.Forms.TextBox
  Friend WithEvents Label9 As Windows.Forms.Label
  Friend WithEvents Label10 As Windows.Forms.Label
  Friend WithEvents PictureBox1 As Windows.Forms.PictureBox
  Friend WithEvents PictureBox2 As Windows.Forms.PictureBox
  Friend WithEvents GroupBox1 As Windows.Forms.GroupBox
  Friend WithEvents GroupBox2 As Windows.Forms.GroupBox
  Friend WithEvents txtPhone As Windows.Forms.TextBox
  Friend WithEvents GroupBox3 As Windows.Forms.GroupBox
  Friend WithEvents cmbGender As Windows.Forms.ComboBox
  Friend WithEvents lblGender As Windows.Forms.Label
  Friend WithEvents GroupBox6 As Windows.Forms.GroupBox
  Friend WithEvents txtNotes As Windows.Forms.TextBox
  Friend WithEvents Label12 As Windows.Forms.Label
  Friend WithEvents GroupBox4 As Windows.Forms.GroupBox
  Friend WithEvents dgvResourceNameValue As Windows.Forms.DataGridView
  Friend WithEvents GroupBox7 As Windows.Forms.GroupBox
  Friend WithEvents cmbResourceListItemNames As Windows.Forms.ComboBox
  Friend WithEvents lstListItems As Windows.Forms.ListView
  Friend WithEvents Label4 As Windows.Forms.Label
  Friend WithEvents Label5 As Windows.Forms.Label
  Friend WithEvents btnCancelListItems As Windows.Forms.Button
  Friend WithEvents btnSaveListItems As Windows.Forms.Button
End Class
