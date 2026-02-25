<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()>
Partial Class Availability
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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(Availability))
        Me.btnCancel = New System.Windows.Forms.Button()
        Me.grpMode = New System.Windows.Forms.GroupBox()
        Me.PictureBox4 = New System.Windows.Forms.PictureBox()
        Me.optAvailable = New System.Windows.Forms.RadioButton()
        Me.optUnavailable = New System.Windows.Forms.RadioButton()
        Me.grpTime = New System.Windows.Forms.GroupBox()
        Me.PictureBox2 = New System.Windows.Forms.PictureBox()
        Me.chkAllDay = New System.Windows.Forms.CheckBox()
        Me.lblEnd = New System.Windows.Forms.Label()
        Me.lblStarteTime = New System.Windows.Forms.Label()
        Me.dtpEndTime = New System.Windows.Forms.DateTimePicker()
        Me.dtpStartTime = New System.Windows.Forms.DateTimePicker()
        Me.grpPattern = New System.Windows.Forms.GroupBox()
        Me.PictureBox3 = New System.Windows.Forms.PictureBox()
        Me.grpPatternOption = New System.Windows.Forms.GroupBox()
        Me.optMonthly = New System.Windows.Forms.RadioButton()
        Me.optWeekly = New System.Windows.Forms.RadioButton()
        Me.optDateRange = New System.Windows.Forms.RadioButton()
    Me.tabPattern = New ResourceManagement.CustomControls.TabControlNoTabs()
    Me.pagDateRange = New System.Windows.Forms.TabPage()
        Me.dtpEndDate = New System.Windows.Forms.DateTimePicker()
        Me.lblEndDate = New System.Windows.Forms.Label()
        Me.dtpStartDate = New System.Windows.Forms.DateTimePicker()
        Me.lblStartDate = New System.Windows.Forms.Label()
        Me.pagWeekly = New System.Windows.Forms.TabPage()
        Me.chkSunday = New System.Windows.Forms.CheckBox()
        Me.chkSaturday = New System.Windows.Forms.CheckBox()
        Me.chkFriday = New System.Windows.Forms.CheckBox()
        Me.chkThursday = New System.Windows.Forms.CheckBox()
        Me.chkWednesday = New System.Windows.Forms.CheckBox()
        Me.chkTuesday = New System.Windows.Forms.CheckBox()
        Me.chkMonday = New System.Windows.Forms.CheckBox()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.nudRecurWeeks = New System.Windows.Forms.NumericUpDown()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.pagMonthly = New System.Windows.Forms.TabPage()
        Me.Label8 = New System.Windows.Forms.Label()
        Me.Label7 = New System.Windows.Forms.Label()
        Me.nudRecurMonthsWeek = New System.Windows.Forms.NumericUpDown()
        Me.cmbOrdinalMonthsWeek = New System.Windows.Forms.ComboBox()
        Me.Label6 = New System.Windows.Forms.Label()
        Me.nudRecurMonthsDay = New System.Windows.Forms.NumericUpDown()
        Me.Label5 = New System.Windows.Forms.Label()
        Me.cmbMonthlyDayOfWeek = New System.Windows.Forms.ComboBox()
        Me.cmbOrdinalMonthsDay = New System.Windows.Forms.ComboBox()
        Me.Label4 = New System.Windows.Forms.Label()
        Me.nudRecurMonthsDate = New System.Windows.Forms.NumericUpDown()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.nudDayOfMonth = New System.Windows.Forms.NumericUpDown()
        Me.grpMonthlyOption = New System.Windows.Forms.GroupBox()
        Me.optMonthlyWeek = New System.Windows.Forms.RadioButton()
        Me.optMonthlyDay = New System.Windows.Forms.RadioButton()
        Me.optMonthlyDate = New System.Windows.Forms.RadioButton()
        Me.grpRangeofPattern = New System.Windows.Forms.GroupBox()
        Me.PictureBox1 = New System.Windows.Forms.PictureBox()
        Me.Label10 = New System.Windows.Forms.Label()
        Me.nudRangeEndAfter = New System.Windows.Forms.NumericUpDown()
        Me.optRangeNoEndDate = New System.Windows.Forms.RadioButton()
        Me.optRangeEndAfter = New System.Windows.Forms.RadioButton()
        Me.optRangeEndBy = New System.Windows.Forms.RadioButton()
        Me.dtpRangeEndDate = New System.Windows.Forms.DateTimePicker()
        Me.dtpRangeStartDate = New System.Windows.Forms.DateTimePicker()
        Me.Label9 = New System.Windows.Forms.Label()
        Me.btnDelete = New System.Windows.Forms.Button()
        Me.btnSave = New System.Windows.Forms.Button()
        Me.ToolTip = New System.Windows.Forms.ToolTip(Me.components)
        Me.grpMode.SuspendLayout()
        CType(Me.PictureBox4, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.grpTime.SuspendLayout()
        CType(Me.PictureBox2, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.grpPattern.SuspendLayout()
        CType(Me.PictureBox3, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.grpPatternOption.SuspendLayout()
        Me.tabPattern.SuspendLayout()
        Me.pagDateRange.SuspendLayout()
        Me.pagWeekly.SuspendLayout()
        CType(Me.nudRecurWeeks, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.pagMonthly.SuspendLayout()
        CType(Me.nudRecurMonthsWeek, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.nudRecurMonthsDay, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.nudRecurMonthsDate, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.nudDayOfMonth, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.grpMonthlyOption.SuspendLayout()
        Me.grpRangeofPattern.SuspendLayout()
        CType(Me.PictureBox1, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.nudRangeEndAfter, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'btnCancel
        '
        Me.btnCancel.Image = Global.ResourceManagement.My.Resources.Resources._24close
        Me.btnCancel.Location = New System.Drawing.Point(601, 401)
        Me.btnCancel.Margin = New System.Windows.Forms.Padding(1)
        Me.btnCancel.Name = "btnCancel"
        Me.btnCancel.Size = New System.Drawing.Size(32, 32)
        Me.btnCancel.TabIndex = 0
        Me.btnCancel.UseVisualStyleBackColor = True
        '
        'grpMode
        '
        Me.grpMode.Controls.Add(Me.PictureBox4)
        Me.grpMode.Controls.Add(Me.optAvailable)
        Me.grpMode.Controls.Add(Me.optUnavailable)
        Me.grpMode.Location = New System.Drawing.Point(14, 3)
        Me.grpMode.Margin = New System.Windows.Forms.Padding(4)
        Me.grpMode.Name = "grpMode"
        Me.grpMode.Padding = New System.Windows.Forms.Padding(4)
        Me.grpMode.Size = New System.Drawing.Size(126, 105)
        Me.grpMode.TabIndex = 2
        Me.grpMode.TabStop = False
        Me.grpMode.Text = "Mode"
        '
        'PictureBox4
        '
        Me.PictureBox4.Image = Global.ResourceManagement.My.Resources.Resources._16toggle
        Me.PictureBox4.Location = New System.Drawing.Point(7, 20)
        Me.PictureBox4.Name = "PictureBox4"
        Me.PictureBox4.Size = New System.Drawing.Size(24, 20)
        Me.PictureBox4.SizeMode = System.Windows.Forms.PictureBoxSizeMode.CenterImage
        Me.PictureBox4.TabIndex = 42
        Me.PictureBox4.TabStop = False
        '
        'optAvailable
        '
        Me.optAvailable.AutoSize = True
        Me.optAvailable.Location = New System.Drawing.Point(15, 76)
        Me.optAvailable.Margin = New System.Windows.Forms.Padding(4)
        Me.optAvailable.Name = "optAvailable"
        Me.optAvailable.Size = New System.Drawing.Size(78, 21)
        Me.optAvailable.TabIndex = 1
        Me.optAvailable.TabStop = True
        Me.optAvailable.Text = "Available"
        Me.optAvailable.UseVisualStyleBackColor = True
        '
        'optUnavailable
        '
        Me.optUnavailable.AutoSize = True
        Me.optUnavailable.Checked = True
        Me.optUnavailable.Font = New System.Drawing.Font("Segoe UI", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.optUnavailable.Location = New System.Drawing.Point(15, 47)
        Me.optUnavailable.Margin = New System.Windows.Forms.Padding(4)
        Me.optUnavailable.Name = "optUnavailable"
        Me.optUnavailable.Size = New System.Drawing.Size(93, 21)
        Me.optUnavailable.TabIndex = 0
        Me.optUnavailable.TabStop = True
        Me.optUnavailable.Text = "Unavailable"
        Me.optUnavailable.UseVisualStyleBackColor = True
        '
        'grpTime
        '
        Me.grpTime.Controls.Add(Me.PictureBox2)
        Me.grpTime.Controls.Add(Me.chkAllDay)
        Me.grpTime.Controls.Add(Me.lblEnd)
        Me.grpTime.Controls.Add(Me.lblStarteTime)
        Me.grpTime.Controls.Add(Me.dtpEndTime)
        Me.grpTime.Controls.Add(Me.dtpStartTime)
        Me.grpTime.Location = New System.Drawing.Point(148, 3)
        Me.grpTime.Margin = New System.Windows.Forms.Padding(4)
        Me.grpTime.Name = "grpTime"
        Me.grpTime.Padding = New System.Windows.Forms.Padding(4)
        Me.grpTime.Size = New System.Drawing.Size(485, 105)
        Me.grpTime.TabIndex = 3
        Me.grpTime.TabStop = False
        Me.grpTime.Text = "Time"
        '
        'PictureBox2
        '
        Me.PictureBox2.Image = Global.ResourceManagement.My.Resources.Resources._16clock
        Me.PictureBox2.Location = New System.Drawing.Point(7, 25)
        Me.PictureBox2.Name = "PictureBox2"
        Me.PictureBox2.Size = New System.Drawing.Size(24, 20)
        Me.PictureBox2.SizeMode = System.Windows.Forms.PictureBoxSizeMode.CenterImage
        Me.PictureBox2.TabIndex = 41
        Me.PictureBox2.TabStop = False
        '
        'chkAllDay
        '
        Me.chkAllDay.AutoSize = True
        Me.chkAllDay.Location = New System.Drawing.Point(62, 69)
        Me.chkAllDay.Margin = New System.Windows.Forms.Padding(4)
        Me.chkAllDay.Name = "chkAllDay"
        Me.chkAllDay.Size = New System.Drawing.Size(67, 21)
        Me.chkAllDay.TabIndex = 4
        Me.chkAllDay.Text = "All Day"
        Me.chkAllDay.UseVisualStyleBackColor = True
        '
        'lblEnd
        '
        Me.lblEnd.AutoSize = True
        Me.lblEnd.Location = New System.Drawing.Point(246, 30)
        Me.lblEnd.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.lblEnd.Name = "lblEnd"
        Me.lblEnd.Size = New System.Drawing.Size(33, 17)
        Me.lblEnd.TabIndex = 3
        Me.lblEnd.Text = "End:"
        '
        'lblStarteTime
        '
        Me.lblStarteTime.AutoSize = True
        Me.lblStarteTime.Location = New System.Drawing.Point(57, 32)
        Me.lblStarteTime.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.lblStarteTime.Name = "lblStarteTime"
        Me.lblStarteTime.Size = New System.Drawing.Size(38, 17)
        Me.lblStarteTime.TabIndex = 2
        Me.lblStarteTime.Text = "Start:"
        '
        'dtpEndTime
        '
        Me.dtpEndTime.Format = System.Windows.Forms.DateTimePickerFormat.Time
        Me.dtpEndTime.Location = New System.Drawing.Point(287, 26)
        Me.dtpEndTime.Margin = New System.Windows.Forms.Padding(4)
        Me.dtpEndTime.Name = "dtpEndTime"
        Me.dtpEndTime.ShowUpDown = True
        Me.dtpEndTime.Size = New System.Drawing.Size(119, 25)
        Me.dtpEndTime.TabIndex = 1
        '
        'dtpStartTime
        '
        Me.dtpStartTime.Format = System.Windows.Forms.DateTimePickerFormat.Time
        Me.dtpStartTime.Location = New System.Drawing.Point(101, 28)
        Me.dtpStartTime.Margin = New System.Windows.Forms.Padding(4)
        Me.dtpStartTime.Name = "dtpStartTime"
        Me.dtpStartTime.ShowUpDown = True
        Me.dtpStartTime.Size = New System.Drawing.Size(119, 25)
        Me.dtpStartTime.TabIndex = 0
        '
        'grpPattern
        '
        Me.grpPattern.Controls.Add(Me.PictureBox3)
        Me.grpPattern.Controls.Add(Me.grpPatternOption)
        Me.grpPattern.Controls.Add(Me.tabPattern)
        Me.grpPattern.Location = New System.Drawing.Point(13, 108)
        Me.grpPattern.Margin = New System.Windows.Forms.Padding(4)
        Me.grpPattern.Name = "grpPattern"
        Me.grpPattern.Padding = New System.Windows.Forms.Padding(4)
        Me.grpPattern.Size = New System.Drawing.Size(620, 159)
        Me.grpPattern.TabIndex = 4
        Me.grpPattern.TabStop = False
        Me.grpPattern.Text = "Pattern"
        '
        'PictureBox3
        '
        Me.PictureBox3.Image = Global.ResourceManagement.My.Resources.Resources._16pattern
        Me.PictureBox3.Location = New System.Drawing.Point(8, 24)
        Me.PictureBox3.Name = "PictureBox3"
        Me.PictureBox3.Size = New System.Drawing.Size(24, 20)
        Me.PictureBox3.SizeMode = System.Windows.Forms.PictureBoxSizeMode.CenterImage
        Me.PictureBox3.TabIndex = 42
        Me.PictureBox3.TabStop = False
        '
        'grpPatternOption
        '
        Me.grpPatternOption.Controls.Add(Me.optMonthly)
        Me.grpPatternOption.Controls.Add(Me.optWeekly)
        Me.grpPatternOption.Controls.Add(Me.optDateRange)
        Me.grpPatternOption.Location = New System.Drawing.Point(39, 38)
        Me.grpPatternOption.Margin = New System.Windows.Forms.Padding(4)
        Me.grpPatternOption.Name = "grpPatternOption"
        Me.grpPatternOption.Padding = New System.Windows.Forms.Padding(4)
        Me.grpPatternOption.Size = New System.Drawing.Size(108, 105)
        Me.grpPatternOption.TabIndex = 0
        Me.grpPatternOption.TabStop = False
        '
        'optMonthly
        '
        Me.optMonthly.AutoSize = True
        Me.optMonthly.Location = New System.Drawing.Point(8, 75)
        Me.optMonthly.Margin = New System.Windows.Forms.Padding(4)
        Me.optMonthly.Name = "optMonthly"
        Me.optMonthly.Size = New System.Drawing.Size(73, 21)
        Me.optMonthly.TabIndex = 2
        Me.optMonthly.TabStop = True
        Me.optMonthly.Text = "Monthly"
        Me.optMonthly.UseVisualStyleBackColor = True
        '
        'optWeekly
        '
        Me.optWeekly.AutoSize = True
        Me.optWeekly.Location = New System.Drawing.Point(8, 46)
        Me.optWeekly.Margin = New System.Windows.Forms.Padding(4)
        Me.optWeekly.Name = "optWeekly"
        Me.optWeekly.Size = New System.Drawing.Size(66, 21)
        Me.optWeekly.TabIndex = 1
        Me.optWeekly.TabStop = True
        Me.optWeekly.Text = "Weekly"
        Me.optWeekly.UseVisualStyleBackColor = True
        '
        'optDateRange
        '
        Me.optDateRange.AutoSize = True
        Me.optDateRange.Checked = True
        Me.optDateRange.Location = New System.Drawing.Point(8, 17)
        Me.optDateRange.Margin = New System.Windows.Forms.Padding(4)
        Me.optDateRange.Name = "optDateRange"
        Me.optDateRange.Size = New System.Drawing.Size(94, 21)
        Me.optDateRange.TabIndex = 0
        Me.optDateRange.TabStop = True
        Me.optDateRange.Text = "Date Range"
        Me.optDateRange.UseVisualStyleBackColor = True
        '
        'tabPattern
        '
        Me.tabPattern.Controls.Add(Me.pagDateRange)
        Me.tabPattern.Controls.Add(Me.pagWeekly)
        Me.tabPattern.Controls.Add(Me.pagMonthly)
        Me.tabPattern.Location = New System.Drawing.Point(155, 39)
        Me.tabPattern.Margin = New System.Windows.Forms.Padding(4)
        Me.tabPattern.Name = "tabPattern"
        Me.tabPattern.SelectedIndex = 0
        Me.tabPattern.Size = New System.Drawing.Size(457, 109)
        Me.tabPattern.TabIndex = 1
        '
        'pagDateRange
        '
        Me.pagDateRange.BackColor = System.Drawing.SystemColors.Control
        Me.pagDateRange.Controls.Add(Me.dtpEndDate)
        Me.pagDateRange.Controls.Add(Me.lblEndDate)
        Me.pagDateRange.Controls.Add(Me.dtpStartDate)
        Me.pagDateRange.Controls.Add(Me.lblStartDate)
        Me.pagDateRange.Location = New System.Drawing.Point(0, -20)
        Me.pagDateRange.Margin = New System.Windows.Forms.Padding(4)
        Me.pagDateRange.Name = "pagDateRange"
        Me.pagDateRange.Padding = New System.Windows.Forms.Padding(4)
        Me.pagDateRange.Size = New System.Drawing.Size(457, 129)
        Me.pagDateRange.TabIndex = 0
        Me.pagDateRange.Text = "Date Range"
        '
        'dtpEndDate
        '
        Me.dtpEndDate.Checked = False
        Me.dtpEndDate.Format = System.Windows.Forms.DateTimePickerFormat.[Short]
        Me.dtpEndDate.Location = New System.Drawing.Point(288, 36)
        Me.dtpEndDate.Margin = New System.Windows.Forms.Padding(4)
        Me.dtpEndDate.Name = "dtpEndDate"
        Me.dtpEndDate.ShowCheckBox = True
        Me.dtpEndDate.Size = New System.Drawing.Size(119, 25)
        Me.dtpEndDate.TabIndex = 3
        '
        'lblEndDate
        '
        Me.lblEndDate.AutoSize = True
        Me.lblEndDate.Location = New System.Drawing.Point(247, 40)
        Me.lblEndDate.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.lblEndDate.Name = "lblEndDate"
        Me.lblEndDate.Size = New System.Drawing.Size(33, 17)
        Me.lblEndDate.TabIndex = 2
        Me.lblEndDate.Text = "End:"
        '
        'dtpStartDate
        '
        Me.dtpStartDate.Checked = False
        Me.dtpStartDate.Format = System.Windows.Forms.DateTimePickerFormat.[Short]
        Me.dtpStartDate.Location = New System.Drawing.Point(65, 36)
        Me.dtpStartDate.Margin = New System.Windows.Forms.Padding(4)
        Me.dtpStartDate.Name = "dtpStartDate"
        Me.dtpStartDate.ShowCheckBox = True
        Me.dtpStartDate.Size = New System.Drawing.Size(119, 25)
        Me.dtpStartDate.TabIndex = 1
        '
        'lblStartDate
        '
        Me.lblStartDate.AutoSize = True
        Me.lblStartDate.Location = New System.Drawing.Point(19, 40)
        Me.lblStartDate.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.lblStartDate.Name = "lblStartDate"
        Me.lblStartDate.Size = New System.Drawing.Size(38, 17)
        Me.lblStartDate.TabIndex = 0
        Me.lblStartDate.Text = "Start:"
        '
        'pagWeekly
        '
        Me.pagWeekly.BackColor = System.Drawing.SystemColors.Control
        Me.pagWeekly.Controls.Add(Me.chkSunday)
        Me.pagWeekly.Controls.Add(Me.chkSaturday)
        Me.pagWeekly.Controls.Add(Me.chkFriday)
        Me.pagWeekly.Controls.Add(Me.chkThursday)
        Me.pagWeekly.Controls.Add(Me.chkWednesday)
        Me.pagWeekly.Controls.Add(Me.chkTuesday)
        Me.pagWeekly.Controls.Add(Me.chkMonday)
        Me.pagWeekly.Controls.Add(Me.Label2)
        Me.pagWeekly.Controls.Add(Me.nudRecurWeeks)
        Me.pagWeekly.Controls.Add(Me.Label1)
        Me.pagWeekly.Location = New System.Drawing.Point(0, -20)
        Me.pagWeekly.Margin = New System.Windows.Forms.Padding(4)
        Me.pagWeekly.Name = "pagWeekly"
        Me.pagWeekly.Padding = New System.Windows.Forms.Padding(4)
        Me.pagWeekly.Size = New System.Drawing.Size(457, 129)
        Me.pagWeekly.TabIndex = 1
        Me.pagWeekly.Text = "Weekly"
        '
        'chkSunday
        '
        Me.chkSunday.AutoSize = True
        Me.chkSunday.Location = New System.Drawing.Point(203, 99)
        Me.chkSunday.Margin = New System.Windows.Forms.Padding(4)
        Me.chkSunday.Name = "chkSunday"
        Me.chkSunday.Size = New System.Drawing.Size(69, 21)
        Me.chkSunday.TabIndex = 9
        Me.chkSunday.Text = "Sunday"
        Me.chkSunday.UseVisualStyleBackColor = True
        '
        'chkSaturday
        '
        Me.chkSaturday.AutoSize = True
        Me.chkSaturday.Location = New System.Drawing.Point(108, 99)
        Me.chkSaturday.Margin = New System.Windows.Forms.Padding(4)
        Me.chkSaturday.Name = "chkSaturday"
        Me.chkSaturday.Size = New System.Drawing.Size(78, 21)
        Me.chkSaturday.TabIndex = 8
        Me.chkSaturday.Text = "Saturday"
        Me.chkSaturday.UseVisualStyleBackColor = True
        '
        'chkFriday
        '
        Me.chkFriday.AutoSize = True
        Me.chkFriday.Location = New System.Drawing.Point(16, 99)
        Me.chkFriday.Margin = New System.Windows.Forms.Padding(4)
        Me.chkFriday.Name = "chkFriday"
        Me.chkFriday.Size = New System.Drawing.Size(62, 21)
        Me.chkFriday.TabIndex = 7
        Me.chkFriday.Text = "Friday"
        Me.chkFriday.UseVisualStyleBackColor = True
        '
        'chkThursday
        '
        Me.chkThursday.AutoSize = True
        Me.chkThursday.Location = New System.Drawing.Point(323, 70)
        Me.chkThursday.Margin = New System.Windows.Forms.Padding(4)
        Me.chkThursday.Name = "chkThursday"
        Me.chkThursday.Size = New System.Drawing.Size(80, 21)
        Me.chkThursday.TabIndex = 6
        Me.chkThursday.Text = "Thursday"
        Me.chkThursday.UseVisualStyleBackColor = True
        '
        'chkWednesday
        '
        Me.chkWednesday.AutoSize = True
        Me.chkWednesday.Location = New System.Drawing.Point(203, 70)
        Me.chkWednesday.Margin = New System.Windows.Forms.Padding(4)
        Me.chkWednesday.Name = "chkWednesday"
        Me.chkWednesday.Size = New System.Drawing.Size(94, 21)
        Me.chkWednesday.TabIndex = 5
        Me.chkWednesday.Text = "Wednesday"
        Me.chkWednesday.UseVisualStyleBackColor = True
        '
        'chkTuesday
        '
        Me.chkTuesday.AutoSize = True
        Me.chkTuesday.Location = New System.Drawing.Point(108, 70)
        Me.chkTuesday.Margin = New System.Windows.Forms.Padding(4)
        Me.chkTuesday.Name = "chkTuesday"
        Me.chkTuesday.Size = New System.Drawing.Size(75, 21)
        Me.chkTuesday.TabIndex = 4
        Me.chkTuesday.Text = "Tuesday"
        Me.chkTuesday.UseVisualStyleBackColor = True
        '
        'chkMonday
        '
        Me.chkMonday.AutoSize = True
        Me.chkMonday.Location = New System.Drawing.Point(16, 70)
        Me.chkMonday.Margin = New System.Windows.Forms.Padding(4)
        Me.chkMonday.Name = "chkMonday"
        Me.chkMonday.Size = New System.Drawing.Size(75, 21)
        Me.chkMonday.TabIndex = 3
        Me.chkMonday.Text = "Monday"
        Me.chkMonday.UseVisualStyleBackColor = True
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(159, 39)
        Me.Label2.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(73, 17)
        Me.Label2.TabIndex = 2
        Me.Label2.Text = "week(s) on:"
        '
        'nudRecurWeeks
        '
        Me.nudRecurWeeks.Location = New System.Drawing.Point(101, 37)
        Me.nudRecurWeeks.Margin = New System.Windows.Forms.Padding(4)
        Me.nudRecurWeeks.Name = "nudRecurWeeks"
        Me.nudRecurWeeks.Size = New System.Drawing.Size(48, 25)
        Me.nudRecurWeeks.TabIndex = 1
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(13, 38)
        Me.Label1.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(76, 17)
        Me.Label1.TabIndex = 0
        Me.Label1.Text = "Recur every"
        '
        'pagMonthly
        '
        Me.pagMonthly.BackColor = System.Drawing.SystemColors.Control
        Me.pagMonthly.Controls.Add(Me.Label8)
        Me.pagMonthly.Controls.Add(Me.Label7)
        Me.pagMonthly.Controls.Add(Me.nudRecurMonthsWeek)
        Me.pagMonthly.Controls.Add(Me.cmbOrdinalMonthsWeek)
        Me.pagMonthly.Controls.Add(Me.Label6)
        Me.pagMonthly.Controls.Add(Me.nudRecurMonthsDay)
        Me.pagMonthly.Controls.Add(Me.Label5)
        Me.pagMonthly.Controls.Add(Me.cmbMonthlyDayOfWeek)
        Me.pagMonthly.Controls.Add(Me.cmbOrdinalMonthsDay)
        Me.pagMonthly.Controls.Add(Me.Label4)
        Me.pagMonthly.Controls.Add(Me.nudRecurMonthsDate)
        Me.pagMonthly.Controls.Add(Me.Label3)
        Me.pagMonthly.Controls.Add(Me.nudDayOfMonth)
        Me.pagMonthly.Controls.Add(Me.grpMonthlyOption)
        Me.pagMonthly.Location = New System.Drawing.Point(0, -20)
        Me.pagMonthly.Margin = New System.Windows.Forms.Padding(4)
        Me.pagMonthly.Name = "pagMonthly"
        Me.pagMonthly.Size = New System.Drawing.Size(457, 129)
        Me.pagMonthly.TabIndex = 2
        Me.pagMonthly.Text = "Monthly"
        '
        'Label8
        '
        Me.Label8.AutoSize = True
        Me.Label8.Location = New System.Drawing.Point(181, 99)
        Me.Label8.Name = "Label8"
        Me.Label8.Size = New System.Drawing.Size(88, 17)
        Me.Label8.TabIndex = 13
        Me.Label8.Text = "week of every"
        '
        'Label7
        '
        Me.Label7.AutoSize = True
        Me.Label7.Location = New System.Drawing.Point(329, 99)
        Me.Label7.Name = "Label7"
        Me.Label7.Size = New System.Drawing.Size(59, 17)
        Me.Label7.TabIndex = 12
        Me.Label7.Text = "month(s)"
        '
        'nudRecurMonthsWeek
        '
        Me.nudRecurMonthsWeek.Location = New System.Drawing.Point(275, 96)
        Me.nudRecurMonthsWeek.Maximum = New Decimal(New Integer() {36, 0, 0, 0})
        Me.nudRecurMonthsWeek.Name = "nudRecurMonthsWeek"
        Me.nudRecurMonthsWeek.Size = New System.Drawing.Size(48, 25)
        Me.nudRecurMonthsWeek.TabIndex = 11
        '
        'cmbOrdinalMonthsWeek
        '
        Me.cmbOrdinalMonthsWeek.FormattingEnabled = True
        Me.cmbOrdinalMonthsWeek.Location = New System.Drawing.Point(79, 95)
        Me.cmbOrdinalMonthsWeek.Name = "cmbOrdinalMonthsWeek"
        Me.cmbOrdinalMonthsWeek.Size = New System.Drawing.Size(96, 25)
        Me.cmbOrdinalMonthsWeek.TabIndex = 10
        '
        'Label6
        '
        Me.Label6.AutoSize = True
        Me.Label6.Location = New System.Drawing.Point(399, 67)
        Me.Label6.Name = "Label6"
        Me.Label6.Size = New System.Drawing.Size(59, 17)
        Me.Label6.TabIndex = 9
        Me.Label6.Text = "month(s)"
        '
        'nudRecurMonthsDay
        '
        Me.nudRecurMonthsDay.Location = New System.Drawing.Point(345, 61)
        Me.nudRecurMonthsDay.Maximum = New Decimal(New Integer() {36, 0, 0, 0})
        Me.nudRecurMonthsDay.Name = "nudRecurMonthsDay"
        Me.nudRecurMonthsDay.Size = New System.Drawing.Size(48, 25)
        Me.nudRecurMonthsDay.TabIndex = 8
        '
        'Label5
        '
        Me.Label5.AutoSize = True
        Me.Label5.Location = New System.Drawing.Point(284, 67)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(55, 17)
        Me.Label5.TabIndex = 7
        Me.Label5.Text = "of every"
        '
        'cmbMonthlyDayOfWeek
        '
        Me.cmbMonthlyDayOfWeek.FormattingEnabled = True
        Me.cmbMonthlyDayOfWeek.Location = New System.Drawing.Point(182, 64)
        Me.cmbMonthlyDayOfWeek.Name = "cmbMonthlyDayOfWeek"
        Me.cmbMonthlyDayOfWeek.Size = New System.Drawing.Size(96, 25)
        Me.cmbMonthlyDayOfWeek.TabIndex = 6
        '
        'cmbOrdinalMonthsDay
        '
        Me.cmbOrdinalMonthsDay.FormattingEnabled = True
        Me.cmbOrdinalMonthsDay.Location = New System.Drawing.Point(80, 64)
        Me.cmbOrdinalMonthsDay.Name = "cmbOrdinalMonthsDay"
        Me.cmbOrdinalMonthsDay.Size = New System.Drawing.Size(96, 25)
        Me.cmbOrdinalMonthsDay.TabIndex = 5
        '
        'Label4
        '
        Me.Label4.AutoSize = True
        Me.Label4.Location = New System.Drawing.Point(242, 34)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(59, 17)
        Me.Label4.TabIndex = 4
        Me.Label4.Text = "month(s)"
        '
        'nudRecurMonthsDate
        '
        Me.nudRecurMonthsDate.Location = New System.Drawing.Point(188, 32)
        Me.nudRecurMonthsDate.Maximum = New Decimal(New Integer() {36, 0, 0, 0})
        Me.nudRecurMonthsDate.Name = "nudRecurMonthsDate"
        Me.nudRecurMonthsDate.Size = New System.Drawing.Size(48, 25)
        Me.nudRecurMonthsDate.TabIndex = 3
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Location = New System.Drawing.Point(134, 34)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(55, 17)
        Me.Label3.TabIndex = 2
        Me.Label3.Text = "of every"
        '
        'nudDayOfMonth
        '
        Me.nudDayOfMonth.Location = New System.Drawing.Point(80, 32)
        Me.nudDayOfMonth.Maximum = New Decimal(New Integer() {31, 0, 0, 0})
        Me.nudDayOfMonth.Name = "nudDayOfMonth"
        Me.nudDayOfMonth.Size = New System.Drawing.Size(48, 25)
        Me.nudDayOfMonth.TabIndex = 1
        '
        'grpMonthlyOption
        '
        Me.grpMonthlyOption.Controls.Add(Me.optMonthlyWeek)
        Me.grpMonthlyOption.Controls.Add(Me.optMonthlyDay)
        Me.grpMonthlyOption.Controls.Add(Me.optMonthlyDate)
        Me.grpMonthlyOption.Location = New System.Drawing.Point(9, 20)
        Me.grpMonthlyOption.Name = "grpMonthlyOption"
        Me.grpMonthlyOption.Size = New System.Drawing.Size(64, 104)
        Me.grpMonthlyOption.TabIndex = 0
        Me.grpMonthlyOption.TabStop = False
        '
        'optMonthlyWeek
        '
        Me.optMonthlyWeek.AutoSize = True
        Me.optMonthlyWeek.Location = New System.Drawing.Point(6, 77)
        Me.optMonthlyWeek.Name = "optMonthlyWeek"
        Me.optMonthlyWeek.Size = New System.Drawing.Size(47, 21)
        Me.optMonthlyWeek.TabIndex = 2
        Me.optMonthlyWeek.TabStop = True
        Me.optMonthlyWeek.Text = "The"
        Me.optMonthlyWeek.UseVisualStyleBackColor = True
        '
        'optMonthlyDay
        '
        Me.optMonthlyDay.AutoSize = True
        Me.optMonthlyDay.Location = New System.Drawing.Point(6, 46)
        Me.optMonthlyDay.Name = "optMonthlyDay"
        Me.optMonthlyDay.Size = New System.Drawing.Size(47, 21)
        Me.optMonthlyDay.TabIndex = 1
        Me.optMonthlyDay.TabStop = True
        Me.optMonthlyDay.Text = "The"
        Me.optMonthlyDay.UseVisualStyleBackColor = True
        '
        'optMonthlyDate
        '
        Me.optMonthlyDate.AutoSize = True
        Me.optMonthlyDate.Location = New System.Drawing.Point(6, 14)
        Me.optMonthlyDate.Name = "optMonthlyDate"
        Me.optMonthlyDate.Size = New System.Drawing.Size(48, 21)
        Me.optMonthlyDate.TabIndex = 0
        Me.optMonthlyDate.TabStop = True
        Me.optMonthlyDate.Text = "Day"
        Me.optMonthlyDate.UseVisualStyleBackColor = True
        '
        'grpRangeofPattern
        '
        Me.grpRangeofPattern.Controls.Add(Me.PictureBox1)
        Me.grpRangeofPattern.Controls.Add(Me.Label10)
        Me.grpRangeofPattern.Controls.Add(Me.nudRangeEndAfter)
        Me.grpRangeofPattern.Controls.Add(Me.optRangeNoEndDate)
        Me.grpRangeofPattern.Controls.Add(Me.optRangeEndAfter)
        Me.grpRangeofPattern.Controls.Add(Me.optRangeEndBy)
        Me.grpRangeofPattern.Controls.Add(Me.dtpRangeEndDate)
        Me.grpRangeofPattern.Controls.Add(Me.dtpRangeStartDate)
        Me.grpRangeofPattern.Controls.Add(Me.Label9)
        Me.grpRangeofPattern.Location = New System.Drawing.Point(14, 274)
        Me.grpRangeofPattern.Name = "grpRangeofPattern"
        Me.grpRangeofPattern.Size = New System.Drawing.Size(619, 123)
        Me.grpRangeofPattern.TabIndex = 5
        Me.grpRangeofPattern.TabStop = False
        Me.grpRangeofPattern.Text = "Range of Pattern"
        '
        'PictureBox1
        '
        Me.PictureBox1.Image = Global.ResourceManagement.My.Resources.Resources._16leftandrightarrows
        Me.PictureBox1.Location = New System.Drawing.Point(8, 27)
        Me.PictureBox1.Name = "PictureBox1"
        Me.PictureBox1.Size = New System.Drawing.Size(24, 20)
        Me.PictureBox1.SizeMode = System.Windows.Forms.PictureBoxSizeMode.CenterImage
        Me.PictureBox1.TabIndex = 40
        Me.PictureBox1.TabStop = False
        '
        'Label10
        '
        Me.Label10.AutoSize = True
        Me.Label10.Location = New System.Drawing.Point(516, 59)
        Me.Label10.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.Label10.Name = "Label10"
        Me.Label10.Size = New System.Drawing.Size(78, 17)
        Me.Label10.TabIndex = 11
        Me.Label10.Text = "occurrences"
        '
        'nudRangeEndAfter
        '
        Me.nudRangeEndAfter.Location = New System.Drawing.Point(461, 57)
        Me.nudRangeEndAfter.Name = "nudRangeEndAfter"
        Me.nudRangeEndAfter.Size = New System.Drawing.Size(48, 25)
        Me.nudRangeEndAfter.TabIndex = 10
        '
        'optRangeNoEndDate
        '
        Me.optRangeNoEndDate.AutoSize = True
        Me.optRangeNoEndDate.Location = New System.Drawing.Point(345, 89)
        Me.optRangeNoEndDate.Name = "optRangeNoEndDate"
        Me.optRangeNoEndDate.Size = New System.Drawing.Size(100, 21)
        Me.optRangeNoEndDate.TabIndex = 9
        Me.optRangeNoEndDate.TabStop = True
        Me.optRangeNoEndDate.Text = "No end date"
        Me.optRangeNoEndDate.UseVisualStyleBackColor = True
        '
        'optRangeEndAfter
        '
        Me.optRangeEndAfter.AutoSize = True
        Me.optRangeEndAfter.Location = New System.Drawing.Point(345, 59)
        Me.optRangeEndAfter.Name = "optRangeEndAfter"
        Me.optRangeEndAfter.Size = New System.Drawing.Size(82, 21)
        Me.optRangeEndAfter.TabIndex = 8
        Me.optRangeEndAfter.TabStop = True
        Me.optRangeEndAfter.Text = "End after:"
        Me.optRangeEndAfter.UseVisualStyleBackColor = True
        '
        'optRangeEndBy
        '
        Me.optRangeEndBy.AutoSize = True
        Me.optRangeEndBy.Location = New System.Drawing.Point(345, 27)
        Me.optRangeEndBy.Name = "optRangeEndBy"
        Me.optRangeEndBy.Size = New System.Drawing.Size(69, 21)
        Me.optRangeEndBy.TabIndex = 7
        Me.optRangeEndBy.TabStop = True
        Me.optRangeEndBy.Text = "End by:"
        Me.optRangeEndBy.UseVisualStyleBackColor = True
        '
        'dtpRangeEndDate
        '
        Me.dtpRangeEndDate.Checked = False
        Me.dtpRangeEndDate.Format = System.Windows.Forms.DateTimePickerFormat.[Short]
        Me.dtpRangeEndDate.Location = New System.Drawing.Point(462, 25)
        Me.dtpRangeEndDate.Margin = New System.Windows.Forms.Padding(4)
        Me.dtpRangeEndDate.Name = "dtpRangeEndDate"
        Me.dtpRangeEndDate.ShowCheckBox = True
        Me.dtpRangeEndDate.Size = New System.Drawing.Size(119, 25)
        Me.dtpRangeEndDate.TabIndex = 6
        '
        'dtpRangeStartDate
        '
        Me.dtpRangeStartDate.Format = System.Windows.Forms.DateTimePickerFormat.[Short]
        Me.dtpRangeStartDate.Location = New System.Drawing.Point(197, 25)
        Me.dtpRangeStartDate.Margin = New System.Windows.Forms.Padding(4)
        Me.dtpRangeStartDate.Name = "dtpRangeStartDate"
        Me.dtpRangeStartDate.ShowCheckBox = True
        Me.dtpRangeStartDate.Size = New System.Drawing.Size(119, 25)
        Me.dtpRangeStartDate.TabIndex = 5
        '
        'Label9
        '
        Me.Label9.AutoSize = True
        Me.Label9.Location = New System.Drawing.Point(151, 31)
        Me.Label9.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.Label9.Name = "Label9"
        Me.Label9.Size = New System.Drawing.Size(38, 17)
        Me.Label9.TabIndex = 4
        Me.Label9.Text = "Start:"
        '
        'btnDelete
        '
        Me.btnDelete.Image = Global.ResourceManagement.My.Resources.Resources._24delete
        Me.btnDelete.Location = New System.Drawing.Point(567, 401)
        Me.btnDelete.Margin = New System.Windows.Forms.Padding(1)
        Me.btnDelete.Name = "btnDelete"
        Me.btnDelete.Size = New System.Drawing.Size(32, 32)
        Me.btnDelete.TabIndex = 6
        Me.btnDelete.UseVisualStyleBackColor = True
        '
        'btnSave
        '
        Me.btnSave.Image = Global.ResourceManagement.My.Resources.Resources._24save
        Me.btnSave.Location = New System.Drawing.Point(533, 401)
        Me.btnSave.Margin = New System.Windows.Forms.Padding(1)
        Me.btnSave.Name = "btnSave"
        Me.btnSave.Size = New System.Drawing.Size(32, 32)
        Me.btnSave.TabIndex = 7
        Me.btnSave.UseVisualStyleBackColor = True
        '
        'Availability
        '
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.None
        Me.ClientSize = New System.Drawing.Size(646, 441)
        Me.Controls.Add(Me.btnSave)
        Me.Controls.Add(Me.btnDelete)
        Me.Controls.Add(Me.grpRangeofPattern)
        Me.Controls.Add(Me.grpPattern)
        Me.Controls.Add(Me.grpTime)
        Me.Controls.Add(Me.grpMode)
        Me.Controls.Add(Me.btnCancel)
        Me.Font = New System.Drawing.Font("Segoe UI", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.Margin = New System.Windows.Forms.Padding(4)
        Me.Name = "Availability"
        Me.SizeGripStyle = System.Windows.Forms.SizeGripStyle.Hide
        Me.Text = "Availability"
        Me.grpMode.ResumeLayout(False)
        Me.grpMode.PerformLayout()
        CType(Me.PictureBox4, System.ComponentModel.ISupportInitialize).EndInit()
        Me.grpTime.ResumeLayout(False)
        Me.grpTime.PerformLayout()
        CType(Me.PictureBox2, System.ComponentModel.ISupportInitialize).EndInit()
        Me.grpPattern.ResumeLayout(False)
        CType(Me.PictureBox3, System.ComponentModel.ISupportInitialize).EndInit()
        Me.grpPatternOption.ResumeLayout(False)
        Me.grpPatternOption.PerformLayout()
        Me.tabPattern.ResumeLayout(False)
        Me.pagDateRange.ResumeLayout(False)
        Me.pagDateRange.PerformLayout()
        Me.pagWeekly.ResumeLayout(False)
        Me.pagWeekly.PerformLayout()
        CType(Me.nudRecurWeeks, System.ComponentModel.ISupportInitialize).EndInit()
        Me.pagMonthly.ResumeLayout(False)
        Me.pagMonthly.PerformLayout()
        CType(Me.nudRecurMonthsWeek, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.nudRecurMonthsDay, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.nudRecurMonthsDate, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.nudDayOfMonth, System.ComponentModel.ISupportInitialize).EndInit()
        Me.grpMonthlyOption.ResumeLayout(False)
        Me.grpMonthlyOption.PerformLayout()
        Me.grpRangeofPattern.ResumeLayout(False)
        Me.grpRangeofPattern.PerformLayout()
        CType(Me.PictureBox1, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.nudRangeEndAfter, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)

    End Sub

    Friend WithEvents btnCancel As Windows.Forms.Button
  Friend WithEvents tabPattern As ResourceManagement.CustomControls.TabControlNoTabs
  Friend WithEvents pagDateRange As Windows.Forms.TabPage
    Friend WithEvents pagWeekly As Windows.Forms.TabPage
    Friend WithEvents grpMode As Windows.Forms.GroupBox
    Friend WithEvents optAvailable As Windows.Forms.RadioButton
    Friend WithEvents optUnavailable As Windows.Forms.RadioButton
    Friend WithEvents grpTime As Windows.Forms.GroupBox
    Friend WithEvents dtpStartTime As Windows.Forms.DateTimePicker
    Friend WithEvents chkAllDay As Windows.Forms.CheckBox
    Friend WithEvents lblEnd As Windows.Forms.Label
    Friend WithEvents lblStarteTime As Windows.Forms.Label
    Friend WithEvents dtpEndTime As Windows.Forms.DateTimePicker
    Friend WithEvents grpPattern As Windows.Forms.GroupBox
    Friend WithEvents grpPatternOption As Windows.Forms.GroupBox
    Friend WithEvents optMonthly As Windows.Forms.RadioButton
    Friend WithEvents optWeekly As Windows.Forms.RadioButton
    Friend WithEvents optDateRange As Windows.Forms.RadioButton
    Friend WithEvents dtpStartDate As Windows.Forms.DateTimePicker
    Friend WithEvents lblStartDate As Windows.Forms.Label
    Friend WithEvents pagMonthly As Windows.Forms.TabPage
    Friend WithEvents dtpEndDate As Windows.Forms.DateTimePicker
    Friend WithEvents lblEndDate As Windows.Forms.Label
    Friend WithEvents nudRecurWeeks As Windows.Forms.NumericUpDown
    Friend WithEvents Label1 As Windows.Forms.Label
    Friend WithEvents chkSunday As Windows.Forms.CheckBox
    Friend WithEvents chkSaturday As Windows.Forms.CheckBox
    Friend WithEvents chkFriday As Windows.Forms.CheckBox
    Friend WithEvents chkThursday As Windows.Forms.CheckBox
    Friend WithEvents chkWednesday As Windows.Forms.CheckBox
    Friend WithEvents chkTuesday As Windows.Forms.CheckBox
    Friend WithEvents chkMonday As Windows.Forms.CheckBox
    Friend WithEvents Label2 As Windows.Forms.Label
    Friend WithEvents grpMonthlyOption As Windows.Forms.GroupBox
    Friend WithEvents optMonthlyDay As Windows.Forms.RadioButton
    Friend WithEvents optMonthlyDate As Windows.Forms.RadioButton
    Friend WithEvents Label4 As Windows.Forms.Label
    Friend WithEvents nudRecurMonthsDate As Windows.Forms.NumericUpDown
    Friend WithEvents Label3 As Windows.Forms.Label
    Friend WithEvents nudDayOfMonth As Windows.Forms.NumericUpDown
    Friend WithEvents optMonthlyWeek As Windows.Forms.RadioButton
    Friend WithEvents cmbOrdinalMonthsDay As Windows.Forms.ComboBox
    Friend WithEvents Label6 As Windows.Forms.Label
    Friend WithEvents nudRecurMonthsDay As Windows.Forms.NumericUpDown
    Friend WithEvents Label5 As Windows.Forms.Label
    Friend WithEvents cmbMonthlyDayOfWeek As Windows.Forms.ComboBox
    Friend WithEvents Label8 As Windows.Forms.Label
    Friend WithEvents Label7 As Windows.Forms.Label
    Friend WithEvents nudRecurMonthsWeek As Windows.Forms.NumericUpDown
    Friend WithEvents cmbOrdinalMonthsWeek As Windows.Forms.ComboBox
    Friend WithEvents grpRangeofPattern As Windows.Forms.GroupBox
    Friend WithEvents dtpRangeEndDate As Windows.Forms.DateTimePicker
    Friend WithEvents dtpRangeStartDate As Windows.Forms.DateTimePicker
    Friend WithEvents Label9 As Windows.Forms.Label
    Friend WithEvents Label10 As Windows.Forms.Label
    Friend WithEvents nudRangeEndAfter As Windows.Forms.NumericUpDown
    Friend WithEvents optRangeNoEndDate As Windows.Forms.RadioButton
    Friend WithEvents optRangeEndAfter As Windows.Forms.RadioButton
    Friend WithEvents optRangeEndBy As Windows.Forms.RadioButton
    Friend WithEvents btnDelete As Windows.Forms.Button
    Friend WithEvents btnSave As Windows.Forms.Button
    Friend WithEvents ToolTip As Windows.Forms.ToolTip
    Friend WithEvents PictureBox1 As Windows.Forms.PictureBox
    Friend WithEvents PictureBox4 As Windows.Forms.PictureBox
    Friend WithEvents PictureBox2 As Windows.Forms.PictureBox
    Friend WithEvents PictureBox3 As Windows.Forms.PictureBox
End Class
