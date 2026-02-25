Imports System.Drawing
Imports System.Windows.Forms
Imports ResourceManagement.My

' ==========================================================================================
' Form: Availability
' Purpose:
'   Edit a single availability record for a resource using the UIModelAvailability contract.
'   Uses UILoaderSaverAvailability to load/save the normalised tables.
'
'   Friend properties:
'     - resourceID     : owning resource
'     - availabilityID : existing availability ("" for Add)
'     - wasSaved       : True when a successful save/delete occurred
'
' Notes:
'   - Snapshot/working copy pattern:
'       * _model.Availability        = snapshot
'       * _model.ActionAvailability  = working copy
'   - Form writes only to ActionAvailability.
' ==========================================================================================
Friend Class Availability

  ' === Friend form contract ===
  Friend Property wasSaved As Boolean
  Friend Property resourceID As String
  Friend Property availabilityID As String

  ' === Private model ===
  Private _model As UIModelAvailability

  Friend Sub New()
    Try
      ' Disable WinForms autoscaling completely
      Me.AutoScaleMode = AutoScaleMode.None
      InitializeComponent()

      ' Apply DPI scaling AFTER controls exist
      ApplyDpiScaling(Me)

    Catch ex As Exception
      ErrorHandler.UnHandleError(ex)

    Finally
      ' Cleanup
    End Try
  End Sub

  ' ==========================================================================================
  ' Routine: Availability_Load
  ' Purpose:
  '   Initialize the Availability form, load the UI model, and render the working copy into
  '   the controls.
  ' Parameters:
  '   sender - event sender
  '   e      - event args
  ' Returns:
  '   None
  ' Notes:
  '   - Uses UILoaderSaverAvailability.LoadAvailabilityModel.
  '   - For Add, assigns ResourceID on the working copy.
  ' ==========================================================================================
  Private Sub Availability_Load(sender As Object, e As EventArgs) Handles MyBase.Load

    Try
      ' Center relative to Excel (ignore if helper not available)
      Try
        FormHelpers.CenterFormOnExcel(Me)
      Catch
        ' ignore
      End Try

      InitialiseTooltips()
      InitialiseMonthlyCombos()

      ' Load model (snapshot + working copy)
      _model = UILoaderSaverAvailability.LoadAvailabilityModel(
                 If(String.IsNullOrEmpty(availabilityID), "", availabilityID))

      ' Ensure working copy has owning ResourceID
      _model.ActionAvailability.ResourceID = resourceID

      If String.IsNullOrEmpty(availabilityID) Then
        ' New availability
        InitialiseDefaults()
      Else
        ' Existing availability
        RenderModelToControls()
      End If

      UpdateRangeEndControls()

    Catch ex As Exception
      ErrorHandler.UnHandleError(ex)
    End Try

  End Sub


  ' ==========================================================================================
  ' Routine: InitialiseTooltips
  ' Purpose:
  '   Configure tooltip text for date controls using current short date mask.
  ' Parameters:
  '   None
  ' Returns:
  '   None
  ' ==========================================================================================
  Private Sub InitialiseTooltips()

    'Dim mask As String = GetShortDateMask()

    'ToolTip.SetToolTip(dtpStartDate, "Enter date as " & mask)
    'ToolTip.SetToolTip(dtpEndDate, "Enter date as " & mask)
    'ToolTip.SetToolTip(dtpRangeStartDate, "Enter date as " & mask)
    'ToolTip.SetToolTip(dtpRangeEndDate, "Enter date as " & mask)

  End Sub


  ' ==========================================================================================
  ' Routine: InitialiseMonthlyCombos
  ' Purpose:
  '   Populate ordinal and day-of-week combo boxes for monthly patterns.
  ' Parameters:
  '   None
  ' Returns:
  '   None
  ' ==========================================================================================
  Private Sub InitialiseMonthlyCombos()

    With cmbOrdinalMonthsDay.Items
      .Clear()
      .Add("first")
      .Add("second")
      .Add("third")
      .Add("fourth")
      .Add("last")
    End With

    With cmbOrdinalMonthsWeek.Items
      .Clear()
      .Add("first")
      .Add("second")
      .Add("third")
      .Add("fourth")
      .Add("last")
    End With

    With cmbMonthlyDayOfWeek.Items
      .Clear()
      .Add("Monday")
      .Add("Tuesday")
      .Add("Wednesday")
      .Add("Thursday")
      .Add("Friday")
      .Add("Saturday")
      .Add("Sunday")
    End With

  End Sub

  ' ==========================================================================================
  ' Routine: InitialiseDefaults
  ' Purpose:
  '   Apply default values for a new availability record.
  '   - Mode defaults to the OPPOSITE of the database DefaultMode
  '   - AllDay = True
  '   - PatternType = DateRange
  '   - StartDate = Today
  '   - RangeStartDate = Today
  '   - RangeEndType = None
  ' ==========================================================================================
  Private Sub InitialiseDefaults()

    Dim defaultMode As String = LoadMetadataValue("DefaultMode")

    ' Mode is opposite of DB default
    If String.Equals(defaultMode, "Available", StringComparison.OrdinalIgnoreCase) Then
      optUnavailable.Checked = True
    Else
      optAvailable.Checked = True
    End If

    ' All day
    chkAllDay.Checked = True

    ' Pattern defaults
    optDateRange.Checked = True
    tabPattern.SelectedIndex = 0

    ' DateRange defaults
    dtpStartDate.Checked = True
    dtpStartDate.Value = Date.Today

    dtpEndDate.Checked = True
    dtpEndDate.Value = Date.Today

    ' Range defaults
    dtpRangeStartDate.Checked = True
    dtpRangeStartDate.Value = Date.Today

    optRangeNoEndDate.Checked = True
    dtpRangeEndDate.Checked = False
    nudRangeEndAfter.Value = 1

  End Sub

  ' ==========================================================================================
  ' Routine: RenderModelToControls
  ' Purpose:
  '   Render the working copy (ActionAvailability) from the model into the form controls.
  ' Parameters:
  '   None
  ' Returns:
  '   None
  ' Notes:
  '   - Uses ActionAvailability, not the snapshot.
  ' ==========================================================================================
  Private Sub RenderModelToControls()

    Dim a As UIAvailabilityRow = _model.ActionAvailability

    ' --- Mode ---
    If String.Equals(a.Mode, "Unavailable", StringComparison.OrdinalIgnoreCase) Then
      optUnavailable.Checked = True
    Else
      optAvailable.Checked = True
    End If

    ' --- Time / All day ---
    chkAllDay.Checked = a.AllDay

    If a.AllDay Then
      ' Default times when all-day (times not really used)
      dtpStartTime.Value = Date.Today
      dtpEndTime.Value = Date.Today
    Else
      dtpStartTime.Value = Date.Today.Add(
        If(a.StartTime = TimeSpan.Zero, TimeSpan.FromHours(9), a.StartTime))
      dtpEndTime.Value = Date.Today.Add(
        If(a.EndTime = TimeSpan.Zero, TimeSpan.FromHours(17), a.EndTime))
    End If

    ' --- Pattern type + tab selection ---
    Select Case a.PatternType
      Case "Weekly"
        optWeekly.Checked = True
        tabPattern.SelectedIndex = 1
      Case "Monthly"
        optMonthly.Checked = True
        tabPattern.SelectedIndex = 2
      Case Else
        optDateRange.Checked = True
        tabPattern.SelectedIndex = 0
    End Select

    ' --- DateRange pattern tab ---
    If a.DateRangeStart <> Date.MinValue Then
      dtpStartDate.Checked = True
      dtpStartDate.Value = a.DateRangeStart
    Else
      dtpStartDate.Checked = False
    End If

    If a.DateRangeEnd <> Date.MinValue Then
      dtpEndDate.Checked = True
      dtpEndDate.Value = a.DateRangeEnd
    Else
      dtpEndDate.Checked = False
    End If

    ' --- Range of pattern ---
    If a.RangeStart <> Date.MinValue Then
      dtpRangeStartDate.Checked = True
      dtpRangeStartDate.Value = a.RangeStart
    Else
      dtpRangeStartDate.Checked = False
    End If

    Select Case a.RangeEndType
      Case "By"
        optRangeEndBy.Checked = True
        If a.RangeEndDate <> Date.MinValue Then
          dtpRangeEndDate.Checked = True
          dtpRangeEndDate.Value = a.RangeEndDate
        Else
          dtpRangeEndDate.Checked = False
        End If

      Case "After"
        optRangeEndAfter.Checked = True
        nudRangeEndAfter.Value = Math.Max(1, CDec(a.RangeEndAfterOccurrences))

      Case Else
        optRangeNoEndDate.Checked = True
        dtpRangeEndDate.Checked = False
        nudRangeEndAfter.Value = 1
    End Select

    ' --- Weekly pattern tab ---
    If a.RecurWeeks > 0 Then
      nudRecurWeeks.Value = CDec(a.RecurWeeks)
    Else
      nudRecurWeeks.Value = 1
    End If

    chkMonday.Checked = a.DaysOfWeek.Contains("Monday")
    chkTuesday.Checked = a.DaysOfWeek.Contains("Tuesday")
    chkWednesday.Checked = a.DaysOfWeek.Contains("Wednesday")
    chkThursday.Checked = a.DaysOfWeek.Contains("Thursday")
    chkFriday.Checked = a.DaysOfWeek.Contains("Friday")
    chkSaturday.Checked = a.DaysOfWeek.Contains("Saturday")
    chkSunday.Checked = a.DaysOfWeek.Contains("Sunday")

    ' --- Monthly pattern tab ---
    Select Case a.MonthlyType
      Case "DayOfMonth"
        optMonthlyDate.Checked = True
        nudDayOfMonth.Value = If(a.DayOfMonth > 0, CDec(a.DayOfMonth), 1D)
        nudRecurMonthsDate.Value = If(a.RecurMonths > 0, CDec(a.RecurMonths), 1D)

      Case "OrdinalDay"
        optMonthlyDay.Checked = True
        cmbOrdinalMonthsDay.Text = a.Ordinal
        cmbMonthlyDayOfWeek.Text = a.MonthlyDayOfWeek
        nudRecurMonthsDay.Value = If(a.RecurMonths > 0, CDec(a.RecurMonths), 1D)

      Case "OrdinalWeek"
        optMonthlyWeek.Checked = True
        cmbOrdinalMonthsWeek.Text = a.Ordinal
        nudRecurMonthsWeek.Value = If(a.RecurMonths > 0, CDec(a.RecurMonths), 1D)

      Case Else
        ' Default to date-based monthly if pattern type = Monthly but no subtype set
        If optMonthly.Checked Then
          optMonthlyDate.Checked = True
          nudDayOfMonth.Value = 1D
          nudRecurMonthsDate.Value = 1D
        End If
    End Select

  End Sub


  ' ==========================================================================================
  ' Routine: ApplyControlsToModel
  ' Purpose:
  '   Copy current control values into the working copy (ActionAvailability).
  ' Parameters:
  '   None
  ' Returns:
  '   None
  ' Notes:
  '   - Does NOT set PendingAction.
  '   - Caller must set model.PendingAction (Add / Update / Delete).
  ' ==========================================================================================
  Private Sub ApplyControlsToModel()

    Dim a As UIAvailabilityRow = _model.ActionAvailability

    ' Owning resource
    a.ResourceID = resourceID

    ' Mode
    If optUnavailable.Checked Then
      a.Mode = "Unavailable"
    Else
      a.Mode = "Available"
    End If

    ' PatternType
    If optDateRange.Checked Then
      a.PatternType = "DateRange"
    ElseIf optWeekly.Checked Then
      a.PatternType = "Weekly"
    ElseIf optMonthly.Checked Then
      a.PatternType = "Monthly"
    Else
      a.PatternType = ""
    End If

    ' All day / times
    a.AllDay = chkAllDay.Checked
    If a.AllDay Then
      a.StartTime = TimeSpan.Zero
      a.EndTime = TimeSpan.Zero
    Else
      a.StartTime = dtpStartTime.Value.TimeOfDay
      a.EndTime = dtpEndTime.Value.TimeOfDay
    End If

    ' DateRange pattern
    If dtpStartDate.Checked Then
      a.DateRangeStart = dtpStartDate.Value.Date
    Else
      a.DateRangeStart = Date.MinValue
    End If

    If dtpEndDate.Checked Then
      a.DateRangeEnd = dtpEndDate.Value.Date
    Else
      a.DateRangeEnd = Date.MinValue
    End If

    ' Range of pattern
    If dtpRangeStartDate.Checked Then
      a.RangeStart = dtpRangeStartDate.Value.Date
    Else
      a.RangeStart = Date.MinValue
    End If

    If optRangeEndBy.Checked Then
      a.RangeEndType = "By"
      If dtpRangeEndDate.Checked Then
        a.RangeEndDate = dtpRangeEndDate.Value.Date
      Else
        a.RangeEndDate = Date.MinValue
      End If
      a.RangeEndAfterOccurrences = 0

    ElseIf optRangeEndAfter.Checked Then
      a.RangeEndType = "After"
      a.RangeEndAfterOccurrences = CLng(nudRangeEndAfter.Value)
      a.RangeEndDate = Date.MinValue

    Else
      a.RangeEndType = "None"
      a.RangeEndDate = Date.MinValue
      a.RangeEndAfterOccurrences = 0
    End If

    ' Weekly pattern
    a.RecurWeeks = CLng(nudRecurWeeks.Value)
    a.DaysOfWeek.Clear()

    If chkMonday.Checked Then a.DaysOfWeek.Add("Monday")
    If chkTuesday.Checked Then a.DaysOfWeek.Add("Tuesday")
    If chkWednesday.Checked Then a.DaysOfWeek.Add("Wednesday")
    If chkThursday.Checked Then a.DaysOfWeek.Add("Thursday")
    If chkFriday.Checked Then a.DaysOfWeek.Add("Friday")
    If chkSaturday.Checked Then a.DaysOfWeek.Add("Saturday")
    If chkSunday.Checked Then a.DaysOfWeek.Add("Sunday")

    ' Monthly pattern
    Select Case True

      Case optMonthlyDate.Checked
        a.MonthlyType = "DayOfMonth"
        a.DayOfMonth = CLng(nudDayOfMonth.Value)
        a.RecurMonths = CLng(nudRecurMonthsDate.Value)
        a.Ordinal = ""
        a.MonthlyDayOfWeek = ""

      Case optMonthlyDay.Checked
        a.MonthlyType = "OrdinalDay"
        a.Ordinal = cmbOrdinalMonthsDay.Text
        a.MonthlyDayOfWeek = cmbMonthlyDayOfWeek.Text
        a.RecurMonths = CLng(nudRecurMonthsDay.Value)
        a.DayOfMonth = 0

      Case optMonthlyWeek.Checked
        a.MonthlyType = "OrdinalWeek"
        a.Ordinal = cmbOrdinalMonthsWeek.Text
        a.RecurMonths = CLng(nudRecurMonthsWeek.Value)
        a.DayOfMonth = 0
        ' MonthlyDayOfWeek is not used for OrdinalWeek pattern in this schema

      Case Else
        a.MonthlyType = ""
        a.DayOfMonth = 0
        a.Ordinal = ""
        a.MonthlyDayOfWeek = ""
        a.RecurMonths = 0

    End Select

  End Sub


  ' ==========================================================================================
  ' Routine: ValidateModelOrThrow
  ' Purpose:
  '   Validate the working availability model before saving.
  ' Parameters:
  '   None
  ' Returns:
  '   None
  ' Notes:
  '   - Throws UserFriendlyException on validation failure.
  '   - UI-level date validation already enforces basic ranges.
  ' ==========================================================================================
  Private Sub ValidateModelOrThrow()

    Dim a As UIAvailabilityRow = _model.ActionAvailability

    ' Time logic when not all-day
    If Not a.AllDay AndAlso a.StartTime >= a.EndTime Then
      Throw New UserFriendlyException("End time must be greater than start time.")
    End If

    ' PatternType must match required fields
    Select Case a.PatternType

      Case "DateRange"
        If a.DateRangeStart = Date.MinValue Or a.DateRangeEnd = Date.MinValue Then
          Throw New UserFriendlyException("Start and End date is required for a Date Range pattern.")
        End If

      Case "Weekly"
        If a.RecurWeeks < 1 Then
          Throw New UserFriendlyException("Recur every (weeks) must be at least 1.")
        End If
        If a.DaysOfWeek.Count = 0 Then
          Throw New UserFriendlyException("Select at least one day of the week.")
        End If

      Case "Monthly"
        Select Case a.MonthlyType

          Case "DayOfMonth"
            If a.DayOfMonth < 1 OrElse a.DayOfMonth > 31 Then
              Throw New UserFriendlyException("Day of month must be between 1 and 31.")
            End If
            If a.RecurMonths < 1 Then
              Throw New UserFriendlyException("Recur every (months) must be at least 1.")
            End If

          Case "OrdinalDay"
            If String.IsNullOrWhiteSpace(a.Ordinal) OrElse
           String.IsNullOrWhiteSpace(a.MonthlyDayOfWeek) Then
              Throw New UserFriendlyException("Select ordinal and day of week for monthly pattern.")
            End If

          Case "OrdinalWeek"
            If String.IsNullOrWhiteSpace(a.Ordinal) Then
              Throw New UserFriendlyException("Select ordinal for monthly pattern.")
            End If

          Case Else
            Throw New UserFriendlyException("Select a valid monthly pattern type.")

        End Select

      Case Else
        Throw New UserFriendlyException("Select a valid pattern type.")

    End Select

  End Sub


  ' ==========================================================================================
  ' Routine: btnSave_Click
  ' Purpose:
  '   Save button handler. Applies current UI to the model, validates, sets PendingAction,
  '   and invokes UILoaderSaverAvailability.SavePendingAvailabilityAction.
  ' Parameters:
  '   sender - event sender
  '   e      - event args
  ' Returns:
  '   None
  ' ==========================================================================================
  Private Sub btnSave_Click(sender As Object, e As EventArgs) Handles btnSave.Click

    Try
      ApplyControlsToModel()
      ValidateModelOrThrow()

      ' Set pending action based on current state
      If String.IsNullOrEmpty(_model.ActionAvailability.AvailabilityID) Then
        _model.PendingAction = AvailabilityAction.Add
      Else
        _model.PendingAction = AvailabilityAction.Update
      End If

      UILoaderSaverAvailability.SavePendingAvailabilityAction(_model)

      ' Propagate new/updated ID back to caller
      availabilityID = _model.ActionAvailability.AvailabilityID
      wasSaved = True
      Me.DialogResult = DialogResult.OK
      Me.Close()

    Catch ex As UserFriendlyException
      MessageBoxHelper.Show(UIMessageOwner.ExcelOwner,
                            ex.Message,
                            "Availability",
                            MessageBoxButtons.OK,
                            MessageBoxIcon.Warning)

    Catch ex As Exception
      ErrorHandler.UnHandleError(ex)
    End Try

  End Sub


  ' ==========================================================================================
  ' Routine: btnDelete_Click
  ' Purpose:
  '   Delete button handler. Marks the availability for delete and commits via saver.
  ' Parameters:
  '   sender - event sender
  '   e      - event args
  ' Returns:
  '   None
  ' Notes:
  '   - Soft-deletes availability and all pattern rows.
  ' ==========================================================================================
  Private Sub btnDelete_Click(sender As Object, e As EventArgs) Handles btnDelete.Click

    Try
      If String.IsNullOrEmpty(_model.ActionAvailability.AvailabilityID) Then
        ' Nothing persisted yet; treat as cancel
        Me.DialogResult = DialogResult.Cancel
        Me.Close()
        Return
      End If

      If MessageBoxHelper.Show(UIMessageOwner.ExcelOwner,
                               "Delete this availability?",
                               "Confirm Delete",
                               MessageBoxButtons.YesNo,
                               MessageBoxIcon.Question) <> DialogResult.Yes Then
        Return
      End If

      _model.PendingAction = AvailabilityAction.Delete
      UILoaderSaverAvailability.SavePendingAvailabilityAction(_model)

      wasSaved = True
      Me.DialogResult = DialogResult.OK
      Me.Close()

    Catch ex As Exception
      ErrorHandler.UnHandleError(ex)
    End Try

  End Sub


  ' ==========================================================================================
  ' Routine: btnCancel_Click
  ' Purpose:
  '   Cancel button handler. Closes the form without saving.
  ' Parameters:
  '   sender - event sender
  '   e      - event args
  ' Returns:
  '   None
  ' ==========================================================================================
  Private Sub btnCancel_Click(sender As Object, e As EventArgs) Handles btnCancel.Click
    Me.DialogResult = DialogResult.Cancel
    Me.Close()
  End Sub


  ' ==========================================================================================
  ' Routine: ComboBox_Validating
  ' Purpose:
  '   Handles validating event for all ComboBoxes to enforce valid selection from list.
  ' Parameters:
  '   sender - event sender
  '   e      - event args
  ' Returns:
  '   None
  ' ==========================================================================================
  Private Sub ComboBox_Validating(sender As Object, e As System.ComponentModel.CancelEventArgs) _
    Handles cmbOrdinalMonthsDay.Validating, cmbMonthlyDayOfWeek.Validating, cmbOrdinalMonthsWeek.Validating

    Try
      ValidateComboBoxSelection(DirectCast(sender, ComboBox), e)
    Catch ex As Exception
      ErrorHandler.UnHandleError(ex)
    End Try

  End Sub


  ' ==========================================================================================
  ' Routine: dtpStartDate_ValueChanged
  ' Purpose:
  '   Ensure EndDate is at least StartDate when StartDate changes.
  ' ==========================================================================================
  Private Sub dtpStartDate_ValueChanged(sender As Object, e As EventArgs) Handles dtpStartDate.ValueChanged
    Try
      If dtpStartDate.Checked AndAlso dtpEndDate.Checked Then
        If dtpEndDate.Value < dtpStartDate.Value Then
          dtpEndDate.Value = dtpStartDate.Value
        End If
      End If
    Catch ex As Exception
      ErrorHandler.UnHandleError(ex)
    End Try
  End Sub


  ' ==========================================================================================
  ' Routine: dtpEndDate_ValueChanged
  ' Purpose:
  '   Validate EndDate (EndDate must be >= StartDate).
  ' ==========================================================================================
  Private Sub dtpEndDate_ValueChanged(sender As Object, e As EventArgs) Handles dtpEndDate.ValueChanged
    Try
      If dtpStartDate.Checked AndAlso dtpEndDate.Checked Then
        If dtpEndDate.Value < dtpStartDate.Value Then
          MessageBoxHelper.Show(UIMessageOwner.ExcelOwner,
                                "End date must be the same or after the start date.",
                                "Invalid Date",
                                MessageBoxButtons.OK,
                                MessageBoxIcon.Warning)
          dtpEndDate.Value = dtpStartDate.Value
        End If
      End If
    Catch ex As Exception
      ErrorHandler.UnHandleError(ex)
    End Try
  End Sub


  ' ==========================================================================================
  ' Routine: dtpRangeStartDate_ValueChanged
  ' Purpose:
  '   Ensure RangeEndDate is at least RangeStartDate when RangeStartDate changes.
  ' ==========================================================================================
  Private Sub dtpRangeStartDate_ValueChanged(sender As Object, e As EventArgs) Handles dtpRangeStartDate.ValueChanged
    Try
      If dtpRangeStartDate.Checked AndAlso dtpRangeEndDate.Checked Then
        If dtpRangeEndDate.Value < dtpRangeStartDate.Value Then
          dtpRangeEndDate.Value = dtpRangeStartDate.Value
        End If
      End If
    Catch ex As Exception
      ErrorHandler.UnHandleError(ex)
    End Try
  End Sub


  ' ==========================================================================================
  ' Routine: dtpRangeEndDate_ValueChanged
  ' Purpose:
  '   Validate RangeEndDate (EndDate must be >= StartDate).
  ' ==========================================================================================
  Private Sub dtpRangeEndDate_ValueChanged(sender As Object, e As EventArgs) Handles dtpRangeEndDate.ValueChanged
    Try
      If dtpRangeStartDate.Checked AndAlso dtpRangeEndDate.Checked Then
        If dtpRangeEndDate.Value < dtpRangeStartDate.Value Then
          MessageBoxHelper.Show(UIMessageOwner.ExcelOwner,
                                "End date must be the same or after the start date.",
                                "Invalid Date",
                                MessageBoxButtons.OK,
                                MessageBoxIcon.Warning)
          dtpRangeEndDate.Value = dtpRangeStartDate.Value
        End If
      End If
    Catch ex As Exception
      ErrorHandler.UnHandleError(ex)
    End Try
  End Sub


  ' ==========================================================================================
  ' Routine: optDateRange_CheckedChanged
  ' Purpose:
  '   Switch pattern tab when DateRange option is selected.
  ' ==========================================================================================
  Private Sub optDateRange_CheckedChanged(sender As Object, e As EventArgs) Handles optDateRange.CheckedChanged
    If optDateRange.Checked Then
      tabPattern.SelectedIndex = 0
      UpdateRangeEndControls()
    End If
  End Sub


  ' ==========================================================================================
  ' Routine: optWeekly_CheckedChanged
  ' Purpose:
  '   Switch pattern tab when Weekly option is selected.
  ' ==========================================================================================
  Private Sub optWeekly_CheckedChanged(sender As Object, e As EventArgs) Handles optWeekly.CheckedChanged
    If optWeekly.Checked Then
      tabPattern.SelectedIndex = 1
      UpdateRangeEndControls()
    End If
  End Sub


  ' ==========================================================================================
  ' Routine: optMonthly_CheckedChanged
  ' Purpose:
  '   Switch pattern tab when Monthly option is selected.
  ' ==========================================================================================
  Private Sub optMonthly_CheckedChanged(sender As Object, e As EventArgs) Handles optMonthly.CheckedChanged
    If optMonthly.Checked Then
      tabPattern.SelectedIndex = 2
      UpdateRangeEndControls()
      UpdateMonthlyPatternControls()
    End If
  End Sub

  ' ==========================================================================================
  ' Routine: UpdateMonthlyPatternControls
  ' Purpose: Enables/disables monthly pattern controls based on selected type.
  ' ==========================================================================================
  Private Sub UpdateMonthlyPatternControls()
    '=== Disable all monthly pattern controls initially ===
    nudDayOfMonth.Enabled = False
    nudRecurMonthsDate.Enabled = False

    cmbOrdinalMonthsDay.Enabled = False
    cmbMonthlyDayOfWeek.Enabled = False
    nudRecurMonthsDay.Enabled = False

    cmbOrdinalMonthsWeek.Enabled = False
    nudRecurMonthsWeek.Enabled = False
    ' === Monthly Date pattern ===
    If optMonthlyDate.Checked Then
      nudDayOfMonth.Enabled = True
      nudRecurMonthsDate.Enabled = True
      ' === Monthly Day pattern ===
    ElseIf optMonthlyDay.Checked Then
      cmbOrdinalMonthsDay.Enabled = True
      cmbMonthlyDayOfWeek.Enabled = True
      nudRecurMonthsDay.Enabled = True
      ' === Monthly Week pattern ===
    ElseIf optMonthlyWeek.Checked Then
      cmbOrdinalMonthsWeek.Enabled = True
      nudRecurMonthsWeek.Enabled = True
    End If
  End Sub

  ' ==========================================================================================
  ' Routine: UpdateRangeEndControls
  ' Purpose: Enables/disables range end controls based on selected type.
  ' ==========================================================================================
  Private Sub UpdateRangeEndControls()
    '=== Disable all range end controls initially ===
    optRangeEndBy.Enabled = False
    optRangeEndAfter.Enabled = False
    optRangeNoEndDate.Enabled = False
    dtpRangeEndDate.Enabled = False
    nudRangeEndAfter.Enabled = False
    dtpRangeStartDate.Enabled = True ' always enabled
    If optDateRange.Checked Then
      ' === There are no range end dates for this so only enable this option ===
      optRangeNoEndDate.Enabled = True
      optRangeNoEndDate.Checked = True
    Else
      '=== Enable all range end options ===
      optRangeEndBy.Enabled = True
      optRangeEndAfter.Enabled = True
      optRangeNoEndDate.Enabled = True
      '==== Enable relevant control depending which option selected===
      If optRangeEndBy.Checked Then
        dtpRangeEndDate.Enabled = True
      ElseIf optRangeEndAfter.Checked Then
        nudRangeEndAfter.Enabled = True
      ElseIf optRangeNoEndDate.Checked Then
        '=== nothing to enable ===
      End If
    End If
  End Sub

  ' ==========================================================================================
  ' Routine: RangeEndControls_CheckedChanged
  ' Purpose: Handles range end control checked changes.
  ' ==========================================================================================
  Private Sub RangeEndControls_CheckedChanged(sender As Object, e As EventArgs) _
    Handles optRangeEndBy.CheckedChanged, optRangeEndAfter.CheckedChanged, optRangeNoEndDate.CheckedChanged
    UpdateRangeEndControls()
  End Sub

  ' ==========================================================================================
  ' Routine: MonthlyPatternControls_CheckedChanged
  ' Purpose: Handles monthly pattern control checked changes.
  ' ==========================================================================================
  Private Sub MonthlyPatternControls_CheckedChanged(sender As Object, e As EventArgs) _
    Handles optMonthlyDate.CheckedChanged, optMonthlyDay.CheckedChanged, optMonthlyWeek.CheckedChanged
    UpdateMonthlyPatternControls()
  End Sub

  Private Sub optMonthlyWeek_CheckedChanged(sender As Object, e As EventArgs)

  End Sub

  Private Sub optMonthlyDay_CheckedChanged(sender As Object, e As EventArgs)

  End Sub

  Private Sub optMonthlyDate_CheckedChanged(sender As Object, e As EventArgs)

  End Sub
End Class


