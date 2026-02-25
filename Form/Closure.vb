Imports System.ComponentModel
Imports System.Windows.Forms

Friend Class Closure
  Private _model As UIModelClosures
  ' === BindingSource for dgvClosures as we sort the grid so need have this level of indirection ===
  Private _bsClosures As New BindingSource()

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
  ' Routine: Closure_Load
  ' Purpose: Initialize form, load the UI model via UILoaderSaverClosure and configure the UI.
  ' Parameters:
  '   sender - event sender
  '   e      - event args
  ' Returns:
  '   None
  ' Notes:
  '   - Centers form on Excel, sets ToolTips, configures dgvClosures columns and binding.
  '   - Uses ErrorHandler.UnHandleError for error reporting.
  ' ==========================================================================================
  Private Sub Closure_Load(sender As Object, e As EventArgs) Handles MyBase.Load
    Try
      ' === Center form relative to Excel ===
      FormHelpers.CenterFormOnExcel(Me)
      ' === Load model ===
      _model = UILoaderSaverClosure.LoadClosuresModel()

      ' === Add ToolTip to controls ===
      ToolTip1.SetToolTip(btnAddNew, "Create a new closure")
      ToolTip1.SetToolTip(btnUpdate, "Update the selected closure")
      ToolTip1.SetToolTip(btnDelete, "Remove the selected closure")

      ' === Configure grid (one-time) ===
      dgvClosures.AutoGenerateColumns = False
      dgvClosures.Columns.Clear()
      dgvClosures.AllowUserToAddRows = False
      dgvClosures.AllowUserToDeleteRows = False
      dgvClosures.ReadOnly = True
      dgvClosures.RowHeadersVisible = False
      dgvClosures.SelectionMode = System.Windows.Forms.DataGridViewSelectionMode.FullRowSelect
      dgvClosures.MultiSelect = False
      dgvClosures.ColumnHeadersVisible = False ' === hide headers

      ' ===  Create columns and set Fill weights ===
      Dim colName As New DataGridViewTextBoxColumn() With {
        .Name = "colClosureName",
        .HeaderText = "ClosureName",
        .DataPropertyName = "ClosureName"
      }
      Dim colStart As New DataGridViewTextBoxColumn() With {
        .Name = "colStartDate",
        .HeaderText = "StartDate",
        .DataPropertyName = "StartDate",
        .DefaultCellStyle = New DataGridViewCellStyle() With {.Format = "d"}
      }
      Dim colEnd As New DataGridViewTextBoxColumn() With {
        .Name = "colEndDate",
        .HeaderText = "EndDate",
        .DataPropertyName = "EndDate",
        .DefaultCellStyle = New DataGridViewCellStyle() With {.Format = "d"}
      }

      ' === Add columns ===
      dgvClosures.Columns.AddRange(New DataGridViewColumn() {colName, colStart, colEnd})

      ' === Use Fill mode so columns resize to available client area (scrollbar area is considered) ===
      dgvClosures.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill

      ' === Allocate relative weights (adjust to taste) ===
      colName.FillWeight = 60   ' 60% of space
      colStart.FillWeight = 20  ' 20% of space
      colEnd.FillWeight = 20    ' 20% of space

      '' === Bind to your BindingList(Of UIClosureRow) ===
      'dgvClosures.DataSource = _model.Closures
      ' === Initialize closures list edit controls ===
      InitializeClosuresListControls()

      ' === Ensure layout recalculation when size or data changes ===
      AddHandler dgvClosures.DataBindingComplete, AddressOf dgvClosures_DataBindingComplete
      AddHandler dgvClosures.SizeChanged, AddressOf dgvClosures_SizeChanged

      ' === Set initial selection to none ===
      dgvClosures.ClearSelection()

    Catch ex As Exception
      ' === Log / show single user-facing message via ErrorHandler only. ===
      ErrorHandler.UnHandleError(ex)
      Return
    End Try
  End Sub

  ' ==========================================================================================
  ' Routine: dgvClosures_DataBindingComplete
  ' Purpose: Recalculate column layout after data binding completes.
  ' Parameters:
  '   sender - event sender
  '   e      - DataGridViewBindingCompleteEventArgs
  ' Returns:
  '   None
  ' Notes:
  '   - Forces Fill autosizing so columns consider current viewport width.
  ' ==========================================================================================
  Private Sub dgvClosures_DataBindingComplete(sender As Object, e As DataGridViewBindingCompleteEventArgs)
    Try
      ' Force recalculation of Fill layout
      dgvClosures.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill
      dgvClosures.Refresh()
    Catch ex As Exception
      ErrorHandler.UnHandleError(ex)
    End Try
  End Sub

  ' ==========================================================================================
  ' Routine: dgvClosures_SizeChanged
  ' Purpose: Reapply column sizing policy when the control is resized.
  ' Parameters:
  '   sender - event sender
  '   e      - event args
  ' Returns:
  '   None
  ' Notes:
  '   - Keeps columns sized to client area so horizontal scrollbar does not appear when vertical scrollbar appears.
  ' ==========================================================================================
  Private Sub dgvClosures_SizeChanged(sender As Object, e As EventArgs)
    Try
      ' Reapply Fill mode so columns account for current client size (and scrollbar)
      dgvClosures.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill
    Catch ex As Exception
      ErrorHandler.UnHandleError(ex)
    End Try
  End Sub

  ' ==========================================================================================
  ' Routine: dgvClosures_SelectionChanged
  ' Purpose: Populate edit controls when a row is selected.
  ' Parameters:
  '   sender - event sender
  '   e      - event args
  ' Returns:
  '   None
  ' Notes:
  '   - Reads selected UIClosureRow from CurrentRow.DataBoundItem.
  ' ==========================================================================================
  Private Sub dgvClosures_SelectionChanged(sender As Object, e As EventArgs) Handles dgvClosures.SelectionChanged
    Try
      Dim selected As UIClosureRow = Nothing
      If dgvClosures.CurrentRow IsNot Nothing Then
        selected = TryCast(dgvClosures.CurrentRow.DataBoundItem, UIClosureRow)
      End If

      If selected Is Nothing Then
        txtClosureName.Text = String.Empty
        dtpStartDate.Value = DateTime.Today
        dtpEndDate.Value = DateTime.Today
        Return
      End If

      txtClosureName.Text = If(String.IsNullOrWhiteSpace(selected.ClosureName), String.Empty, selected.ClosureName)

      ' === Start Date ===
      If selected.StartDate = Date.MinValue Then
        dtpStartDate.Checked = False
      Else
        dtpStartDate.Checked = True
        dtpStartDate.Value = selected.StartDate
      End If

      ' === End Date ===
      If selected.EndDate = Date.MinValue Then
        dtpEndDate.Checked = False
      Else
        dtpEndDate.Checked = True
        dtpEndDate.Value = selected.EndDate
      End If
      'If selected.StartDate = Date.MinValue Then
      '  dtpStartDate.Value = DateTime.Today
      'Else
      '  dtpStartDate.Value = selected.StartDate
      'End If

      'If selected.EndDate = Date.MinValue Then
      '  dtpEndDate.Value = dtpStartDate.Value
      'Else
      '  dtpEndDate.Value = selected.EndDate
      'End If

      If dtpStartDate.Checked AndAlso dtpEndDate.Checked Then
        If dtpEndDate.Value < dtpStartDate.Value Then
          dtpEndDate.Value = dtpStartDate.Value
        End If
      End If

    Catch ex As Exception
      ErrorHandler.UnHandleError(ex)
      Return
    End Try
  End Sub

  ' ==========================================================================================
  ' Routine: dtpStartDate_ValueChanged
  ' Purpose: Ensure EndDate is at least StartDate when StartDate changes.
  ' Parameters:
  '   sender - event sender
  '   e      - event args
  ' Returns:
  '   None
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
      Return
    End Try
  End Sub

  ' ==========================================================================================
  ' Routine: dtpEndDate_ValueChanged
  ' Purpose: Validate EndDate (EndDate must be >= StartDate).
  ' Parameters:
  '   sender - event sender
  '   e      - event args
  ' Returns:
  '   None
  ' Notes:
  '   - Uses MessageBoxHelper to show a user-friendly message when validation fails.
  ' ==========================================================================================
  Private Sub dtpEndDate_ValueChanged(sender As Object, e As EventArgs) Handles dtpEndDate.ValueChanged
    Try
      If dtpStartDate.Checked AndAlso dtpEndDate.Checked Then
        If dtpEndDate.Value < dtpStartDate.Value Then
          MessageBoxHelper.Show(UIMessageOwner.ExcelOwner, "End date must be the same or after the start date.",
                          "Invalid Date",
                          MessageBoxButtons.OK,
                          MessageBoxIcon.Warning)
          dtpEndDate.Value = dtpStartDate.Value
        End If
      End If
    Catch ex As Exception
      ErrorHandler.UnHandleError(ex)
      Return
    End Try
  End Sub

  ' ==========================================================================================
  ' Routine: btnAddNew_Click
  ' Purpose: Create a new closure from UI values and persist via UILoaderSaverClosure.
  ' Parameters:
  '   sender - event sender
  '   e      - event args
  ' Returns:
  '   None
  ' Notes:
  '   - Validates dates, builds payload on model and calls SavePendingClosureAction.
  ' ==========================================================================================
  Private Sub btnAddNew_Click(sender As Object, e As EventArgs) Handles btnAddNew.Click
    Try
      If Not ValidateDates() Then Exit Sub

      ' === Save list of current closure ID's to find the new one later ===
      Dim preservedIds As New HashSet(Of String)(
    _model.Closures.Select(Function(c) c.ClosureID))

      ' === Direct assignment to canonical model property
      _model.ActionClosure.ClosureName = txtClosureName.Text.Trim()
      _model.ActionClosure.StartDate = dtpStartDate.Value.Date
      _model.ActionClosure.EndDate = dtpEndDate.Value.Date
      _model.PendingAction = ClosureAction.Add

      UILoaderSaverClosure.SavePendingClosureAction(_model)

      ' === Find newly added closure ID ===
      Dim newId As String = Nothing
      For Each row In _model.Closures
        If Not preservedIds.Contains(row.ClosureID) Then
          newId = row.ClosureID
          Exit For
        End If
      Next

      InitializeClosuresListControls(newId)
      '_model = UILoaderSaverClosure.LoadClosuresModel()
      'dgvClosures.DataSource = _model.Closures

    Catch ex As Exception
      ErrorHandler.UnHandleError(ex)
      Return
    End Try
  End Sub

  ' ==========================================================================================
  ' Routine: btnUpdate_Click
  ' Purpose: Update the selected closure with UI values and persist via UILoaderSaverClosure.
  ' Parameters:
  '   sender - event sender
  '   e      - event args
  ' Returns:
  '   None
  ' Notes:
  '   - Requires a selected row; uses Selected ClosureID for the update payload.
  ' ==========================================================================================
  Private Sub btnUpdate_Click(sender As Object, e As EventArgs) Handles btnUpdate.Click
    Try
      Dim selected As UIClosureRow = Nothing

      If dgvClosures.CurrentRow IsNot Nothing Then
        selected = TryCast(dgvClosures.CurrentRow.DataBoundItem, UIClosureRow)
      End If

      If selected Is Nothing Then
        MessageBoxHelper.Show(UIMessageOwner.ExcelOwner, "Please select a closure to update.", "No Selection", MessageBoxButtons.OK, MessageBoxIcon.Information)
        Exit Sub
      End If

      If Not ValidateDates() Then Exit Sub

      ' Direct assignment to canonical model property
      _model.ActionClosure.ClosureID = selected.ClosureID
      _model.ActionClosure.ClosureName = txtClosureName.Text.Trim()
      _model.ActionClosure.StartDate = dtpStartDate.Value.Date
      _model.ActionClosure.EndDate = dtpEndDate.Value.Date
      _model.PendingAction = ClosureAction.Update

      UILoaderSaverClosure.SavePendingClosureAction(_model)

      InitializeClosuresListControls(selected.ClosureID)
      '_model = UILoaderSaverClosure.LoadClosuresModel()
      'dgvClosures.DataSource = _model.Closures

    Catch ex As Exception
      ErrorHandler.UnHandleError(ex)
      Return
    End Try
  End Sub

  ' ==========================================================================================
  ' Routine: btnDelete_Click
  ' Purpose: Delete the selected closure (marks as deleted) and persist via UILoaderSaverClosure.
  ' Parameters:
  '   sender - event sender
  '   e      - event args
  ' Returns:
  '   None
  ' Notes:
  '   - Confirms with the user via MessageBoxHelper before deleting.
  ' ==========================================================================================
  Private Sub btnDelete_Click(sender As Object, e As EventArgs) Handles btnDelete.Click
    Try
      Dim selected As UIClosureRow = Nothing
      If dgvClosures.CurrentRow IsNot Nothing Then
        selected = TryCast(dgvClosures.CurrentRow.DataBoundItem, UIClosureRow)
      End If

      If selected Is Nothing Then
        MessageBoxHelper.Show(UIMessageOwner.ExcelOwner, "Please select a closure to delete.", "No Selection", MessageBoxButtons.OK, MessageBoxIcon.Information)
        Exit Sub
      End If

      If MessageBoxHelper.Show(UIMessageOwner.ExcelOwner, "Are you sure you want to delete the selected closure?", "Confirm Delete", MessageBoxButtons.YesNo, MessageBoxIcon.Question) <> DialogResult.Yes Then
        Exit Sub
      End If

      ' Direct assignment to canonical model property
      _model.ActionClosure.ClosureID = selected.ClosureID
      _model.PendingAction = ClosureAction.Delete

      UILoaderSaverClosure.SavePendingClosureAction(_model)

      InitializeClosuresListControls()
      '_model = UILoaderSaverClosure.LoadClosuresModel()
      'dgvClosures.DataSource = _model.Closures

    Catch ex As Exception
      ErrorHandler.UnHandleError(ex)
      Return
    End Try
  End Sub

  ' ==========================================================================================
  ' Routine: btnClose_Click
  ' Purpose: Close the dialog.
  ' Parameters:
  '   sender - event sender
  '   e      - event args
  ' Returns:
  '   None
  ' ==========================================================================================
  Private Sub btnClose_Click(sender As Object, e As EventArgs) Handles btnClose.Click
    Me.Close()
  End Sub

  ' ==========================================================================================
  ' Routine: ValidateDates
  ' Purpose: Ensure EndDate is the same or after StartDate.
  ' Parameters:
  '   None
  ' Returns:
  '   Boolean - True if valid, otherwise False
  ' Notes:
  '   - Shows a user-friendly message via MessageBoxHelper when invalid.
  ' ==========================================================================================
  Private Function ValidateDates() As Boolean
    If dtpEndDate.Value < dtpStartDate.Value Then
      MessageBoxHelper.Show(UIMessageOwner.ExcelOwner, "End date must be the same or after the start date.",
                      "Invalid Date Range",
                      MessageBoxButtons.OK,
                      MessageBoxIcon.Warning)
      Return False
    End If
    Return True
  End Function

  ' ==========================================================================================
  ' Routine: InitializeClosuresListControls
  ' Purpose: Initialize data view grid closures associated controls.
  ' Parameters:
  '   None
  ' Returns:
  '   None
  ' ==========================================================================================
  Private Sub InitializeClosuresListControls(Optional preserveClosureId As String = "")
    Try

      ' === Bind the model list through a BindingSource so the grid gets proper currency, notifications, and sorting support  ===
      _bsClosures.DataSource = _model.Closures
      ' ===  Bind the grid to the BindingSource (never directly to the list) so row selection and sorting behave correctly ===
      dgvClosures.DataSource = _bsClosures
      ' === Invalidate/ Refresh() to ensure visual update for any non-notifying changes. ===
      dgvClosures.Invalidate()
      dgvClosures.Refresh()
      dgvClosures.ClearSelection()
      dgvClosures.CurrentCell = Nothing
      '' === Apply default sort (StartDate descending) ===
      dgvClosures.Columns("colStartDate").SortMode = DataGridViewColumnSortMode.Automatic
      dgvClosures.Sort(dgvClosures.Columns("colStartDate"), ListSortDirection.Descending)
      ' === Clear edit controls ===
      txtClosureName.Text = String.Empty
      dtpStartDate.Value = DateTime.Today
      dtpEndDate.Value = DateTime.Today
      ' === Re-select previous item if it exists ===
      If Not String.IsNullOrWhiteSpace(preserveClosureId) Then
        For i As Integer = 0 To dgvClosures.Rows.Count - 1
          Dim row = TryCast(dgvClosures.Rows(i).DataBoundItem, UIClosureRow)
          If row IsNot Nothing AndAlso row.ClosureID = preserveClosureId Then
            dgvClosures.Rows(i).Selected = True
            Exit For
          End If
        Next
      End If

      'If TypeOf dgvClosures.DataSource Is BindingSource Then
      '  DirectCast(dgvClosures.DataSource, BindingSource).ResetBindings(False)
      'Else
      '  ' If bound Then directly To a BindingList(Of T), it usually notifies automatically.
      '  'Invalidate/ Refresh() to ensure visual update for any non-notifying changes.
      '  dgvClosures.Invalidate()
      '  dgvClosures.Refresh()
      'End If
    Catch ex As Exception
      ErrorHandler.UnHandleError(ex)
    End Try
  End Sub
End Class