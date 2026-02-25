Imports System.Drawing
Imports System.Windows.Forms

' ==========================================================================================
' Routine: Resource (Form Class)
' Purpose:
'   UI dialog for managing a single Resource.
'   Loads UIModelResources, binds fields, validates input, and persists changes.
' Parameters:
'   None (properties ResourceID and WasSaved are set externally)
' Returns:
'   Form instance used by caller
' Notes:
'   - ResourceID = "" means Add mode
'   - ResourceID <> "" means Update/Delete mode
'   - EndDate must be >= StartDate (given current non-nullable DateTimePicker)
' ==========================================================================================
Friend Class Resource

  Friend Property WasSaved As Boolean = False
  Friend Property ResourceID As String = ""

  Private _model As UIModelResources

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
  ' Routine: Resource_Load
  ' Purpose:
  '   Initialize dialog on load.
  '   - Center form relative to Excel
  '   - Load UI model
  '   - Bind lookups, fields, and name/value grid
  ' Parameters:
  '   sender - event sender
  '   e      - event args
  ' Returns:
  '   None
  ' Notes:
  '   Uses UILoaderSaverResource.LoadResourceModel(ResourceID)
  ' ==========================================================================================
  Private Sub Resource_Load(sender As Object, e As EventArgs) Handles MyBase.Load
    Try

      FormHelpers.CenterFormOnExcel(Me)

      _model = LoadResourceModel(ResourceID)

      BindLookups()
      BindResourceFields()
      BindResourceNameValueGrid()
      ConfigureValueCellsPerRow()
      lstListItems.View = View.List
      lstListItems.CheckBoxes = True
      BindResourceListItemNames()

    Catch ex As Exception
      ErrorHandler.UnHandleError(ex)
    End Try
  End Sub

  ' ==========================================================================================
  ' Routine: BindLookups
  ' Purpose:
  '   Bind lookup lists (Salutation, Gender)
  '   to ComboBoxes using UIListItemRow.
  ' Parameters:
  '   None
  ' Returns:
  '   None
  ' Notes:
  '   SelectedValue will be ListItemID (GUID); DisplayMember is ListItemName.
  ' ==========================================================================================
  Private Sub BindLookups()

    ' === Salutation ===
    cmbSalutation.DataSource = _model.Salutations
    cmbSalutation.DisplayMember = NameOf(UIListItemRow.ListItemName)
    cmbSalutation.ValueMember = NameOf(UIListItemRow.ListItemID)
    lblSalutation.Text = _model.SalutationListItemTypeName & ":"

    ' === Gender ===
    cmbGender.DataSource = _model.Genders
    cmbGender.DisplayMember = NameOf(UIListItemRow.ListItemName)
    cmbGender.ValueMember = NameOf(UIListItemRow.ListItemID)
    lblGender.Text = _model.GenderListItemTypeName & ":"

  End Sub

  ' ==========================================================================================
  ' Routine: BindResourceFields
  ' Purpose:
  '   Populate UI controls from _model.ActionResource.
  ' Parameters:
  '   None
  ' Returns:
  '   None
  ' Notes:
  '   Uses ID fields for ComboBoxes (Salutation, Gender).
  ' ==========================================================================================
  Private Sub BindResourceFields()

    Dim a = _model.ActionResource

    txtPreferredName.Text = a.PreferredName
    txtFirstName.Text = a.FirstName
    txtLastName.Text = a.LastName

    ' === Salutation ===
    If Not String.IsNullOrEmpty(a.SalutationID) Then
      cmbSalutation.SelectedValue = a.SalutationID
    Else
      cmbSalutation.SelectedIndex = -1
    End If

    ' === Gender ===
    If Not String.IsNullOrEmpty(a.GenderID) Then
      cmbGender.SelectedValue = a.GenderID
    Else
      cmbGender.SelectedIndex = -1
    End If

    txtEmail.Text = a.Email
    txtPhone.Text = a.Phone

    ' === Start Date ===
    If a.StartDate = Date.MinValue Then
      dtpStartDate.Checked = False
    Else
      dtpStartDate.Checked = True
      dtpStartDate.Value = a.StartDate
    End If

    ' === End Date ===
    If a.EndDate = Date.MinValue Then
      dtpEndDate.Checked = False
    Else
      dtpEndDate.Checked = True
      dtpEndDate.Value = a.EndDate
    End If

    If dtpStartDate.Checked AndAlso dtpEndDate.Checked Then
      If dtpEndDate.Value < dtpStartDate.Value Then
        dtpEndDate.Value = dtpStartDate.Value
      End If
    End If

    txtNotes.Text = a.Notes

  End Sub

  ' ==========================================================================================
  ' Routine: BindResourceNameValueGrid
  ' Purpose:
  '   Bind the name/value pairs grid (dgvResourceNameValue)
  ' Parameters:
  '   None
  ' Returns:
  '   None
  ' Notes:
  '   Value column switches between TextBox and ComboBox in CellFormatting.
  ' ==========================================================================================
  Private Sub BindResourceNameValueGrid()

    dgvResourceNameValue.AutoGenerateColumns = False
    dgvResourceNameValue.Columns.Clear()

    Dim colName As New DataGridViewTextBoxColumn()
    colName.HeaderText = "Name"
    colName.DataPropertyName = NameOf(UIResourceNameValueRow.ResourceListItemName)
    colName.ReadOnly = True
    colName.FillWeight = 50   ' 50% of space
    dgvResourceNameValue.Columns.Add(colName)

    Dim colValue As New DataGridViewTextBoxColumn()
    colValue.HeaderText = "Value"
    ' IMPORTANT: do NOT bind this column to ResourceListItemValue
    colValue.DataPropertyName = Nothing
    colValue.Name = "ValueColumn"
    colValue.FillWeight = 50   ' 50% of space
    dgvResourceNameValue.Columns.Add(colValue)

    ' === Ensure layout recalculation when size or data changes ===
    AddHandler dgvResourceNameValue.DataBindingComplete, AddressOf dgvResourceNameValue_DataBindingComplete
    AddHandler dgvResourceNameValue.SizeChanged, AddressOf dgvResourceNameValue_SizeChanged

    '=== Bind data to grid except MultiSelectList ===
    Dim rows = _model.ActionResourceNameValues.
    Where(Function(r) EnumEntries.GetCardinality(r.ValueType) = FieldCardinality.One).
    ToList()
    'Dim rows = _model.ActionResourceNameValues.
    'Where(Function(r) r.ValueType <> ResourceListItemValueType.MultiSelectList).
    'ToList()
    Dim bindingList As New SortableBindingList(Of UIResourceNameValueRow)
    For Each row In rows
      bindingList.Add(row)
    Next
    dgvResourceNameValue.DataSource = bindingList

  End Sub

  ' ==========================================================================================
  ' Routine: BindResourceListItemNames
  ' Purpose:
  '   Populate cmbResourceListItemNames with all ResourceListItem definitions.
  ' ==========================================================================================
  Private Sub BindResourceListItemNames()
    Dim rows = _model.ActionResourceNameValues.
    Where(Function(r) EnumEntries.GetCardinality(r.ValueType) = FieldCardinality.Many).
    ToList()
    'Dim rows = _model.ActionResourceNameValues.
    'Where(Function(r) r.ValueType = ResourceListItemValueType.MultiSelectList).
    'ToList()
    Dim bindingList As New SortableBindingList(Of UIResourceNameValueRow)
    For Each row In rows
      bindingList.Add(row)
    Next
    cmbResourceListItemNames.DataSource = bindingList
    cmbResourceListItemNames.DisplayMember = NameOf(UIResourceNameValueRow.ResourceListItemName)
    cmbResourceListItemNames.ValueMember = NameOf(UIResourceNameValueRow.ResourceListItemID)
  End Sub

  ' ==========================================================================================
  ' Routine: dgvResourceNameValue_EditingControlShowing
  ' Purpose:
  '   Attach to the ComboBox editing control used for lookup rows in the ValueColumn.
  '   Ensures that when the user selects a new dropdown value, the corresponding
  '   UIResourceNameValueRow is updated immediately, even if the user does not exit the cell.
  ' Parameters:
  '   sender - the DataGridView raising the event
  '   e      - provides the editing control instance
  ' Returns:
  '   None
  ' Notes:
  '   - Only attaches when the current cell is in the ValueColumn and is a ComboBox.
  '   - Prevents stale ListItemID values when sorting or repainting occurs.
  ' ==========================================================================================
  Private Sub dgvResourceNameValue_EditingControlShowing(sender As Object,
                                                       e As DataGridViewEditingControlShowingEventArgs) _
                                                       Handles dgvResourceNameValue.EditingControlShowing
    Try
      Dim grid = dgvResourceNameValue

      If grid.CurrentCell Is Nothing Then Exit Sub
      If grid.Columns(grid.CurrentCell.ColumnIndex).Name <> "ValueColumn" Then Exit Sub

      Dim combo = TryCast(e.Control, ComboBox)
      If combo Is Nothing Then Exit Sub

      ' Avoid multiple subscriptions
      RemoveHandler combo.SelectedIndexChanged, AddressOf ValueCombo_SelectedIndexChanged
      AddHandler combo.SelectedIndexChanged, AddressOf ValueCombo_SelectedIndexChanged

    Catch ex As Exception
      ErrorHandler.UnHandleError(ex)
    End Try
  End Sub

  ' ==========================================================================================
  ' Routine: ValueCombo_SelectedIndexChanged
  ' Purpose:
  '   Update the UIResourceNameValueRow for lookup rows when the user selects a new value
  '   in the ComboBox editing control. Ensures the model reflects the user's choice
  '   immediately, without requiring the user to exit the cell.
  ' Parameters:
  '   sender - the ComboBox editing control
  '   e      - event arguments
  ' Returns:
  '   None
  ' Notes:
  '   - Applies only to lookup rows (ListItemTypeID not empty).
  '   - Ensures sorting and repainting display the updated value correctly.
  ' ==========================================================================================
  Private Sub ValueCombo_SelectedIndexChanged(sender As Object, e As EventArgs)
    Try
      Dim combo = TryCast(sender, ComboBox)
      If combo Is Nothing Then Exit Sub

      Dim grid = dgvResourceNameValue
      Dim row = grid.CurrentRow
      If row Is Nothing Then Exit Sub

      Dim uiRow = TryCast(row.DataBoundItem, UIResourceNameValueRow)
      If uiRow Is Nothing Then Exit Sub

      ' Lookup rows only
      If Not String.IsNullOrEmpty(uiRow.ListItemTypeID) Then
        uiRow.SelectedListItemID = TryCast(combo.SelectedValue, String)
        uiRow.ResourceListItemValue = Nothing
      End If

    Catch ex As Exception
      ErrorHandler.UnHandleError(ex)
    End Try
  End Sub

  ' ==========================================================================================
  ' Routine: dgvResourceNameValue_CellValueChanged
  ' Purpose:
  '   Update UIResourceNameValueRow for free-text rows when the user edits the ValueColumn.
  '   Lookup rows are handled by the ComboBox editing control and are excluded here.
  ' Parameters:
  '   sender - the DataGridView raising the event
  '   e      - identifies the changed cell
  ' Returns:
  '   None
  ' Notes:
  '   - Required for free-text rows because they do not raise SelectedIndexChanged.
  ' ==========================================================================================
  Private Sub dgvResourceNameValue_CellValueChanged(sender As Object,
                                                  e As DataGridViewCellEventArgs) _
                                                  Handles dgvResourceNameValue.CellValueChanged
    Try
      If e.RowIndex < 0 Then Exit Sub
      If dgvResourceNameValue.Columns(e.ColumnIndex).Name <> "ValueColumn" Then Exit Sub

      Dim row = dgvResourceNameValue.Rows(e.RowIndex)
      Dim uiRow = TryCast(row.DataBoundItem, UIResourceNameValueRow)
      If uiRow Is Nothing Then Exit Sub

      ' Free-text rows only
      If String.IsNullOrEmpty(uiRow.ListItemTypeID) Then
        uiRow.ResourceListItemValue = TryCast(row.Cells(e.ColumnIndex).Value, String)
        uiRow.SelectedListItemID = Nothing
      End If

    Catch ex As Exception
      ErrorHandler.UnHandleError(ex)
    End Try
  End Sub

  ' ==========================================================================================
  ' Routine: dgvResourceNameValue_Sorted
  ' Purpose:
  '   Reapply ComboBox/TextBox cells after the grid has been sorted, because sorting causes
  '   the DataGridView to recreate all rows and revert cells to their column templates.
  ' ==========================================================================================
  Private Sub dgvResourceNameValue_Sorted(sender As Object, e As EventArgs) _
    Handles dgvResourceNameValue.Sorted

    ConfigureValueCellsPerRow()

  End Sub

  ' ==========================================================================================
  ' Routine: dgvResourceNameValue_DataBindingComplete
  ' Purpose: Recalculate column layout after data binding completes.
  ' Parameters:
  '   sender - event sender
  '   e      - DataGridViewBindingCompleteEventArgs
  ' Returns:
  '   None
  ' Notes:
  '   - Forces Fill autosizing so columns consider current viewport width.
  ' ==========================================================================================
  Private Sub dgvResourceNameValue_DataBindingComplete(sender As Object, e As DataGridViewBindingCompleteEventArgs)
    Try
      ' Force recalculation of Fill layout
      dgvResourceNameValue.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill
      dgvResourceNameValue.Refresh()
    Catch ex As Exception
      ErrorHandler.UnHandleError(ex)
    End Try
  End Sub

  ' ==========================================================================================
  ' Routine: dgvResourceNameValue_SizeChanged
  ' Purpose: Reapply column sizing policy when the control is resized.
  ' Parameters:
  '   sender - event sender
  '   e      - event args
  ' Returns:
  '   None
  ' Notes:
  '   - Keeps columns sized to client area so horizontal scrollbar does not appear when vertical scrollbar appears.
  ' ==========================================================================================
  Private Sub dgvResourceNameValue_SizeChanged(sender As Object, e As EventArgs)
    Try
      ' Reapply Fill mode so columns account for current client size (and scrollbar)
      dgvResourceNameValue.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill
    Catch ex As Exception
      ErrorHandler.UnHandleError(ex)
    End Try
  End Sub

  ' ==========================================================================================
  ' Routine: ConfigureValueCellsPerRow
  ' Purpose: Assign the correct cell type (ComboBox or TextBox) to the ValueColumn for each row.
  ' Parameters:
  '   None
  ' Returns:
  '   None
  ' Notes:
  '   Must be called after binding; performs all per-row cell replacement outside paint/format.
  ' ==========================================================================================
  Private Sub ConfigureValueCellsPerRow()

    For Each gridRow As DataGridViewRow In dgvResourceNameValue.Rows

      If gridRow.IsNewRow Then Continue For

      Dim uiRow = TryCast(gridRow.DataBoundItem, UIResourceNameValueRow)
      If uiRow Is Nothing Then Continue For

      Dim cell = gridRow.Cells("ValueColumn")

      If Not String.IsNullOrEmpty(uiRow.ListItemTypeID) Then
        ' Lookup row → ComboBox cell
        Dim combo As New DataGridViewComboBoxCell()
        combo.DisplayMember = NameOf(UIListItemRow.ListItemName)
        combo.ValueMember = NameOf(UIListItemRow.ListItemID)
        combo.DataSource = uiRow.ListItems
        combo.Value = uiRow.SelectedListItemID
        gridRow.Cells("ValueColumn") = combo

      Else
        ' Free-text row → TextBox cell
        Dim txt As New DataGridViewTextBoxCell()
        txt.Value = uiRow.ResourceListItemValue
        gridRow.Cells("ValueColumn") = txt

      End If

    Next

  End Sub

  ' ==========================================================================================
  ' Routine: PushGridToModel
  ' Purpose:
  '   Push edited name/value rows from dgvResourceNameValue into the working model and
  '   set per-row PendingAction (Add/Update/Delete/None) based on snapshot vs working values.
  ' Parameters:
  '   None
  ' Returns:
  '   None
  ' Notes:
  '   - Snapshot: _model.ResourceNameValues
  '   - Working:  _model.ActionResourceNameValues (bound to dgvResourceNameValue)
  '   - Rules:
  '       * No snapshot + empty new value  -> None
  '       * No snapshot + non-empty value  -> Add
  '       * Snapshot value -> empty value  -> Delete
  '       * Snapshot value -> different value -> Update
  '       * Snapshot value -> same value   -> None
  ' ==========================================================================================
  Private Sub PushGridToModel()

    Try
      ' --- Ensure any in-progress edits are committed ---
      dgvResourceNameValue.EndEdit()
      dgvResourceNameValue.CommitEdit(DataGridViewDataErrorContexts.Commit)

      For Each uiRow As UIResourceNameValueRow In _model.ActionResourceNameValues

        Dim gridRow As DataGridViewRow =
          dgvResourceNameValue.Rows.Cast(Of DataGridViewRow)().
            FirstOrDefault(Function(r) r.DataBoundItem Is uiRow)

        If gridRow Is Nothing Then Continue For

        Dim cell = gridRow.Cells("ValueColumn")

        ' --- Find snapshot row for comparison (by ResourceListItemID) ---
        Dim snapRow As UIResourceNameValueRow =
          _model.ResourceNameValues.
            FirstOrDefault(Function(r) r.ResourceListItemID = uiRow.ResourceListItemID)

        Dim snapListItemID As String = Nothing
        Dim snapText As String = Nothing

        If snapRow IsNot Nothing Then
          snapListItemID = snapRow.SelectedListItemID
          snapText = snapRow.ResourceListItemValue
        End If

        Dim newListItemID As String = Nothing
        Dim newText As String = Nothing

        ' --- Read new values from grid cell into working row ---
        If TypeOf cell Is DataGridViewComboBoxCell Then
          ' Lookup row
          newListItemID = TryCast(cell.Value, String)
          newText = Nothing

          uiRow.SelectedListItemID = newListItemID
          uiRow.ResourceListItemValue = Nothing

        ElseIf TypeOf cell Is DataGridViewTextBoxCell Then
          ' Free-text row
          newListItemID = Nothing
          newText = TryCast(cell.Value, String)

          uiRow.SelectedListItemID = Nothing
          uiRow.ResourceListItemValue = newText

        End If

        ' --- Decide PendingAction based on snapshot vs new values ---
        Dim hadSnapshotValue As Boolean =
          (Not String.IsNullOrEmpty(snapListItemID)) OrElse
          (Not String.IsNullOrEmpty(snapText))

        Dim hasNewValue As Boolean =
          (Not String.IsNullOrEmpty(newListItemID)) OrElse
          (Not String.IsNullOrEmpty(newText))

        uiRow.PendingAction = ResourceNameValueAction.None

        If Not hadSnapshotValue AndAlso Not hasNewValue Then
          ' No value before, no value now -> None

        ElseIf Not hadSnapshotValue AndAlso hasNewValue Then
          ' No previous value, now has a value -> Add
          uiRow.PendingAction = ResourceNameValueAction.Add

        ElseIf hadSnapshotValue AndAlso Not hasNewValue Then
          ' Had a value, now cleared -> Delete
          uiRow.PendingAction = ResourceNameValueAction.Delete

        Else
          ' Had a value, still has a value -> compare
          Dim snapshotKey As String = If(snapListItemID, snapText)
          Dim newKey As String = If(newListItemID, newText)

          If Not String.Equals(snapshotKey, newKey, StringComparison.Ordinal) Then
            uiRow.PendingAction = ResourceNameValueAction.Update
          Else
            uiRow.PendingAction = ResourceNameValueAction.None
          End If

        End If

      Next

    Catch ex As Exception
      ErrorHandler.UnHandleError(ex)
    End Try

  End Sub

  ' ==========================================================================================
  ' Routine: dtpStartDate_ValueChanged
  ' Purpose:
  '   Ensure EndDate is at least StartDate when StartDate changes.
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
  ' Purpose:
  '   Validate EndDate (EndDate must be >= StartDate).
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
      Return
    End Try
  End Sub

  ' ==========================================================================================
  ' Routine: cmbResourceListItemNames_SelectedIndexChanged
  ' Purpose:
  '   For the selected ListItem call LoadListItemsEditor to populate lstListItems.
  ' Parameters:
  '   sender - event sender
  '   e      - event args
  ' Returns:
  '   None
  ' Notes:
  ' ==========================================================================================

  Private Sub cmbResourceListItemNames_SelectedIndexChanged(sender As Object, e As EventArgs) _
    Handles cmbResourceListItemNames.SelectedIndexChanged
    Try
      Dim uiRow = TryCast(cmbResourceListItemNames.SelectedItem, UIResourceNameValueRow)
      If uiRow Is Nothing Then Exit Sub
      LoadListItemsEditor(uiRow)
    Catch ex As Exception
      ErrorHandler.UnHandleError(ex)
    End Try
  End Sub

  ' ==========================================================================================
  ' Routine: LoadListItemsEditor
  ' Purpose:
  '   Populate lstListItems based on the selected ResourceListItem definition.
  '   Supports MultiSelectList only.
  ' Parameters:
  '   uiRow - the selected UIResourceNameValueRow
  ' Returns:
  '   None
  ' Notes:
  ' ==========================================================================================
  Private Sub LoadListItemsEditor(uiRow As UIResourceNameValueRow)
    lstListItems.Items.Clear()
    'If uiRow.ValueType <> ResourceListItemValueType.MultiSelectList Then
    If EnumEntries.GetCardinality(uiRow.ValueType) = FieldCardinality.One Then
      lstListItems.Enabled = False
      btnSaveListItems.Enabled = False
      btnCancelListItems.Enabled = False
      Return
    End If
    lstListItems.Enabled = True
    btnSaveListItems.Enabled = True
    btnCancelListItems.Enabled = True
    ' Populate list with checkboxes
    For Each li In uiRow.ListItems
      Dim item As New ListViewItem(li.ListItemName)
      item.Tag = li.ListItemID
      item.Checked = uiRow.SelectedListItemIDs.Contains(li.ListItemID)
      lstListItems.Items.Add(item)
    Next
  End Sub

  ' ==========================================================================================
  ' Routine: btnSaveListItems_Click
  ' Purpose:
  '   Push checked items from lstListItems into the selected UIResourceNameValueRow.
  ' Parameters:
  '   sender - event sender
  '   e      - event args
  ' Returns:
  '   None
  ' Notes:
  ' ==========================================================================================
  Private Sub btnSaveListItems_Click(sender As Object, e As EventArgs) _
    Handles btnSaveListItems.Click
    Try
      Dim uiRow = TryCast(cmbResourceListItemNames.SelectedItem, UIResourceNameValueRow)
      If uiRow Is Nothing Then Exit Sub
      'If uiRow.ValueType <> ResourceListItemValueType.MultiSelectList Then Exit Sub
      If EnumEntries.GetCardinality(uiRow.ValueType) <> FieldCardinality.Many Then Exit Sub
      Dim selected As New List(Of String)
      For Each item As ListViewItem In lstListItems.Items
        If item.Checked Then
          selected.Add(CStr(item.Tag))
        End If
      Next
      uiRow.SelectedListItemIDs = selected
      ' Mark as Update/Add/Delete later in PushGridToModel
      uiRow.PendingAction = ResourceNameValueAction.Update
    Catch ex As Exception
      ErrorHandler.UnHandleError(ex)
    End Try
  End Sub

  ' ==========================================================================================
  ' Routine: btnCancelListItems_Click
  ' Purpose:
  '   Revert lstListItems to the values stored in the UI model.
  ' Parameters:
  '   sender - event sender
  '   e      - event args
  ' Returns:
  '   None
  ' Notes:
  ' ==========================================================================================
  Private Sub btnCancelListItems_Click(sender As Object, e As EventArgs) _
    Handles btnCancelListItems.Click
    Try
      Dim uiRow = TryCast(cmbResourceListItemNames.SelectedItem, UIResourceNameValueRow)
      If uiRow Is Nothing Then Exit Sub
      LoadListItemsEditor(uiRow)
    Catch ex As Exception
      ErrorHandler.UnHandleError(ex)
    End Try
  End Sub

  ' ==========================================================================================
  ' Routine: btnSave_Click
  ' Purpose:
  '   Commit changes:
  '     - Push UI values into ActionResource
  '     - Push grid values into ActionResourceNameValues
  '     - Validate model
  '     - Set PendingAction (Add/Update)
  '     - Persist using SavePendingResourceAction
  ' Parameters:
  '   sender - event sender
  '   e      - event args
  ' Returns:
  '   None
  ' Notes:
  '   Sets WasSaved = True and ResourceID so caller can refresh.
  ' ==========================================================================================
  Private Sub btnSave_Click(sender As Object, e As EventArgs) Handles btnSave.Click
    Try
      PushUIToModel()
      PushGridToModel()
      ValidateModelOrThrow()

      If String.IsNullOrEmpty(ResourceID) Then
        _model.PendingAction = ResourceAction.Add
      Else
        _model.PendingAction = ResourceAction.Update
      End If

      SavePendingResourceAction(_model)

      WasSaved = True
      ResourceID = _model.ActionResource.ResourceID
      Me.Close()

    Catch ex As UserFriendlyException
      MessageBoxHelper.Show(UIMessageOwner.ExcelOwner,
                            ex.Message,
                            "Validation error",
                            MessageBoxButtons.OK,
                            MessageBoxIcon.Warning)

    Catch ex As Exception
      ErrorHandler.UnHandleError(ex)
    End Try
  End Sub

  ' ==========================================================================================
  ' Routine: PushUIToModel
  ' Purpose:
  '   Map current UI control values into _model.ActionResource.
  ' Parameters:
  '   None
  ' Returns:
  '   None
  ' Notes:
  '   Writes GUID/ID values for Salutation, Gender.
  ' ==========================================================================================
  Private Sub PushUIToModel()

    Dim a = _model.ActionResource

    a.PreferredName = txtPreferredName.Text
    a.FirstName = txtFirstName.Text
    a.LastName = txtLastName.Text

    a.SalutationID = TryCast(cmbSalutation.SelectedValue, String)
    a.GenderID = TryCast(cmbGender.SelectedValue, String)

    a.Email = txtEmail.Text
    a.Phone = txtPhone.Text

    If dtpStartDate.Checked Then
      a.StartDate = dtpStartDate.Value
    Else
      a.StartDate = Date.MinValue
    End If

    If dtpEndDate.Checked Then
      a.EndDate = dtpEndDate.Value
    Else
      a.EndDate = Date.MinValue
    End If

    a.Notes = txtNotes.Text

  End Sub

  ' ==========================================================================================
  ' Routine: ValidateModelOrThrow
  ' Purpose:
  '   Validate ActionResource before saving.
  ' Parameters:
  '   None
  ' Returns:
  '   None
  ' Notes:
  '   - PreferredName is required
  '   - EndDate must be >= StartDate if it is set (i.e., not Date.MinValue)
  ' ==========================================================================================
  Private Sub ValidateModelOrThrow()

    Dim a = _model.ActionResource

    If String.IsNullOrWhiteSpace(a.PreferredName) Then
      Throw New UserFriendlyException("Preferred Name is required.")
    End If

    If a.EndDate <> Date.MinValue AndAlso a.EndDate < a.StartDate Then
      Throw New UserFriendlyException("End Date must be greater than or equal to Start Date.")
    End If

  End Sub

  ' ==========================================================================================
  ' Routine: btnDelete_Click
  ' Purpose:
  '   Delete the current resource (soft delete).
  ' Parameters:
  '   sender - event sender
  '   e      - event args
  ' Returns:
  '   None
  ' Notes:
  '   Only valid when ResourceID is set.
  ' ==========================================================================================
  Private Sub btnDelete_Click(sender As Object, e As EventArgs) Handles btnDelete.Click
    Try
      If String.IsNullOrEmpty(ResourceID) Then Exit Sub

      _model.PendingAction = ResourceAction.Delete
      SavePendingResourceAction(_model)

      WasSaved = True
      Me.Close()

    Catch ex As Exception
      ErrorHandler.UnHandleError(ex)
    End Try
  End Sub

  ' ==========================================================================================
  ' Routine: btnClose_Click
  ' Purpose:
  '   Close the dialog without saving.
  ' Parameters:
  '   sender - event sender
  '   e      - event args
  ' Returns:
  '   None
  ' Notes:
  '   Ensures WasSaved = False so caller does not refresh.
  ' ==========================================================================================
  Private Sub btnClose_Click(sender As Object, e As EventArgs) Handles btnClose.Click
    WasSaved = False
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
  ' Notes:
  ' ==========================================================================================
  Private Sub ComboBox_Validating(sender As Object, e As System.ComponentModel.CancelEventArgs) _
    Handles cmbGender.Validating, cmbSalutation.Validating

    Try
      '=== Forward to shared validator ===
      ValidateComboBoxSelection(DirectCast(sender, ComboBox), e)

    Catch ex As Exception
      ErrorHandler.UnHandleError(ex)
    End Try
  End Sub

End Class