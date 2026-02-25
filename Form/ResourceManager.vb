Option Explicit On
Imports System.Windows.Forms
Imports System.Drawing
Friend Class ResourceManager

  Private _model As UIModelResourceManager

  '=== useed for sorting lstResource ===
  Private _sortColumn As Integer = -1
  Private _sortAscending As Boolean = True
  Friend Sub New()
    Try
      ' Disable WinForms autoscaling completely
      Me.AutoScaleMode = AutoScaleMode.None

      InitializeComponent()

      ' Apply manual DPI scaling AFTER controls exist
      ApplyDpiScaling(Me)

    Catch ex As Exception
      ErrorHandler.UnHandleError(ex)

    Finally
      ' Cleanup
    End Try
  End Sub

  ' ==========================================================================================
  ' Routine: ResourceManager_Load
  ' Purpose: Initialise form, center on Excel, load model, populate resource list.
  ' Parameters:
  '   sender - event sender
  '   e      - event args
  ' Returns:
  '   (None)
  ' Notes:
  '   Uses UILoaderSaverResourceManager to load model.
  ' ==========================================================================================
  Private Sub ResourceManager_Load(sender As Object, e As EventArgs) Handles MyBase.Load
    Try
      '=== Center form on Excel ===
      FormHelpers.CenterFormOnExcel(Me)

      '=== Display the defaukt availability mode ===
      Dim defaultMode As String = LoadMetadataValue("DefaultMode")
      lblDefaultMode.Text = "Resources are " & defaultMode & " by default"

      '=== Initialise lstResource ===
      lstResources.View = View.Details
      lstResources.FullRowSelect = True
      lstResources.GridLines = True
      lstResources.Columns.Clear()
      lstResources.Columns.Add("Preferred Name")
      lstResources.Columns.Add("Full Name")
      'lstResources.Columns.Add("Employee ID", 100)
      lstResources.Columns.Add("Status")
      lstResources.ColumnPercents = {35, 50, 15}
      lstResources.FullRowSelect = True

      '=== Load model ===
      UILoaderSaverResourceManager.LoadResourceManagerModel(_model)

      '=== Populate resources ===
      PopulateResourceList()

      '=== Initialise options ===
      optPreferredName.Checked = True
      chkShowInactive.Checked = False
      txtFilter.Text = ""

    Catch ex As Exception
      ErrorHandler.UnHandleError(ex)
    End Try
  End Sub

  ' ==========================================================================================
  ' Routine: PopulateResourceList
  ' Purpose: Populate lstResources manually using model.ResourceSummaries.
  ' Parameters:
  '   selectResourceID - ID to reselect after refresh (optional)
  ' Returns:
  '   (None)
  ' Notes:
  '   Applies filtering and inactive toggle.
  ' ==========================================================================================
  Private Sub PopulateResourceList(Optional ByVal selectResourceID As String = "")
    Try
      lstResources.BeginUpdate()
      lstResources.Items.Clear()

      '=== Filtering ===
      Dim filterText As String = txtFilter.Text.Trim().ToLower()
      Dim filterMode As ResourceManagerFilterMode = GetFilterMode()

      For Each r In _model.ResourceSummaries

        '=== Skip inactive if checkbox unchecked ===
        If Not chkShowInactive.Checked AndAlso r.IsInactive Then Continue For

        '=== Apply filter ===
        If filterText <> "" Then
          Dim fieldValue As String = ""

          Select Case filterMode
            Case ResourceManagerFilterMode.PreferredName
              fieldValue = r.PreferredName
            Case ResourceManagerFilterMode.FullName
              fieldValue = r.FullName
              'Case ResourceManagerFilterMode.EmployeeID
              '  fieldValue = r.EmployeeID
          End Select
          If fieldValue Is Nothing Then Continue For
          Dim text As String = fieldValue.ToString() ' makes it is string before calling ToLower
          If Not text.ToLower().Contains(filterText) Then Continue For

        End If

        '=== Build ListViewItem ===
        Dim item As New ListViewItem(r.PreferredName)
        item.SubItems.Add(r.FullName)
        'item.SubItems.Add(r.EmployeeID)
        item.SubItems.Add(If(r.IsInactive, "Inactive", ""))

        item.Tag = r.ResourceID

        lstResources.Items.Add(item)
      Next

      lstResources.EndUpdate()

      '=== Restore selection ===
      If selectResourceID <> "" Then
        For Each item As ListViewItem In lstResources.Items
          If CStr(item.Tag) = selectResourceID Then
            item.Selected = True
            item.EnsureVisible()
            Exit For
          End If
        Next
      End If

    Catch ex As Exception
      ErrorHandler.UnHandleError(ex)
    End Try
  End Sub

  ' ==========================================================================================
  ' Routine: GetFilterMode
  ' Purpose: Determine which field to filter on based on radio buttons.
  ' Parameters:
  '   (None)
  ' Returns:
  '   ResourceManagerFilterMode
  ' Notes:
  '   Defaults to EmployeeID.
  ' ==========================================================================================
  Private Function GetFilterMode() As ResourceManagerFilterMode
    If optPreferredName.Checked Then Return ResourceManagerFilterMode.PreferredName
    If optFullName.Checked Then Return ResourceManagerFilterMode.FullName
    Return ResourceManagerFilterMode.EmployeeID
  End Function

  ' ==========================================================================================
  ' Routine: lstResources_SelectedIndexChanged
  ' Purpose: Update model.SelectedResourceID and load availability list.
  ' Parameters:
  '   sender - event sender
  '   e      - event args
  ' Returns:
  '   (None)
  ' Notes:
  '   Clears availability list if no resource selected.
  ' ==========================================================================================
  Private Sub lstResources_SelectedIndexChanged(sender As Object, e As EventArgs) Handles lstResources.SelectedIndexChanged
    Try
      If lstResources.SelectedItems.Count = 0 Then
        _model.SelectedResourceID = ""
        _model.SelectedAvailabilityID = ""
        lstResourceAvailability.DataSource = Nothing
        Exit Sub
      End If

      '=== Update selected resource ===
      Dim selectedID As String = CStr(lstResources.SelectedItems(0).Tag)
      _model.SelectedResourceID = selectedID

      '=== Reload availability ===
      UILoaderSaverResourceManager.LoadResourceManagerModel(_model)
      PopulateAvailabilityList()

    Catch ex As Exception
      ErrorHandler.UnHandleError(ex)
    End Try
  End Sub

  ' ==========================================================================================
  ' Routine: lstResources_ColumnClick
  ' Purpose: Sort ListView by clicked column.
  ' Parameters:
  '   sender - event sender
  '   e      - column click args
  ' Returns:
  '   (None)
  ' Notes:
  '   Toggles ascending/descending when clicking same column.
  ' ==========================================================================================
  Private Sub lstResources_ColumnClick(sender As Object, e As ColumnClickEventArgs) Handles lstResources.ColumnClick
    Try
      '=== Toggle sort direction if same column ===
      If e.Column = _sortColumn Then
        _sortAscending = Not _sortAscending
      Else
        _sortColumn = e.Column
        _sortAscending = True
      End If

      '=== Apply sorter ===
      lstResources.ListViewItemSorter = New ListViewItemComparer(_sortColumn, _sortAscending)
      lstResources.Sort()

    Catch ex As Exception
      ErrorHandler.UnHandleError(ex)
    End Try
  End Sub

  ' ==========================================================================================
  ' Routine: lstResources_DoubleClick
  ' Purpose: Forward double-click on resource list to Edit Resource.
  ' Parameters:
  '   sender - event sender
  '   e      - event args
  ' Returns:
  '   (None)
  ' Notes:
  '   Calls btnEditResource_Click to reuse logic.
  ' ==========================================================================================
  Private Sub lstResources_DoubleClick(sender As Object, e As EventArgs) Handles lstResources.MouseDoubleClick
    btnEditResource_Click(sender, e)
  End Sub

  ' ==========================================================================================
  ' Routine: lstResourceAvailability_DoubleClick
  ' Purpose: Forward double-click on availability list to Edit Availability.
  ' Parameters:
  '   sender - event sender
  '   e      - event args
  ' Returns:
  '   (None)
  ' Notes:
  '   Calls btnEditAvailability_Click to reuse logic.
  ' ==========================================================================================
  Private Sub lstResourceAvailability_DoubleClick(sender As Object, e As EventArgs) Handles lstResourceAvailability.DoubleClick
    btnEditAvailability_Click(sender, e)
  End Sub

  ' ==========================================================================================
  ' Routine: PopulateAvailabilityList
  ' Purpose: Bind availability list to model.AvailabilitySummaries.
  ' Parameters:
  '   selectAvailabilityID - ID to reselect after refresh (optional)
  ' Returns:
  '   (None)
  ' Notes:
  '   Uses ToString() override on UIResourceManagerAvailabilitySummaryRow.
  ' ==========================================================================================
  Private Sub PopulateAvailabilityList(Optional ByVal selectAvailabilityID As String = "")
    Try
      lstResourceAvailability.DataSource = Nothing
      lstResourceAvailability.DataSource = _model.AvailabilitySummaries

      '=== Restore selection ===
      If selectAvailabilityID <> "" Then
        For i As Integer = 0 To lstResourceAvailability.Items.Count - 1
          Dim row = DirectCast(lstResourceAvailability.Items(i), UIResourceManagerAvailabilitySummaryRow)
          If row.AvailabilityID = selectAvailabilityID Then
            lstResourceAvailability.SelectedIndex = i
            Exit For
          End If
        Next
      End If

    Catch ex As Exception
      ErrorHandler.UnHandleError(ex)
    End Try
  End Sub

  ' ==========================================================================================
  ' Routine: FilterChanged
  ' Purpose: Rebuild resource list when filter controls change.
  ' Parameters:
  '   sender - event sender
  '   e      - event args
  ' Returns:
  '   (None)
  ' Notes:
  '   Applies current model.SelectedResourceID.
  ' ==========================================================================================
  Private Sub txtFilter_TextChanged(sender As Object, e As EventArgs) Handles txtFilter.TextChanged
    Try
      PopulateResourceList(_model.SelectedResourceID)
    Catch ex As Exception
      ErrorHandler.UnHandleError(ex)
    End Try
  End Sub

  Private Sub chkShowInactive_CheckedChanged(sender As Object, e As EventArgs) Handles chkShowInactive.CheckedChanged
    Try
      PopulateResourceList(_model.SelectedResourceID)
    Catch ex As Exception
      ErrorHandler.UnHandleError(ex)
    End Try
  End Sub

  Private Sub optPreferredName_CheckedChanged(sender As Object, e As EventArgs) Handles optPreferredName.CheckedChanged
    Try
      PopulateResourceList(_model.SelectedResourceID)
    Catch ex As Exception
      ErrorHandler.UnHandleError(ex)
    End Try
  End Sub

  Private Sub optFullName_CheckedChanged(sender As Object, e As EventArgs) Handles optFullName.CheckedChanged
    Try
      PopulateResourceList(_model.SelectedResourceID)
    Catch ex As Exception
      ErrorHandler.UnHandleError(ex)
    End Try
  End Sub

  'Private Sub optEmployeeID_CheckedChanged(sender As Object, e As EventArgs) Handles
  '  Try
  '    PopulateResourceList(_model.SelectedResourceID)
  '  Catch ex As Exception
  '    ErrorHandler.UnHandleError(ex)
  '  End Try
  'End Sub

  ' ==========================================================================================
  ' Routine: btnNewResource_Click
  ' Purpose: Open Resource form for new resource, reload list on save.
  ' Parameters:
  '   sender - event sender
  '   e      - event args
  ' Returns:
  '   (None)
  ' Notes:
  '   Reselects new resource after save.
  ' ==========================================================================================
  Private Sub btnNewResource_Click(sender As Object, e As EventArgs) Handles btnNewResource.Click
    Try
      Dim f As New Resource()
      f.ResourceID = ""
      f.ShowDialog()

      If f.WasSaved Then
        ' Force full resource reload (no availabilities yet)
        _model = Nothing
        UILoaderSaverResourceManager.LoadResourceManagerModel(_model)

        PopulateResourceList(f.ResourceID)   ' rebind, reselect edited resource
        ' If the availability panel is driven by resource selection,
        ' its existing selection-changed logic will reload availabilities as needed.
      End If

    Catch ex As Exception
      ErrorHandler.UnHandleError(ex)
    End Try
  End Sub

  ' ==========================================================================================
  ' Routine: btnEditResource_Click
  ' Purpose: Open Resource form for editing, reload list on save.
  ' Parameters:
  '   sender - event sender
  '   e      - event args
  ' Returns:
  '   (None)
  ' Notes:
  '   Reselects edited resource after save.
  ' ==========================================================================================
  Private Sub btnEditResource_Click(sender As Object, e As EventArgs) Handles btnEditResource.Click
    Try
      If lstResources.SelectedItems.Count = 0 Then Exit Sub

      Dim id As String = CStr(lstResources.SelectedItems(0).Tag)

      Dim f As New Resource()
      f.ResourceID = id
      f.ShowDialog()

      If f.WasSaved Then
        ' Force full resource reload (no availabilities yet)
        _model = Nothing
        UILoaderSaverResourceManager.LoadResourceManagerModel(_model)

        PopulateResourceList(f.ResourceID)   ' rebind, reselect edited resource
        ' If the availability panel is driven by resource selection,
        ' its existing selection-changed logic will reload availabilities as needed.
      End If

    Catch ex As Exception
      ErrorHandler.UnHandleError(ex)
    End Try
  End Sub


  ' ==========================================================================================
  ' Routine: btnNewAvailability_Click
  ' Purpose: Open Availability form for new availability, reload list on save.
  ' Parameters:
  '   sender - event sender
  '   e      - event args
  ' Returns:
  '   (None)
  ' Notes:
  '   Reselects new availability after save.
  ' ==========================================================================================
  Private Sub btnNewAvailability_Click(sender As Object, e As EventArgs) Handles btnNewAvailability.Click
    Try
      If _model.SelectedResourceID = "" Then Exit Sub

      Dim f As New Availability()
      f.resourceID = _model.SelectedResourceID
      f.availabilityID = ""
      f.ShowDialog()

      If f.wasSaved Then
        _model.SelectedAvailabilityID = f.availabilityID
        UILoaderSaverResourceManager.LoadResourceManagerModel(_model)
        PopulateAvailabilityList(f.availabilityID)
      End If

    Catch ex As Exception
      ErrorHandler.UnHandleError(ex)
    End Try
  End Sub

  ' ==========================================================================================
  ' Routine: btnEditAvailability_Click
  ' Purpose: Open Availability form for editing, reload list on save.
  ' Parameters:
  '   sender - event sender
  '   e      - event args
  ' Returns:
  '   (None)
  ' Notes:
  '   Reselects edited availability after save.
  ' ==========================================================================================
  Private Sub btnEditAvailability_Click(sender As Object, e As EventArgs) Handles btnEditAvailability.Click
    Try
      If lstResourceAvailability.SelectedIndex < 0 Then Exit Sub

      Dim row = DirectCast(lstResourceAvailability.SelectedItem, UIResourceManagerAvailabilitySummaryRow)
      Dim id As String = row.AvailabilityID

      Dim f As New Availability()
      f.resourceID = _model.SelectedResourceID
      f.availabilityID = id
      f.ShowDialog()

      If f.wasSaved Then
        _model.SelectedAvailabilityID = f.availabilityID
        UILoaderSaverResourceManager.LoadResourceManagerModel(_model)
        PopulateAvailabilityList(f.availabilityID)
      End If

    Catch ex As Exception
      ErrorHandler.UnHandleError(ex)
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
  ' Class: ListViewItemComparer
  ' Purpose:
  '   Provides column‑based sorting for a ListView when the user clicks a column header.
  '   Implements IComparer so the ListView can sort its items using this comparer.
  ' Parameters:
  '   (Constructor)
  '     column    - The zero‑based index of the column to sort by.
  '     ascending - True for ascending sort, False for descending.
  ' Returns:
  '   (Compare method)
  '     Integer indicating sort order between two ListViewItem objects.
  ' Notes:
  '   - Used by assigning an instance to ListView.ListViewItemSorter.
  '   - Sorting is string‑based using the text of the selected column.
  '   - The ResourceManager form toggles ascending/descending when the same column is clicked.
  ' ==========================================================================================
  Private Class ListViewItemComparer
    Implements IComparer

    Private ReadOnly _column As Integer
    Private ReadOnly _ascending As Boolean

    '=== Constructor: store column index and sort direction ===
    Friend Sub New(column As Integer, ascending As Boolean)
      _column = column
      _ascending = ascending
    End Sub

    ' ==========================================================================================
    ' Routine: Compare
    ' Purpose: Compare two ListViewItem objects based on the selected column.
    ' Parameters:
    '   x - First ListViewItem
    '   y - Second ListViewItem
    ' Returns:
    '   Integer:
    '     < 0  if x < y
    '     = 0  if x = y
    '     > 0  if x > y
    ' Notes:
    '   - Comparison is case‑insensitive.
    '   - Direction is controlled by _ascending.
    ' ==========================================================================================
    Friend Function Compare(x As Object, y As Object) As Integer Implements IComparer.Compare
      '=== Cast to ListViewItem ===
      Dim itemX As ListViewItem = DirectCast(x, ListViewItem)
      Dim itemY As ListViewItem = DirectCast(y, ListViewItem)

      '=== Compare the text in the selected column ===
      Dim result As Integer = String.Compare(
        itemX.SubItems(_column).Text,
        itemY.SubItems(_column).Text,
        StringComparison.CurrentCultureIgnoreCase)

      '=== Apply ascending/descending ===
      Return If(_ascending, result, -result)
    End Function

  End Class

  Private Sub lstResources_DoubleClick(sender As Object, e As MouseEventArgs) Handles lstResources.MouseDoubleClick

  End Sub
End Class

