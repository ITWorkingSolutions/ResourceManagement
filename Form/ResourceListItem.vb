Option Explicit On
Imports System.ComponentModel
Imports System.Windows.Forms

' ==========================================================================================
' Class: ResourceListItem
' Purpose: Maintenance form for ResourceListItems, allowing view, add, update, and delete, with optional ListItemType assignment (including None).
' Notes:
'   Uses UIModelResourceListItem as the UI model and UILoaderSaverResourceListItem for all persistence. Form knows nothing about the database.
' ==========================================================================================
Friend Class ResourceListItem

  Private _model As UIModelResourceListItem
  Private _suppressEvents As Boolean = False

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
  ' Routine: ResourceListItem_Load
  ' Purpose: Initialise the form, load the model, and bind controls.
  ' Parameters:
  '   sender - Event sender.
  '   e      - Event args.
  ' Returns:
  '   None
  ' Notes:
  '   Uses the standard two-path loader pattern via UILoaderSaverResourceListItem.LoadResourceListItemModel.
  ' ==========================================================================================
  Private Sub ResourceListItem_Load(sender As Object, e As EventArgs) Handles MyBase.Load
    Try
      Try
        FormHelpers.CenterFormOnExcel(Me)
      Catch
        ' Ignore if helper not available.
      End Try

      UILoaderSaverResourceListItem.LoadResourceListItemModel(_model)

      InitialiseListView()
      BindListItemTypes()
      BindResourceListItems()
      BindValueTypes()
      ClearEditor()

    Catch ex As Exception
      ErrorHandler.UnHandleError(ex)
    End Try
  End Sub

  ' ==========================================================================================
  ' Routine: InitialiseListView
  ' Purpose: Configure the ListView columns and behaviour for ResourceListItems.
  ' Parameters:
  '   None
  ' Returns:
  '   None
  ' Notes:
  '   Adds two columns: Resource List Item Name and List Item Type, with full-row select enabled.
  ' ==========================================================================================
  Private Sub InitialiseListView()
    lstResourceListItems.View = View.Details
    lstResourceListItems.FullRowSelect = True
    lstResourceListItems.Columns.Clear()

    lstResourceListItems.Columns.Add("Resource List Item Name", 207)
    lstResourceListItems.Columns.Add("Value Type", 114)
    lstResourceListItems.Columns.Add("List Item Type", 207)

    DisableListViewHeader(lstResourceListItems)
  End Sub

  ' ==========================================================================================
  ' Routine: BindListItemTypes
  ' Purpose: Bind the ListItemTypes combo box to the model.
  ' Parameters:
  '   None
  ' Returns:
  '   None
  ' Notes:
  '   Assumes model.ListItemTypes includes an explicit "None" row with ListItemTypeID = "".
  ' ==========================================================================================
  Private Sub BindListItemTypes()
    cmbListItemTypes.DataSource = _model.ListItemTypes
    cmbListItemTypes.DisplayMember = "ListItemTypeName"
    cmbListItemTypes.ValueMember = "ListItemTypeID"
  End Sub

  ' ==========================================================================================
  ' Routine: BindResourceListItems
  ' Purpose: Populate the ListView with ResourceListItems from the model.
  ' Parameters:
  '   None
  ' Returns:
  '   None
  ' Notes:
  '   Each ListViewItem.Tag stores the corresponding UIResourceListItemRow instance.
  ' ==========================================================================================
  Private Sub BindResourceListItems()
    lstResourceListItems.Items.Clear()

    For Each row As UIResourceListItemRow In _model.ResourceListItems
      Dim typeRow As UIListItemTypeRow =
        _model.ListItemTypes.FirstOrDefault(Function(t) t.ListItemTypeID = row.ListItemTypeID)

      Dim typeName As String = If(typeRow Is Nothing, "", typeRow.ListItemTypeName)

      Dim item As New ListViewItem(row.ResourceListItemName)
      item.SubItems.Add(row.ValueType.ToString())
      item.SubItems.Add(typeName)
      item.Tag = row

      lstResourceListItems.Items.Add(item)
    Next
  End Sub
  ' ==========================================================================================
  ' Routine: BindValueTypes
  ' Purpose: Populate the combo box cmbValueTypes with ResourceNameValueType from the model.
  ' Parameters:
  '   None
  ' Returns:
  '   None
  ' Notes:
  ' ==========================================================================================
  Private Sub BindValueTypes()
    cmbValueTypes.DisplayMember = "Display"
    cmbValueTypes.ValueMember = "EnumValue"
    cmbValueTypes.DataSource = ResourceListItemValueTypeMap.BindingList()
  End Sub

  ' ==========================================================================================
  ' Routine: lstResourceListItems_SelectedIndexChanged
  ' Purpose: Update editor fields when the user selects a ResourceListItem in the ListView.
  ' Parameters:
  '   sender - Event sender.
  '   e      - Event args.
  ' Returns:
  '   None
  ' Notes:
  '   Uses the ListViewItem.Tag to retrieve the UIResourceListItemRow for the selection.
  ' ==========================================================================================
  Private Sub lstResourceListItems_SelectedIndexChanged(sender As Object, e As EventArgs) _
      Handles lstResourceListItems.SelectedIndexChanged

    If _suppressEvents Then Exit Sub

    Try
      If lstResourceListItems.SelectedItems.Count = 0 Then
        ClearEditor()
        Exit Sub
      End If

      Dim row As UIResourceListItemRow =
        TryCast(lstResourceListItems.SelectedItems(0).Tag, UIResourceListItemRow)

      If row Is Nothing Then
        ClearEditor()
        Exit Sub
      End If

      _suppressEvents = True

      txtResourceListItemName.Text = row.ResourceListItemName
      'cmbValueTypes.SelectedItem = [Enum].Parse(GetType(ResourceListItemValueType), row.ValueType)
      cmbValueTypes.SelectedValue = row.ValueType
      If row.ListItemTypeID IsNot Nothing Then
        cmbListItemTypes.SelectedValue = row.ListItemTypeID
      End If

      _suppressEvents = False

      ApplyValueTypeRules()

    Catch ex As Exception
      ErrorHandler.UnHandleError(ex)
    End Try
  End Sub

  ' ==========================================================================================
  ' Routine: ApplyValueTypeRules
  ' Purpose: Enforce enable/disable logic on cmbListItemTypes based on selected value type.
  ' Parameters:
  '   None
  ' Returns:
  '   None
  ' Notes:
  ' ==========================================================================================
  Private Sub ApplyValueTypeRules()
    Dim selectedType As ResourceListItemValueType =
      CType(cmbValueTypes.SelectedValue, ResourceListItemValueType)

    Select Case selectedType

      Case ResourceListItemValueType.Text
        cmbListItemTypes.SelectedValue = ""
        cmbListItemTypes.Enabled = False

      Case ResourceListItemValueType.SingleSelectList, ResourceListItemValueType.MultiSelectList
        cmbListItemTypes.Enabled = True

    End Select
  End Sub

  ' ==========================================================================================
  ' Routine: cmbValueTypes_SelectedIndexChanged
  ' Purpose: Calls ApplyValueTypeRules when the selected value type changes.
  ' Parameters:
  '   sender - Event sender.
  '   e      - Event args.
  ' Returns:
  '   None
  ' Notes:
  ' ==========================================================================================
  Private Sub cmbValueTypes_SelectedIndexChanged(sender As Object, e As EventArgs) _
    Handles cmbValueTypes.SelectedIndexChanged

    If _suppressEvents Then Exit Sub
    ApplyValueTypeRules()
  End Sub

  '' ==========================================================================================
  '' Routine: lstResourceListItems_Layout
  '' Purpose: Hide the headers of the ListView every time it is laid out.
  '' Parameters:
  ''   sender - Event sender.
  ''   e      - Event args.
  '' Returns:
  ''   None
  '' Notes:
  '' ==========================================================================================

  'Private Sub lstResourceListItems_Layout(sender As Object, e As LayoutEventArgs) _
  '  Handles lstResourceListItems.Layout

  '  DisableListViewHeader(lstResourceListItems)
  'End Sub

  ' ==========================================================================================
  ' Routine: ClearEditor
  ' Purpose: Reset the editor controls to their default state.
  ' Parameters:
  '   None
  ' Returns:
  '   None
  ' Notes:
  '   Clears the name textbox and selects the "None" ListItemType where available.
  ' ==========================================================================================
  Private Sub ClearEditor()
    txtResourceListItemName.Text = ""
    cmbValueTypes.SelectedValue = ResourceListItemValueType.Text
    cmbListItemTypes.SelectedValue = ""
    cmbListItemTypes.Enabled = False
  End Sub

  ' ==========================================================================================
  ' Routine: ValidateResourceListItemName
  ' Purpose: Ensure the ResourceListItemName field is not empty.
  ' Parameters:
  '   None
  ' Returns:
  '   Boolean - True if the name is valid; False otherwise.
  ' Notes:
  '   Shows a warning message if validation fails.
  ' ==========================================================================================
  Private Function ValidateResourceListItemName() As Boolean
    If String.IsNullOrWhiteSpace(txtResourceListItemName.Text) Then
      MessageBoxHelper.Show(UIMessageOwner.ExcelOwner,
                            "Please enter a resource list item name.",
                            "Validation",
                            MessageBoxButtons.OK,
                            MessageBoxIcon.Warning)
      Return False
    End If
    Return True
  End Function

  ' ==========================================================================================
  ' Routine: ValidateValueType
  ' Purpose: Ensure a ListItemType is selected if the ValueType requires it.
  ' Parameters:
  '   None
  ' Returns:
  '   Boolean - True if the name is valid; False otherwise.
  ' Notes:
  '   Shows a warning message if validation fails.
  ' ==========================================================================================
  Private Function ValidateValueType() As Boolean
    Dim vt As ResourceListItemValueType =
    CType(cmbValueTypes.SelectedValue, ResourceListItemValueType)

    If vt = ResourceListItemValueType.SingleSelectList OrElse vt = ResourceListItemValueType.MultiSelectList Then
      If String.IsNullOrWhiteSpace(CStr(cmbListItemTypes.SelectedValue)) Then
        MessageBoxHelper.Show(UIMessageOwner.ExcelOwner,
                            "Please select a List Item Type for this value type.",
                            "Validation",
                            MessageBoxButtons.OK,
                            MessageBoxIcon.Warning)
        Return False
      End If
    End If

    Return True
  End Function

  ' ==========================================================================================
  ' Routine: btnAddNew_Click
  ' Purpose: Add a new ResourceListItem based on the editor values.
  ' Parameters:
  '   sender - Event sender.
  '   e      - Event args.
  ' Returns:
  '   None
  ' Notes:
  '   Constructs a new UIResourceListItemRow and sets PendingResourceListItemAction to Add.
  ' ==========================================================================================
  Private Sub btnAddNew_Click(sender As Object, e As EventArgs) Handles btnAddNew.Click
    Try
      If Not ValidateResourceListItemName() Then Exit Sub
      If Not ValidateValueType() Then Exit Sub

      _model.ActionResourceListItem = New UIResourceListItemRow With {
        .ResourceListItemName = txtResourceListItemName.Text.Trim(),
        .ListItemTypeID = CStr(cmbListItemTypes.SelectedValue),
        .ValueType = CType(cmbValueTypes.SelectedValue, ResourceListItemValueType)
      }

      _model.PendingResourceListItemAction = ResourceListItemAction.Add

      UILoaderSaverResourceListItem.SaveResourceListItemAction(_model)

      UILoaderSaverResourceListItem.LoadResourceListItemModel(_model)
      BindResourceListItems()
      ClearEditor()

    Catch ex As Exception
      ErrorHandler.UnHandleError(ex)
    End Try
  End Sub

  ' ==========================================================================================
  ' Routine: btnUpdate_Click
  ' Purpose: Update the selected ResourceListItem with the editor values.
  ' Parameters:
  '   sender - Event sender.
  '   e      - Event args.
  ' Returns:
  '   None
  ' Notes:
  '   Creates a new UIResourceListItemRow for the update; never mutates the existing row instance.
  ' ==========================================================================================
  Private Sub btnUpdate_Click(sender As Object, e As EventArgs) Handles btnUpdate.Click
    Try
      If lstResourceListItems.SelectedItems.Count = 0 Then
        MessageBoxHelper.Show(UIMessageOwner.ExcelOwner,
                              "Please select a resource list item to update.",
                              "No Selection",
                              MessageBoxButtons.OK,
                              MessageBoxIcon.Information)
        Exit Sub
      End If

      If Not ValidateResourceListItemName() Then Exit Sub
      If Not ValidateValueType() Then Exit Sub

      Dim selected As UIResourceListItemRow =
        TryCast(lstResourceListItems.SelectedItems(0).Tag, UIResourceListItemRow)

      _model.ActionResourceListItem = New UIResourceListItemRow With {
        .ResourceListItemID = selected.ResourceListItemID,
        .ResourceListItemName = txtResourceListItemName.Text.Trim(),
        .ListItemTypeID = CStr(cmbListItemTypes.SelectedValue),
        .ValueType = CType(cmbValueTypes.SelectedValue, ResourceListItemValueType)
      }

      _model.PendingResourceListItemAction = ResourceListItemAction.Update

      UILoaderSaverResourceListItem.SaveResourceListItemAction(_model)

      UILoaderSaverResourceListItem.LoadResourceListItemModel(_model)
      BindResourceListItems()
      ClearEditor()

    Catch ex As Exception
      ErrorHandler.UnHandleError(ex)
    End Try
  End Sub

  ' ==========================================================================================
  ' Routine: btnDelete_Click
  ' Purpose: Delete the selected ResourceListItem.
  ' Parameters:
  '   sender - Event sender.
  '   e      - Event args.
  ' Returns:
  '   None
  ' Notes:
  '   Uses the selected UIResourceListItemRow as the ActionResourceListItem and sets action to Delete.
  ' ==========================================================================================
  Private Sub btnDelete_Click(sender As Object, e As EventArgs) Handles btnDelete.Click
    Try
      If lstResourceListItems.SelectedItems.Count = 0 Then
        MessageBoxHelper.Show(UIMessageOwner.ExcelOwner,
                              "Please select a resource list item to delete.",
                              "No Selection",
                              MessageBoxButtons.OK,
                              MessageBoxIcon.Information)
        Exit Sub
      End If

      If MessageBoxHelper.Show(UIMessageOwner.ExcelOwner,
                               "Are you sure you want to delete the selected resource list item?",
                               "Confirm Delete",
                               MessageBoxButtons.YesNo,
                               MessageBoxIcon.Question) <> DialogResult.Yes Then
        Exit Sub
      End If

      Dim selected As UIResourceListItemRow =
        TryCast(lstResourceListItems.SelectedItems(0).Tag, UIResourceListItemRow)

      _model.ActionResourceListItem = selected
      _model.PendingResourceListItemAction = ResourceListItemAction.Delete

      UILoaderSaverResourceListItem.SaveResourceListItemAction(_model)

      UILoaderSaverResourceListItem.LoadResourceListItemModel(_model)
      BindResourceListItems()
      ClearEditor()

    Catch ex As Exception
      ErrorHandler.UnHandleError(ex)
    End Try
  End Sub

  ' ==========================================================================================
  ' Routine: btnClose_Click
  ' Purpose: Close the ResourceListItem form.
  ' Parameters:
  '   sender - Event sender.
  '   e      - Event args.
  ' Returns:
  '   None
  ' Notes:
  '   Standard close handler.
  ' ==========================================================================================
  Private Sub btnClose_Click(sender As Object, e As EventArgs) Handles btnClose.Click
    Me.Close()
  End Sub

End Class