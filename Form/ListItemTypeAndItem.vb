Option Explicit On
Imports System.ComponentModel
Imports System.Windows.Forms

' ==========================================================================================
' Class: ListItemTypeAndItem
' Purpose:
'   Maintenance form for ListItemTypes and ListItems. Allows the user to:
'     - Select a ListItemType
'     - Add / Update / Delete ListItemTypes (except system types)
'     - View ListItems belonging to the selected type
'     - Add / Update / Delete ListItems
' Notes:
'   - Uses UIModelListItemTypeAndItem as the UI model.
'   - Uses UILoaderSaverListItemTypeAndItem for all persistence.
'   - Follows the standard two-path loader pattern for the model.
' ==========================================================================================
Friend Class ListItemTypeAndItem

  ' ==========================================================================================
  ' Private fields
  ' ==========================================================================================
  Private _model As UIModelListItemTypeAndItem
  Private _suppressEvents As Boolean = False

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
  ' Routine: ListItem_Load
  ' Purpose:
  '   Initialise the form, load the model, and bind controls.
  ' Parameters:
  '   sender - event sender
  '   e      - event args
  ' Returns:
  '   None
  ' ==========================================================================================
  Private Sub ListItem_Load(sender As Object, e As EventArgs) Handles MyBase.Load
    Try
      Try
        FormHelpers.CenterFormOnExcel(Me)
      Catch
        ' Ignore if helper not available
      End Try

      UILoaderSaverListItemTypeAndItem.LoadListItemTypeAndItemModel(_model)
      BindListItemTypes()
      ClearListItemTypeEditor()
      ClearListItemEditor()

    Catch ex As Exception
      ErrorHandler.UnHandleError(ex)
    End Try
  End Sub

  ' ==========================================================================================
  ' Routine: BindListItemTypes
  ' Purpose:
  '   Bind the ListItemTypes listbox to the model.
  ' Parameters:
  '   None
  ' Returns:
  '   None
  ' ==========================================================================================
  Private Sub BindListItemTypes()
    _suppressEvents = True
    lstListItemTypes.DataSource = _model.ListItemTypes
    lstListItemTypes.DisplayMember = "ListItemTypeName"
    lstListItemTypes.ValueMember = "ListItemTypeID"
    _suppressEvents = False
  End Sub

  ' ==========================================================================================
  ' Routine: BindListItems
  ' Purpose:
  '   Bind the ListItems listbox to the model for the selected ListItemType.
  ' Parameters:
  '   None
  ' Returns:
  '   None
  ' ==========================================================================================
  Private Sub BindListItems()
    lstListItems.DataSource = _model.ListItems
    lstListItems.DisplayMember = "ListItemName"
    lstListItems.ValueMember = "ListItemID"
  End Sub

  ' ==========================================================================================
  ' Routine: lstListItemTypes_SelectedIndexChanged
  ' Purpose:
  '   When the user selects a ListItemType, refresh ListItems and update the type editor.
  ' Parameters:
  '   sender - event sender
  '   e      - event args
  ' Returns:
  '   None
  ' ==========================================================================================
  Private Sub lstListItemTypes_SelectedIndexChanged(sender As Object, e As EventArgs) _
      Handles lstListItemTypes.SelectedIndexChanged

    If _suppressEvents Then Exit Sub

    Try
      Dim selected As UIListItemTypeRow = TryCast(lstListItemTypes.SelectedItem, UIListItemTypeRow)
      _model.SelectedListItemType = selected

      If selected Is Nothing Then
        ClearListItemTypeEditor()
        _model.ListItems.Clear()
        lstListItems.DataSource = Nothing
        ClearListItemEditor()
        Exit Sub
      End If

      txtListItemTypeName.Text = selected.ListItemTypeName

      ' Loader pattern: existing model, so loader refreshes items for current type
      UILoaderSaverListItemTypeAndItem.LoadListItemTypeAndItemModel(_model)
      BindListItems()
      ClearListItemEditor()

    Catch ex As Exception
      ErrorHandler.UnHandleError(ex)
    End Try
  End Sub

  ' ==========================================================================================
  ' Routine: lstListItems_SelectedIndexChanged
  ' Purpose:
  '   When the user selects a ListItem, update the ListItem editor textbox.
  ' Parameters:
  '   sender - event sender
  '   e      - event args
  ' Returns:
  '   None
  ' ==========================================================================================
  Private Sub lstListItems_SelectedIndexChanged(sender As Object, e As EventArgs) _
      Handles lstListItems.SelectedIndexChanged

    Try
      Dim selected As UIListItemRow = TryCast(lstListItems.SelectedItem, UIListItemRow)

      If selected Is Nothing Then
        ClearListItemEditor()
        Exit Sub
      End If

      txtListItemName.Text = selected.ListItemName

    Catch ex As Exception
      ErrorHandler.UnHandleError(ex)
    End Try
  End Sub

  ' ==========================================================================================
  ' Routine: ClearListItemTypeEditor
  ' Purpose:
  '   Reset the ListItemType editor fields.
  ' Parameters:
  '   None
  ' Returns:
  '   None
  ' ==========================================================================================
  Private Sub ClearListItemTypeEditor()
    txtListItemTypeName.Text = ""
  End Sub

  ' ==========================================================================================
  ' Routine: ClearListItemEditor
  ' Purpose:
  '   Reset the ListItem editor fields.
  ' Parameters:
  '   None
  ' Returns:
  '   None
  ' ==========================================================================================
  Private Sub ClearListItemEditor()
    txtListItemName.Text = ""
  End Sub

  ' ==========================================================================================
  ' Routine: ValidateListItemTypeName
  ' Purpose:
  '   Ensure txtListItemTypeName is not empty.
  ' Parameters:
  '   None
  ' Returns:
  '   Boolean - True if valid, False otherwise.
  ' ==========================================================================================
  Private Function ValidateListItemTypeName() As Boolean
    If String.IsNullOrWhiteSpace(txtListItemTypeName.Text) Then
      MessageBoxHelper.Show(UIMessageOwner.ExcelOwner,
                            "Please enter a list item type name.",
                            "Validation",
                            MessageBoxButtons.OK,
                            MessageBoxIcon.Warning)
      Return False
    End If
    Return True
  End Function

  ' ==========================================================================================
  ' Routine: ValidateListItemName
  ' Purpose:
  '   Ensure txtListItemName is not empty.
  ' Parameters:
  '   None
  ' Returns:
  '   Boolean - True if valid, False otherwise.
  ' ==========================================================================================
  Private Function ValidateListItemName() As Boolean
    If String.IsNullOrWhiteSpace(txtListItemName.Text) Then
      MessageBoxHelper.Show(UIMessageOwner.ExcelOwner,
                            "Please enter a list item name.",
                            "Validation",
                            MessageBoxButtons.OK,
                            MessageBoxIcon.Warning)
      Return False
    End If
    Return True
  End Function

  ' ==========================================================================================
  ' Routine: btnAddNewListItemType_Click
  ' Purpose:
  '   Add a new ListItemType.
  ' Parameters:
  '   sender - event sender
  '   e      - event args
  ' Returns:
  '   None
  ' ==========================================================================================
  Private Sub btnAddNewListItemType_Click(sender As Object, e As EventArgs) _
      Handles btnAddNewListItemType.Click

    Try
      If Not ValidateListItemTypeName() Then Exit Sub

      _model.ActionListItemType = New UIListItemTypeRow With {
        .ListItemTypeName = txtListItemTypeName.Text.Trim(),
        .IsSystemType = False
      }
      _model.PendingListItemTypeAction = ListItemTypeAction.Add

      UILoaderSaverListItemTypeAndItem.SaveListItemTypeAction(_model)

      BindListItemTypes()
      ClearListItemTypeEditor()
      ClearListItemEditor()
      lstListItems.DataSource = Nothing

    Catch ex As Exception
      ErrorHandler.UnHandleError(ex)
    End Try
  End Sub

  ' ==========================================================================================
  ' Routine: btnUpdateListItemType_Click
  ' Purpose:
  '   Update the selected ListItemType.
  ' Parameters:
  '   sender - event sender
  '   e      - event args
  ' Returns:
  '   None
  ' ==========================================================================================
  Private Sub btnUpdateListItemType_Click(sender As Object, e As EventArgs) _
      Handles btnUpdateListItemType.Click

    Try
      Dim selected As UIListItemTypeRow = TryCast(lstListItemTypes.SelectedItem, UIListItemTypeRow)
      If selected Is Nothing Then
        MessageBoxHelper.Show(UIMessageOwner.ExcelOwner,
                              "Please select a list item type to update.",
                              "No Selection",
                              MessageBoxButtons.OK,
                              MessageBoxIcon.Information)
        Exit Sub
      End If

      If Not ValidateListItemTypeName() Then Exit Sub

      _model.ActionListItemType = New UIListItemTypeRow With {
        .ListItemTypeID = selected.ListItemTypeID,
        .ListItemTypeName = txtListItemTypeName.Text.Trim(),
        .IsSystemType = selected.IsSystemType
      }
      _model.PendingListItemTypeAction = ListItemTypeAction.Update

      UILoaderSaverListItemTypeAndItem.SaveListItemTypeAction(_model)

      BindListItemTypes()
      ClearListItemTypeEditor()
      ClearListItemEditor()
      lstListItems.DataSource = Nothing

    Catch ex As Exception
      ErrorHandler.UnHandleError(ex)
    End Try
  End Sub

  ' ==========================================================================================
  ' Routine: btnDeleteListItemType_Click
  ' Purpose:
  '   Delete the selected ListItemType (unless it is a system type).
  ' Parameters:
  '   sender - event sender
  '   e      - event args
  ' Returns:
  '   None
  ' ==========================================================================================
  Private Sub btnDeleteListItemType_Click(sender As Object, e As EventArgs) _
      Handles btnDeleteListItemType.Click

    Try
      Dim selected As UIListItemTypeRow = TryCast(lstListItemTypes.SelectedItem, UIListItemTypeRow)
      If selected Is Nothing Then
        MessageBoxHelper.Show(UIMessageOwner.ExcelOwner,
                              "Please select a list item type to delete.",
                              "No Selection",
                              MessageBoxButtons.OK,
                              MessageBoxIcon.Information)
        Exit Sub
      End If

      If selected.IsSystemType Then
        MessageBoxHelper.Show(UIMessageOwner.ExcelOwner,
                              "System list item types cannot be deleted.",
                              "Not Allowed",
                              MessageBoxButtons.OK,
                              MessageBoxIcon.Warning)
        Exit Sub
      End If

      If MessageBoxHelper.Show(UIMessageOwner.ExcelOwner,
                               "Are you sure you want to delete the selected list item type?",
                               "Confirm Delete",
                               MessageBoxButtons.YesNo,
                               MessageBoxIcon.Question) <> DialogResult.Yes Then
        Exit Sub
      End If

      _model.ActionListItemType = selected
      _model.PendingListItemTypeAction = ListItemTypeAction.Delete

      UILoaderSaverListItemTypeAndItem.SaveListItemTypeAction(_model)

      BindListItemTypes()
      ClearListItemTypeEditor()
      ClearListItemEditor()
      lstListItems.DataSource = Nothing

    Catch ex As Exception
      ErrorHandler.UnHandleError(ex)
    End Try
  End Sub

  ' ==========================================================================================
  ' Routine: btnAddNewListItem_Click
  ' Purpose:
  '   Add a new ListItem under the selected ListItemType.
  ' Parameters:
  '   sender - event sender
  '   e      - event args
  ' Returns:
  '   None
  ' ==========================================================================================
  Private Sub btnAddNewListItem_Click(sender As Object, e As EventArgs) _
      Handles btnAddNewListItem.Click

    Try
      If _model.SelectedListItemType Is Nothing Then
        MessageBoxHelper.Show(UIMessageOwner.ExcelOwner,
                              "Please select a list item type first.",
                              "No Type Selected",
                              MessageBoxButtons.OK,
                              MessageBoxIcon.Information)
        Exit Sub
      End If

      If Not ValidateListItemName() Then Exit Sub

      _model.ActionListItem = New UIListItemRow With {
        .ListItemTypeID = _model.SelectedListItemType.ListItemTypeID,
        .ListItemName = txtListItemName.Text.Trim()
      }
      _model.PendingListItemAction = ListItemAction.Add

      UILoaderSaverListItemTypeAndItem.SaveListItemAction(_model)

      BindListItems()
      ClearListItemEditor()

    Catch ex As Exception
      ErrorHandler.UnHandleError(ex)
    End Try
  End Sub

  ' ==========================================================================================
  ' Routine: btnUpdateListItem_Click
  ' Purpose:
  '   Update the selected ListItem.
  ' Parameters:
  '   sender - event sender
  '   e      - event args
  ' Returns:
  '   None
  ' ==========================================================================================
  Private Sub btnUpdateListItem_Click(sender As Object, e As EventArgs) _
      Handles btnUpdateListItem.Click

    Try
      Dim selected As UIListItemRow = TryCast(lstListItems.SelectedItem, UIListItemRow)
      If selected Is Nothing Then
        MessageBoxHelper.Show(UIMessageOwner.ExcelOwner,
                              "Please select a list item to update.",
                              "No Selection",
                              MessageBoxButtons.OK,
                              MessageBoxIcon.Information)
        Exit Sub
      End If

      If Not ValidateListItemName() Then Exit Sub

      _model.ActionListItem = New UIListItemRow With {
        .ListItemID = selected.ListItemID,
        .ListItemTypeID = selected.ListItemTypeID,
        .ListItemName = txtListItemName.Text.Trim()
      }
      _model.PendingListItemAction = ListItemAction.Update

      UILoaderSaverListItemTypeAndItem.SaveListItemAction(_model)

      BindListItems()
      ClearListItemEditor()

    Catch ex As Exception
      ErrorHandler.UnHandleError(ex)
    End Try
  End Sub

  ' ==========================================================================================
  ' Routine: btnDeleteListItem_Click
  ' Purpose:
  '   Delete the selected ListItem.
  ' Parameters:
  '   sender - event sender
  '   e      - event args
  ' Returns:
  '   None
  ' ==========================================================================================
  Private Sub btnDeleteListItem_Click(sender As Object, e As EventArgs) _
      Handles btnDeleteListItem.Click

    Try
      Dim selected As UIListItemRow = TryCast(lstListItems.SelectedItem, UIListItemRow)
      If selected Is Nothing Then
        MessageBoxHelper.Show(UIMessageOwner.ExcelOwner,
                              "Please select a list item to delete.",
                              "No Selection",
                              MessageBoxButtons.OK,
                              MessageBoxIcon.Information)
        Exit Sub
      End If

      If MessageBoxHelper.Show(UIMessageOwner.ExcelOwner,
                               "Are you sure you want to delete the selected list item?",
                               "Confirm Delete",
                               MessageBoxButtons.YesNo,
                               MessageBoxIcon.Question) <> DialogResult.Yes Then
        Exit Sub
      End If

      _model.ActionListItem = selected
      _model.PendingListItemAction = ListItemAction.Delete

      UILoaderSaverListItemTypeAndItem.SaveListItemAction(_model)

      BindListItems()
      ClearListItemEditor()

    Catch ex As Exception
      ErrorHandler.UnHandleError(ex)
    End Try
  End Sub

  ' ==========================================================================================
  ' Routine: btnClose_Click
  ' Purpose:
  '   Close the form.
  ' Parameters:
  '   sender - event sender
  '   e      - event args
  ' Returns:
  '   None
  ' ==========================================================================================
  Private Sub btnClose_Click(sender As Object, e As EventArgs) Handles btnClose.Click
    Me.Close()
  End Sub

End Class
