Option Explicit On
Imports System.ComponentModel

' ==========================================================================================
' Module: UILoaderSaverListItemTypeAndItem
' Purpose:
'   Loader and saver logic for the List Item Type + List Item maintenance form.
'   This module is the ONLY supported mechanism for:
'     - Loading ListItemTypes and ListItems into UIModelListItemTypeAndItem
'     - Persisting Add / Update / Delete actions for ListItemTypes
'     - Persisting Add / Update / Delete actions for ListItems
' Responsibilities:
'   - Follow the two-path loader pattern:
'       Path 1: model Is Nothing → initialise model and load ListItemTypes
'       Path 2: model exists → refresh ListItems for the selected type
'   - Enforce system-type protection (IsSystemType = True cannot be deleted)
'   - Use ONLY RecordLoader and RecordSaver (no direct SQL)
'   - Map between domain records and UI rows
' Notes:
'   - All ListItemTypes and ListItems use GUID primary keys.
'   - System types are seeded at DB creation and must never be deleted.
'   - This module must remain deterministic and drift-free.
' ==========================================================================================
Module UILoaderSaverListItemTypeAndItem

  ' ==========================================================================================
  ' Routine: LoadListItemTypeAndItemModel
  ' Purpose:
  '   Load or refresh the UIModelListItemTypeAndItem following the standard two-path loader
  '   pattern used across the entire application.
  ' Parameters:
  '   model (ByRef) - UIModelListItemTypeAndItem instance, or Nothing on first call.
  ' Returns:
  '   None (model is updated in-place).
  ' Notes:
  '   Path 1: model Is Nothing
  '       - Initialise model
  '       - Load all ListItemTypes
  '       - Clear ListItems
  '
  '   Path 2: model already exists
  '       - Reload ListItems for the currently selected ListItemType
  ' ==========================================================================================
  Friend Sub LoadListItemTypeAndItemModel(ByRef model As UIModelListItemTypeAndItem)

    ' === Path 1: First call – initialise model and load ListItemTypes ===
    If model Is Nothing Then
      model = New UIModelListItemTypeAndItem()

      model.ListItemTypes = LoadAllListItemTypes()
      model.ListItems = New SortableBindingList(Of UIListItemRow)()

      model.SelectedListItemType = Nothing
      model.ActionListItemType = New UIListItemTypeRow()
      model.ActionListItem = New UIListItemRow()

      Exit Sub
    End If

    ' === Path 2: Subsequent calls – refresh ListItems for current selection ===
    If model.SelectedListItemType Is Nothing Then
      model.ListItems.Clear()
    Else
      model.ListItems = LoadListItemsForType(model.SelectedListItemType.ListItemTypeID)
    End If

  End Sub

  ' ==========================================================================================
  ' Routine: LoadAllListItemTypes
  ' Purpose:
  '   Load all ListItemTypes from the database and map them to UIListItemTypeRow objects.
  ' Parameters:
  '   None
  ' Returns:
  '   SortableBindingList(Of UIListItemTypeRow)
  ' Notes:
  '   - Uses RecordLoader only.
  '   - System types and user-defined types are both returned.
  ' ==========================================================================================
  Private Function LoadAllListItemTypes() As SortableBindingList(Of UIListItemTypeRow)
    Dim conn As SQLiteConnectionWrapper = Nothing
    Dim list As New SortableBindingList(Of UIListItemTypeRow)

    Try
      conn = OpenDatabase(AddInContext.Current.Config.ActiveDbPath)

      Dim recs = RecordLoader.LoadRecords(Of RecordListItemType)(conn)
      For Each r In recs.OrderBy(Function(x) x.ListItemTypeName)
        list.Add(MapRecordToUIListItemType(r))
      Next

      Return list

    Catch ex As Exception
      ErrorHandler.UnHandleError(ex)
      Throw

    Finally
      If conn IsNot Nothing Then conn.Close()
    End Try
  End Function

  ' ==========================================================================================
  ' Routine: LoadListItemsForType
  ' Purpose:
  '   Load all ListItems belonging to a specific ListItemTypeID.
  ' Parameters:
  '   listItemTypeID - GUID of the ListItemType whose items should be loaded.
  ' Returns:
  '   SortableBindingList(Of UIListItemRow)
  ' Notes:
  '   - Uses RecordLoader only.
  ' ==========================================================================================
  Private Function LoadListItemsForType(listItemTypeID As String) As SortableBindingList(Of UIListItemRow)
    Dim conn As SQLiteConnectionWrapper = Nothing
    Dim list As New SortableBindingList(Of UIListItemRow)

    Try
      conn = OpenDatabase(AddInContext.Current.Config.ActiveDbPath)

      Dim recs = RecordLoader.LoadRecordsByFields(Of RecordListItem)(
        conn,
        {"ListItemTypeID"},
        {listItemTypeID}
      )

      For Each r In recs.OrderBy(Function(x) x.ListItemName)
        list.Add(MapRecordToUIListItem(r))
      Next

      Return list

    Catch ex As Exception
      ErrorHandler.UnHandleError(ex)
      Throw

    Finally
      If conn IsNot Nothing Then conn.Close()
    End Try
  End Function

  ' ==========================================================================================
  ' Routine: SaveListItemTypeAction
  ' Purpose:
  '   Persist Add / Update / Delete actions for ListItemTypes.
  ' Parameters:
  '   model (ByRef) - UIModelListItemTypeAndItem containing PendingListItemTypeAction
  '                   and ActionListItemType.
  ' Returns:
  '   None (model is refreshed after save).
  ' Notes:
  '   - System types (IsSystemType = True) cannot be deleted.
  '   - After saving, the entire model is reloaded (Path 1).
  ' ==========================================================================================
  Friend Sub SaveListItemTypeAction(ByRef model As UIModelListItemTypeAndItem)
    Dim conn As SQLiteConnectionWrapper = Nothing

    Try
      conn = OpenDatabase(AddInContext.Current.Config.ActiveDbPath)

      Select Case model.PendingListItemTypeAction

        Case ListItemTypeAction.Add
          Dim rec = MapNewTypeToRecord(model)
          RecordSaver.SaveRecord(conn, rec)

        Case ListItemTypeAction.Update
          Dim rec = MapUpdateTypeToRecord(model)
          RecordSaver.SaveRecord(conn, rec)

        Case ListItemTypeAction.Delete
          If model.ActionListItemType.IsSystemType Then
            Throw New InvalidOperationException("System ListItemTypes cannot be deleted.")
          End If

          Dim rec As New RecordListItemType With {
            .ListItemTypeID = model.ActionListItemType.ListItemTypeID,
            .IsDeleted = True
          }
          RecordSaver.SaveRecord(conn, rec)

      End Select

      ' === Reload entire model ===
      model = Nothing
      LoadListItemTypeAndItemModel(model)

    Catch ex As Exception
      ErrorHandler.UnHandleError(ex)
      Throw

    Finally
      If conn IsNot Nothing Then conn.Close()
    End Try
  End Sub

  ' ==========================================================================================
  ' Routine: SaveListItemAction
  ' Purpose:
  '   Persist Add / Update / Delete actions for ListItems.
  ' Parameters:
  '   model (ByRef) - UIModelListItemTypeAndItem containing PendingListItemAction
  '                   and ActionListItem.
  ' Returns:
  '   None (ListItems are refreshed after save).
  ' Notes:
  '   - ListItemTypeID is inherited from the currently selected ListItemType.
  ' ==========================================================================================
  Friend Sub SaveListItemAction(ByRef model As UIModelListItemTypeAndItem)
    Dim conn As SQLiteConnectionWrapper = Nothing

    Try
      conn = OpenDatabase(AddInContext.Current.Config.ActiveDbPath)

      Select Case model.PendingListItemAction

        Case ListItemAction.Add
          Dim rec = MapNewItemToRecord(model)
          RecordSaver.SaveRecord(conn, rec)

        Case ListItemAction.Update
          Dim rec = MapUpdateItemToRecord(model)
          RecordSaver.SaveRecord(conn, rec)

        Case ListItemAction.Delete
          Dim rec As New RecordListItem With {
            .ListItemID = model.ActionListItem.ListItemID,
            .IsDeleted = True
          }
          RecordSaver.SaveRecord(conn, rec)

      End Select

      ' === Refresh only the ListItems ===
      model.ListItems = LoadListItemsForType(model.SelectedListItemType.ListItemTypeID)

    Catch ex As Exception
      ErrorHandler.UnHandleError(ex)
      Throw

    Finally
      If conn IsNot Nothing Then conn.Close()
    End Try
  End Sub

  ' ==========================================================================================
  ' Routine: MapRecordToUIListItemType
  ' Purpose:
  '   Convert a RecordListItemType into a UIListItemTypeRow.
  ' Parameters:
  '   r - RecordListItemType
  ' Returns:
  '   UIListItemTypeRow
  ' ==========================================================================================
  Private Function MapRecordToUIListItemType(r As RecordListItemType) As UIListItemTypeRow
    Return New UIListItemTypeRow With {
      .ListItemTypeID = r.ListItemTypeID,
      .ListItemTypeName = r.ListItemTypeName,
      .IsSystemType = (r.IsSystemType = 1)
    }
  End Function

  ' ==========================================================================================
  ' Routine: MapRecordToUIListItem
  ' Purpose:
  '   Convert a RecordListItem into a UIListItemRow.
  ' Parameters:
  '   r - RecordListItem
  ' Returns:
  '   UIListItemRow
  ' ==========================================================================================
  Private Function MapRecordToUIListItem(r As RecordListItem) As UIListItemRow
    Return New UIListItemRow With {
      .ListItemID = r.ListItemID,
      .ListItemTypeID = r.ListItemTypeID,
      .ListItemName = r.ListItemName
    }
  End Function

  ' ==========================================================================================
  ' Routine: MapNewTypeToRecord
  ' Purpose:
  '   Build a new RecordListItemType from the UI model for insertion.
  ' Parameters:
  '   model - UIModelListItemTypeAndItem
  ' Returns:
  '   RecordListItemType (IsNew = True)
  ' ==========================================================================================
  Private Function MapNewTypeToRecord(model As UIModelListItemTypeAndItem) As RecordListItemType
    Dim r As New RecordListItemType()
    r.ListItemTypeID = Guid.NewGuid().ToString()
    r.ListItemTypeName = model.ActionListItemType.ListItemTypeName
    r.IsSystemType = 0
    r.IsNew = True
    Return r
  End Function

  ' ==========================================================================================
  ' Routine: MapUpdateTypeToRecord
  ' Purpose:
  '   Build a RecordListItemType for update from the UI model.
  ' Parameters:
  '   model - UIModelListItemTypeAndItem
  ' Returns:
  '   RecordListItemType (IsDirty = True)
  ' ==========================================================================================
  Private Function MapUpdateTypeToRecord(model As UIModelListItemTypeAndItem) As RecordListItemType
    Dim r As New RecordListItemType()
    r.ListItemTypeID = model.ActionListItemType.ListItemTypeID
    r.ListItemTypeName = model.ActionListItemType.ListItemTypeName
    r.IsSystemType = If(model.ActionListItemType.IsSystemType, 1, 0)
    r.IsDirty = True
    Return r
  End Function

  ' ==========================================================================================
  ' Routine: MapNewItemToRecord
  ' Purpose:
  '   Build a new RecordListItem from the UI model for insertion.
  ' Parameters:
  '   model - UIModelListItemTypeAndItem
  ' Returns:
  '   RecordListItem (IsNew = True)
  ' ==========================================================================================
  Private Function MapNewItemToRecord(model As UIModelListItemTypeAndItem) As RecordListItem
    Dim r As New RecordListItem()
    r.ListItemID = Guid.NewGuid().ToString()
    r.ListItemTypeID = model.SelectedListItemType.ListItemTypeID
    r.ListItemName = model.ActionListItem.ListItemName
    r.IsNew = True
    Return r
  End Function

  ' ==========================================================================================
  ' Routine: MapUpdateItemToRecord
  ' Purpose:
  '   Build a RecordListItem for update from the UI model.
  ' Parameters:
  '   model - UIModelListItemTypeAndItem
  ' Returns:
  '   RecordListItem (IsDirty = True)
  ' ==========================================================================================
  Private Function MapUpdateItemToRecord(model As UIModelListItemTypeAndItem) As RecordListItem
    Dim r As New RecordListItem()
    r.ListItemID = model.ActionListItem.ListItemID
    r.ListItemTypeID = model.ActionListItem.ListItemTypeID
    r.ListItemName = model.ActionListItem.ListItemName
    r.IsDirty = True
    Return r
  End Function

End Module
