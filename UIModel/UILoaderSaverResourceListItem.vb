Option Explicit On
Imports System.ComponentModel

' ==========================================================================================
' Module: UILoaderSaverResourceListItem
'
' Purpose:
'   Loader and saver logic for the Resource List Item maintenance form.
'   This module is the ONLY supported mechanism for:
'     - Loading ResourceListItems into UIModelResourceListItem
'     - Persisting Add / Update / Delete actions for ResourceListItems
'
' Responsibilities:
'   - Follow the two-path loader pattern:
'       Path 1: model Is Nothing → initialise model and load all data
'       Path 2: model exists → refresh ResourceListItems only
'   - Use ONLY RecordLoader and RecordSaver (no direct SQL)
'   - Map between RecordResourceListItem and UIResourceListItemRow
'   - Provide explicit "None" option for ListItemType selection
'
' Notes:
'   - All ResourceListItems use GUID primary keys.
'   - ListItemTypeID may be "" (None) when the item is not tied to a dropdown.
'   - This module must remain deterministic and drift-free.
' ==========================================================================================
Module UILoaderSaverResourceListItem

  ' ==========================================================================================
  ' Routine: LoadResourceListItemModel
  ' Purpose:
  '   Load or refresh the UIModelResourceListItem following the standard two-path loader
  '   pattern used across the entire application.
  ' Parameters:
  '   model (ByRef) - UIModelResourceListItem instance, or Nothing on first call.
  ' Returns:
  '   None (model is updated in-place).
  ' Notes:
  '   Path 1: model Is Nothing
  '       - Initialise model
  '       - Load all ResourceListItems
  '       - Load ListItemTypes lookup (including explicit "None")
  '
  '   Path 2: model already exists
  '       - Reload ResourceListItems only
  ' ==========================================================================================
  Friend Sub LoadResourceListItemModel(ByRef model As UIModelResourceListItem)
    Dim isFirstLoad As Boolean = (model Is Nothing)

    If isFirstLoad Then
      model = New UIModelResourceListItem()

      model.ListItemTypes = LoadAllListItemTypesForResourceListItems()
      model.ResourceListItems = LoadAllResourceListItems()

      model.PendingResourceListItemAction = ResourceListItemAction.None
      model.ActionResourceListItem = New UIResourceListItemRow()

      Exit Sub
    End If

    ' Path 2: refresh only the ResourceListItems
    model.ResourceListItems = LoadAllResourceListItems()

  End Sub

  ' ==========================================================================================
  ' Routine: LoadAllResourceListItems
  ' Purpose:
  '   Load all ResourceListItems from the database and map them to UIResourceListItemRow objects.
  ' Parameters:
  '   None
  ' Returns:
  '   SortableBindingList(Of UIResourceListItemRow)
  ' ==========================================================================================
  Private Function LoadAllResourceListItems() As SortableBindingList(Of UIResourceListItemRow)
    Dim conn As SQLiteConnectionWrapper = Nothing
    Dim list As New SortableBindingList(Of UIResourceListItemRow)

    Try
      conn = OpenDatabase(AddInContext.Current.Config.ActiveDbPath)

      Dim recs = RecordLoader.LoadRecords(Of RecordResourceListItem)(conn)

      For Each r In recs.OrderBy(Function(x) x.ResourceListItemName)
        list.Add(MapRecordToUIResourceListItemRow(r))
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
  ' Routine: LoadAllListItemTypesForResourceListItems
  ' Purpose:
  '   Load all ListItemTypes for use as a lookup when assigning ResourceListItems.
  '   Adds an explicit "None" option so ResourceListItems are not required to bind to a
  '   specific ListItemType.
  ' Parameters:
  '   None
  ' Returns:
  '   SortableBindingList(Of UIListItemTypeRow)
  ' ==========================================================================================
  Private Function LoadAllListItemTypesForResourceListItems() As SortableBindingList(Of UIListItemTypeRow)
    Dim conn As SQLiteConnectionWrapper = Nothing
    Dim list As New SortableBindingList(Of UIListItemTypeRow)

    Try
      ' Explicit "None" option
      list.Add(New UIListItemTypeRow With {
        .ListItemTypeID = "",
        .ListItemTypeName = "None",
        .IsSystemType = False
      })

      conn = OpenDatabase(AddInContext.Current.Config.ActiveDbPath)

      Dim recs = RecordLoader.LoadRecords(Of RecordListItemType)(conn)

      For Each r In recs.OrderBy(Function(x) x.ListItemTypeName)
        list.Add(New UIListItemTypeRow With {
          .ListItemTypeID = r.ListItemTypeID,
          .ListItemTypeName = r.ListItemTypeName,
          .IsSystemType = (r.IsSystemType = 1)
        })
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
  ' Routine: SaveResourceListItemAction
  ' Purpose:
  '   Persist Add / Update / Delete actions for ResourceListItems.
  ' Parameters:
  '   model (ByRef) - UIModelResourceListItem containing PendingResourceListItemAction
  '                   and ActionResourceListItem.
  ' Returns:
  '   None (ResourceListItems are refreshed after save).
  ' Notes:
  '   - New ResourceListItems receive a GUID primary key.
  '   - ListItemTypeID may be "" to indicate "None" (not tied to any ListItemType).
  ' ==========================================================================================
  Friend Sub SaveResourceListItemAction(ByRef model As UIModelResourceListItem)
    Dim conn As SQLiteConnectionWrapper = Nothing

    Try
      conn = OpenDatabase(AddInContext.Current.Config.ActiveDbPath)

      Select Case model.PendingResourceListItemAction

        Case ResourceListItemAction.Add
          Dim rec = MapNewResourceListItemToRecord(model)
          RecordSaver.SaveRecord(conn, rec)

        Case ResourceListItemAction.Update
          Dim rec = MapUpdateResourceListItemToRecord(model)
          RecordSaver.SaveRecord(conn, rec)

        Case ResourceListItemAction.Delete
          Dim rec As New RecordResourceListItem With {
            .ResourceListItemID = model.ActionResourceListItem.ResourceListItemID,
            .IsDeleted = True
          }
          RecordSaver.SaveRecord(conn, rec)

      End Select

      ' Refresh ResourceListItems
      model.ResourceListItems = LoadAllResourceListItems()
      model.PendingResourceListItemAction = ResourceListItemAction.None

    Catch ex As Exception
      ErrorHandler.UnHandleError(ex)
      Throw

    Finally
      If conn IsNot Nothing Then conn.Close()
    End Try
  End Sub

  ' ==========================================================================================
  ' Routine: MapRecordToUIResourceListItemRow
  ' Purpose:
  '   Convert a RecordResourceListItem into a UIResourceListItemRow.
  ' Parameters:
  '   r - RecordResourceListItem
  ' Returns:
  '   UIResourceListItemRow
  ' ==========================================================================================
  Private Function MapRecordToUIResourceListItemRow(r As RecordResourceListItem) As UIResourceListItemRow
    Return New UIResourceListItemRow With {
      .ResourceListItemID = r.ResourceListItemID,
      .ResourceListItemName = r.ResourceListItemName,
      .ListItemTypeID = r.ListItemTypeID,
      .ValueType =
          If(String.IsNullOrWhiteSpace(r.ValueType),
             ResourceListItemValueType.Text,
             CType([Enum].Parse(GetType(ResourceListItemValueType), r.ValueType), ResourceListItemValueType))}
  End Function

  ' ==========================================================================================
  ' Routine: MapNewResourceListItemToRecord
  ' Purpose:
  '   Build a new RecordResourceListItem from the UI model for insertion.
  ' Parameters:
  '   model - UIModelResourceListItem
  ' Returns:
  '   RecordResourceListItem (IsNew = True)
  ' ==========================================================================================
  Private Function MapNewResourceListItemToRecord(model As UIModelResourceListItem) As RecordResourceListItem
    Dim r As New RecordResourceListItem()

    r.ResourceListItemID = Guid.NewGuid().ToString()
    r.ResourceListItemName = model.ActionResourceListItem.ResourceListItemName
    r.ListItemTypeID = model.ActionResourceListItem.ListItemTypeID   ' "" allowed for None
    r.ValueType = model.ActionResourceListItem.ValueType.ToString()
    r.IsNew = True

    Return r
  End Function

  ' ==========================================================================================
  ' Routine: MapUpdateResourceListItemToRecord
  ' Purpose:
  '   Build a RecordResourceListItem for update from the UI model.
  ' Parameters:
  '   model - UIModelResourceListItem
  ' Returns:
  '   RecordResourceListItem (IsDirty = True)
  ' ==========================================================================================
  Private Function MapUpdateResourceListItemToRecord(model As UIModelResourceListItem) As RecordResourceListItem
    Dim r As New RecordResourceListItem()

    r.ResourceListItemID = model.ActionResourceListItem.ResourceListItemID
    r.ResourceListItemName = model.ActionResourceListItem.ResourceListItemName
    r.ListItemTypeID = model.ActionResourceListItem.ListItemTypeID   ' "" allowed for None
    r.ValueType = model.ActionResourceListItem.ValueType.ToString()
    r.IsDirty = True

    Return r
  End Function

End Module