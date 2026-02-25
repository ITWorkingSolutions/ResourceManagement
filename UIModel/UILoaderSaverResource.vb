Option Explicit On

' ==========================================================================================
' Module: UILoaderSaverResource
' Purpose:
'   Load and save a single Resource record (and its Role/Function assignments and
'   Resource Name/Value attributes) using the UIModelResources contract.
'
'   - NO direct SQL. All database interaction is via:
'       * OpenDatabase
'       * RecordLoader
'       * RecordSaver
'
'   - Supports:
'       * Add (no ResourceID passed)
'       * Update (existing ResourceID)
'       * Delete (PendingAction = Delete)
'
'   - Role/Function assignments use per-row PendingAction:
'       * ResourceRoleFunctionAction.None   -> no-op
'       * ResourceRoleFunctionAction.Add    -> insert row
'       * ResourceRoleFunctionAction.Update -> update PrimaryRole only
'       * ResourceRoleFunctionAction.Delete -> delete row
'
'   - Resource Name/Value assignments use snapshot vs working copy diff:
'       * model.ResourceNameValues           = DB snapshot
'       * model.ActionResourceNameValues     = working copy
'       * Saver computes Adds / Updates / Deletes by comparing these sets
'
'   - Lookup lists (Salutation, Gender, Role/Function) are loaded via
'     UILoaderSaverListItemTypeAndItem and exposed as UIListItemRow collections.
' ==========================================================================================
Friend Module UILoaderSaverResource
  ' ==========================================================================================
  ' Routine: LoadResourceModel
  ' Purpose:
  '   Load an existing resource (and its role/function assignments and name/value attributes),
  '   OR return an initialised empty model when no ResourceID is supplied.
  '
  ' Parameters:
  '   resourceID - "" means return an initialised empty model.
  '
  ' Returns:
  '   UIModelResources - containing:
  '       - Resource (snapshot)
  '       - ActionResource (working copy)
  '       - ResourceNameValues (snapshot)
  '       - ActionResourceNameValues (working copy)
  '       - Lookup lists (UIListItemRow)
  '
  ' Notes:
  '   - NO PendingAction set here.
  '   - Loader/Saver does not infer Add/Update/Delete.
  '   - Cloning is done here so UI has a safe working copy.
  ' ==========================================================================================
  Friend Function LoadResourceModel(ByVal resourceID As String) As UIModelResources

    Dim conn As SQLiteConnectionWrapper = Nothing
    Dim model As New UIModelResources()
    Dim rec As RecordResource

    Try
      conn = OpenDatabase(AddInContext.Current.Config.ActiveDbPath)

      ' === Lookup lists (UIListItemRow only, via LoadListItemTypeAndItemModel) ===
      Dim liModel As UIModelListItemTypeAndItem = Nothing
      UILoaderSaverListItemTypeAndItem.LoadListItemTypeAndItemModel(liModel)

      ' Salutation
      liModel.SelectedListItemType =
        liModel.ListItemTypes.First(Function(t) t.ListItemTypeID = ListItemTypeSystemCatalog.SalutationTypeID)
      UILoaderSaverListItemTypeAndItem.LoadListItemTypeAndItemModel(liModel)
      model.Salutations = New List(Of UIListItemRow)(liModel.ListItems)
      ' Get the type name for labeling
      model.SalutationListItemTypeName = liModel.SelectedListItemType.ListItemTypeName

      ' Gender
      liModel.SelectedListItemType =
        liModel.ListItemTypes.First(Function(t) t.ListItemTypeID = ListItemTypeSystemCatalog.GenderTypeID)
      UILoaderSaverListItemTypeAndItem.LoadListItemTypeAndItemModel(liModel)
      model.Genders = New List(Of UIListItemRow)(liModel.ListItems)
      ' Get the type name for labeling
      model.GenderListItemTypeName = liModel.SelectedListItemType.ListItemTypeName

      If String.IsNullOrEmpty(resourceID) Then

        ' --------------------------------------------------------------
        '  No ResourceID: return an initialised empty model
        '  - Resource: empty snapshot
        '  - ActionResource: empty working copy
        '  - ResourceRoleFunctions: empty snapshot
        '  - ActionResourceRoleFunctions: empty working set
        '  - ResourceNameValues: empty snapshot
        '  - ActionResourceNameValues: empty working set
        ' --------------------------------------------------------------
        model.Resource = New UIResourceRow()
        model.ActionResource = CloneResourceRow(model.Resource)

        ' === Resource Name/Value definitions only (no values yet) ===
        Dim defs As List(Of UIResourceNameValueRow)
        defs = LoadResourceNameValues(conn, "")   ' definitions only

        model.ResourceNameValues = New SortableBindingList(Of UIResourceNameValueRow)()
        model.ActionResourceNameValues = New SortableBindingList(Of UIResourceNameValueRow)()

        Dim d As UIResourceNameValueRow
        For Each d In defs
          model.ActionResourceNameValues.Add(d)
        Next

        Return model

      End If

      ' --------------------------------------------------------------
      '  ResourceID supplied: load existing resource + assignments
      ' --------------------------------------------------------------
      Dim pkValues() As Object = {resourceID}
      rec = RecordLoader.LoadRecord(Of RecordResource)(conn, pkValues)

      If rec Is Nothing Then
        Throw New InvalidOperationException($"ResourceID '{resourceID}' not found.")
      End If

      ' Snapshot resource from DB
      model.Resource = MapRecordToUIResourceRow(rec)

      ' Working copy for UI edits
      model.ActionResource = CloneResourceRow(model.Resource)

      ' Snapshot resource name/value attributes
      Dim nameValueSnapshot As List(Of UIResourceNameValueRow)
      nameValueSnapshot = LoadResourceNameValues(conn, resourceID)

      model.ResourceNameValues = New SortableBindingList(Of UIResourceNameValueRow)()
      Dim nv As UIResourceNameValueRow
      For Each nv In nameValueSnapshot
        model.ResourceNameValues.Add(nv)
      Next

      ' Working set for name/value UI edits
      model.ActionResourceNameValues = New SortableBindingList(Of UIResourceNameValueRow)()
      For Each nv In model.ResourceNameValues
        model.ActionResourceNameValues.Add(CloneResourceNameValueRow(nv))
      Next

      Return model

    Catch ex As Exception
      ErrorHandler.UnHandleError(ex)
      Throw
    Finally
      If conn IsNot Nothing Then conn.Close()
    End Try

  End Function

  ' ==========================================================================================
  ' Routine: SavePendingResourceAction
  ' Purpose:
  '   Perform the pending resource-level action encoded in the UI model (Add / Update / Delete)
  '   and commit all ActionResourceRoleFunctions and ActionResourceNameValues.
  ' Parameters:
  '   model - UIModelResources containing PendingAction, ActionResource,
  '           ActionResourceRoleFunctions, and ActionResourceNameValues.
  ' Returns:
  '   None
  ' Notes:
  '   - Does NOT reload or replace the model.
  '   - Caller is responsible for any post-save refresh (e.g. reloading grids or forms).
  ' ==========================================================================================
  Friend Sub SavePendingResourceAction(ByRef model As UIModelResources)

    Dim conn As SQLiteConnectionWrapper = Nothing
    Dim newRec As RecordResource
    Dim updRec As RecordResource
    Dim delRec As RecordResource
    Dim resourceID As String

    If model Is Nothing Then Throw New ArgumentNullException(NameOf(model))

    Try
      conn = OpenDatabase(AddInContext.Current.Config.ActiveDbPath)

      Select Case model.PendingAction

        Case ResourceAction.Add
          ' === Insert new resource ===
          newRec = MapNewValuesToRecord(model)
          RecordSaver.SaveRecord(conn, newRec)
          resourceID = newRec.ResourceID

          ' Keep the new ResourceID on the UI model for caller use
          model.ActionResource.ResourceID = resourceID

          SaveResourceNameValues(conn, resourceID, model.ActionResourceNameValues)

        Case ResourceAction.Update
          ' === Update existing resource ===
          If model.ActionResource Is Nothing OrElse
             String.IsNullOrEmpty(model.ActionResource.ResourceID) Then
            Throw New InvalidOperationException("No resource selected for update.")
          End If

          updRec = MapUpdateValuesToRecord(model)
          RecordSaver.SaveRecord(conn, updRec)
          resourceID = updRec.ResourceID

          SaveResourceNameValues(conn, resourceID, model.ActionResourceNameValues)

        Case ResourceAction.Delete
          ' === Soft-delete resource and its assignments ===
          If model.ActionResource Is Nothing OrElse
             String.IsNullOrEmpty(model.ActionResource.ResourceID) Then
            Throw New InvalidOperationException("No resource selected for delete.")
          End If

          resourceID = model.ActionResource.ResourceID

          delRec = New RecordResource()
          delRec.ResourceID = resourceID
          delRec.IsDeleted = True
          RecordSaver.SaveRecord(conn, delRec)

          DeleteAllResourceNameValues(conn, resourceID)

        Case Else
          ' === Nothing to do ===
          Return

      End Select

    Catch ex As Exception
      ErrorHandler.UnHandleError(ex)
      Throw
    Finally
      If conn IsNot Nothing Then conn.Close()
    End Try

  End Sub

  ' ==========================================================================================
  ' Routine: MapRecordToUIResourceRow
  ' Purpose:
  '   Map a RecordResource domain object to a UIResourceRow for display/edit.
  ' Parameters:
  '   r - RecordResource loaded from DB.
  ' Returns:
  '   UIResourceRow.
  ' Notes:
  '   - UIResourceRow contains IDs only. No lookup names are mapped.
  '   - Dates are parsed from "yyyy-MM-dd" strings into Date.
  ' ==========================================================================================
  Private Function MapRecordToUIResourceRow(ByVal r As RecordResource) As UIResourceRow

    Dim row As New UIResourceRow()

    ' Identity & naming
    row.ResourceID = r.ResourceID
    row.PreferredName = r.PreferredName
    row.SalutationID = r.SalutationID
    row.GenderID = r.GenderID
    row.FirstName = r.FirstName
    row.LastName = r.LastName

    ' Contact
    row.Email = r.Email
    row.Phone = r.Phone

    ' Lifecycle
    If Not String.IsNullOrEmpty(r.StartDate) Then
      row.StartDate = Date.ParseExact(r.StartDate, "yyyy-MM-dd", Nothing)
    End If

    If Not String.IsNullOrEmpty(r.EndDate) Then
      row.EndDate = Date.ParseExact(r.EndDate, "yyyy-MM-dd", Nothing)
    End If

    ' Meta
    row.Notes = r.Notes

    Return row

  End Function

  ' ==========================================================================================
  ' Routine: MapUpdateValuesToRecord
  ' Purpose:
  '   Build a RecordResource for update from UIModelResources.ActionResource.
  ' Parameters:
  '   model - UIModelResources containing ActionResource with updated values.
  ' Returns:
  '   RecordResource ready for RecordSaver.SaveRecord (IsDirty = True).
  ' Notes:
  '   - Preserves ResourceID from ActionResource.
  '   - Dates stored as "yyyy-MM-dd" strings; blank when Date.MinValue.
  ' ==========================================================================================
  Private Function MapUpdateValuesToRecord(ByVal model As UIModelResources) As RecordResource

    Dim r As New RecordResource()
    Dim a As UIResourceRow = model.ActionResource

    ' Primary key
    r.ResourceID = a.ResourceID
    r.IsDirty = True

    ' Identity & naming
    r.PreferredName = a.PreferredName
    r.SalutationID = a.SalutationID
    r.GenderID = a.GenderID
    r.FirstName = a.FirstName
    r.LastName = a.LastName

    ' Contact
    r.Email = a.Email
    r.Phone = a.Phone

    ' Lifecycle
    r.StartDate = If(a.StartDate = Date.MinValue, "", a.StartDate.ToString("yyyy-MM-dd"))
    r.EndDate = If(a.EndDate = Date.MinValue, "", a.EndDate.ToString("yyyy-MM-dd"))

    ' Meta
    r.Notes = a.Notes

    Return r

  End Function

  ' ==========================================================================================
  ' Routine: MapNewValuesToRecord
  ' Purpose:
  '   Build a RecordResource for insertion from UIModelResources.ActionResource.
  ' Parameters:
  '   model - UIModelResources containing ActionResource with new values.
  ' Returns:
  '   RecordResource ready for RecordSaver.SaveRecord (IsNew = True).
  ' Notes:
  '   - Generates a GUID for ResourceID.
  '   - Dates stored as "yyyy-MM-dd" strings; blank when Date.MinValue.
  ' ==========================================================================================
  Private Function MapNewValuesToRecord(ByVal model As UIModelResources) As RecordResource

    Dim r As New RecordResource()
    Dim a As UIResourceRow = model.ActionResource

    ' Primary key
    r.ResourceID = Guid.NewGuid().ToString()
    r.IsNew = True

    ' Identity & naming
    r.PreferredName = a.PreferredName
    r.SalutationID = a.SalutationID
    r.GenderID = a.GenderID
    r.FirstName = a.FirstName
    r.LastName = a.LastName

    ' Contact
    r.Email = a.Email
    r.Phone = a.Phone

    ' Lifecycle
    r.StartDate = If(a.StartDate = Date.MinValue, "", a.StartDate.ToString("yyyy-MM-dd"))
    r.EndDate = If(a.EndDate = Date.MinValue, "", a.EndDate.ToString("yyyy-MM-dd"))

    ' Meta
    r.Notes = a.Notes

    Return r

  End Function

  ' ==========================================================================================
  ' Routine: LoadResourceNameValues
  ' Purpose:
  '   Load ALL ResourceListItem definitions and merge any existing values
  '   from tblResourceNameValue and tblResourceNameValueListItem for the given ResourceID.
  '
  '   Supports:
  '     - Text
  '     - SingleSelectList
  '     - MultiSelectList
  '
  ' Parameters:
  '   conn       - SQLite connection wrapper.
  '   resourceID - ResourceID to load attributes for. "" means new resource.
  '
  ' Returns:
  '   List(Of UIResourceNameValueRow) complete working set.
  '
  ' Notes:
  '   - Ensures every ResourceListItem appears even if no value exists yet.
  '   - Definitions come from tblResourceListItem.
  '   - Single/text values come from tblResourceNameValue.
  '   - Multi-select values come from tblResourceNameValueListItem.
  ' ==========================================================================================
  Private Function LoadResourceNameValues(ByVal conn As SQLiteConnectionWrapper,
                                        ByVal resourceID As String) _
                                        As List(Of UIResourceNameValueRow)

    Dim results As New List(Of UIResourceNameValueRow)()

    ' 1. Load ALL definitions from tblResourceListItem
    Dim definitions As List(Of RecordResourceListItem)
    definitions = RecordLoader.LoadRecords(Of RecordResourceListItem)(conn)

    Dim defList As New List(Of RecordResourceListItem)()
    Dim def As RecordResourceListItem
    For Each def In definitions
      If def IsNot Nothing AndAlso Not def.IsDeleted AndAlso
         Not String.IsNullOrEmpty(def.ResourceListItemID) Then
        defList.Add(def)
      End If
    Next

    ' 2. Load existing single/text values when resourceID is supplied
    Dim singleValuesByID As New Dictionary(Of String, RecordResourceNameValue)(StringComparer.OrdinalIgnoreCase)

    If Not String.IsNullOrEmpty(resourceID) Then

      Dim singleValueRows As List(Of RecordResourceNameValue)
      singleValueRows = RecordLoader.LoadRecordsByFields(Of RecordResourceNameValue)(
                    conn,
                    {"ResourceID"},
                    {resourceID}
                )

      Dim svr As RecordResourceNameValue
      For Each svr In singleValueRows
        If svr IsNot Nothing AndAlso Not svr.IsDeleted AndAlso
             Not String.IsNullOrEmpty(svr.ResourceListItemID) Then

          If Not singleValuesByID.ContainsKey(svr.ResourceListItemID) Then
            singleValuesByID.Add(svr.ResourceListItemID, svr)
          End If
        End If
      Next

    End If

    ' 3. Load existing multi values when resourceID is supplied
    Dim multiValuesByID As New Dictionary(Of String, RecordResourceNameValueListItem)(StringComparer.OrdinalIgnoreCase)

    If Not String.IsNullOrEmpty(resourceID) Then

      Dim multiValueRows As List(Of RecordResourceNameValueListItem)
      multiValueRows = RecordLoader.LoadRecordsByFields(Of RecordResourceNameValueListItem)(
                    conn,
                    {"ResourceID"},
                    {resourceID}
                )

      Dim mvr As RecordResourceNameValueListItem
      For Each mvr In multiValueRows
        If mvr IsNot Nothing AndAlso Not mvr.IsDeleted AndAlso
             Not String.IsNullOrEmpty(mvr.ResourceListItemID) Then

          If Not multiValuesByID.ContainsKey(mvr.ResourceListItemID) Then
            multiValuesByID.Add(mvr.ResourceListItemID, mvr)
          End If
        End If
      Next

    End If

    ' 4. Merge definitions + optional values into UIResourceNameValueRow objects
    Dim d As RecordResourceListItem
    For Each d In defList

      Dim row As New UIResourceNameValueRow()

      ' --- Definition fields ---
      row.ResourceListItemID = d.ResourceListItemID
      row.ResourceListItemName = d.ResourceListItemName
      row.ListItemTypeID = d.ListItemTypeID
      row.ValueType = If(String.IsNullOrWhiteSpace(d.ValueType),
        ResourceListItemValueType.Text,
        CType([Enum].Parse(GetType(ResourceListItemValueType), d.ValueType), ResourceListItemValueType))

      ' --- Load lookup items if applicable ---
      If row.ValueType = ResourceListItemValueType.SingleSelectList OrElse
       row.ValueType = ResourceListItemValueType.MultiSelectList Then

        row.ListItems = LoadListItemsForType(conn, d.ListItemTypeID)
      Else
        row.ListItems = New List(Of UIListItemRow)()
      End If

      ' --- Apply existing value if present  ---
      Dim existingSingleton As RecordResourceNameValue = Nothing
      Dim existingMulti As RecordResourceNameValueListItem = Nothing

      If singleValuesByID.TryGetValue(d.ResourceListItemID, existingSingleton) Then

        Select Case row.ValueType
          Case ResourceListItemValueType.Text
            row.ResourceListItemValue = existingSingleton.ResourceListItemValue
          Case ResourceListItemValueType.SingleSelectList
            row.SelectedListItemID = existingSingleton.ListItemID
        End Select

      ElseIf multiValuesByID.TryGetValue(d.ResourceListItemID, existingMulti) Then
        Select Case row.ValueType
          Case ResourceListItemValueType.MultiSelectList
            row.SelectedListItemIDs =
              LoadMultiSelectValues(conn, resourceID, d.ResourceListItemID)
        End Select
      Else
        ' No existing value
        row.ResourceListItemValue = Nothing
        row.SelectedListItemID = Nothing
        row.SelectedListItemIDs = New List(Of String)()
      End If

      results.Add(row)

    Next

    Return results

  End Function

  ' ==========================================================================================
  ' Routine: LoadListItemsForType
  ' Purpose:
  '   Load all ListItems belonging to a specific ListItemTypeID.
  ' Parameters:
  '   conn           - SQLite connection wrapper.
  '   listItemTypeID - GUID of the ListItemType whose items should be loaded.
  ' Returns:
  '   List(Of UIListItemRow)
  ' Notes:
  '   - Local copy of the pattern used in UILoaderSaverListItemTypeAndItem.
  ' ==========================================================================================
  Private Function LoadListItemsForType(ByVal conn As SQLiteConnectionWrapper,
                                        ByVal listItemTypeID As String) _
                                        As List(Of UIListItemRow)

    Dim list As New List(Of UIListItemRow)()

    If String.IsNullOrEmpty(listItemTypeID) Then Return list

    Dim recs = RecordLoader.LoadRecordsByFields(Of RecordListItem)(
      conn,
      {"ListItemTypeID"},
      {listItemTypeID}
    )

    If recs Is Nothing Then Return list

    Dim r As RecordListItem
    For Each r In recs.OrderBy(Function(x) x.ListItemName)
      Dim ui As New UIListItemRow()
      ui.ListItemID = r.ListItemID
      ui.ListItemTypeID = r.ListItemTypeID
      ui.ListItemName = r.ListItemName
      list.Add(ui)
    Next

    Return list

  End Function

  ' ==========================================================================================
  ' Routine: SaveResourceNameValues
  ' Purpose:
  '   Commit per-row PendingAction for ActionResourceNameValues:
  '     - Add    -> insert new values
  '     - Update -> update values
  '     - Delete -> soft-delete values
  ' Parameters:
  '   conn       - SQLite connection wrapper.
  '   resourceID - ResourceID for the attributes.
  '   actionRows - IEnumerable(Of UIResourceNameValueRow) with PendingAction set.
  ' Returns:
  '   None
  ' Notes:
  '   - Key for single/text is (ResourceID, ResourceListItemID) in tblResourceNameValue.
  '   - Multi-select values are stored in tblResourceNameValueListItem.
  ' ==========================================================================================
  Private Sub SaveResourceNameValues(ByVal conn As SQLiteConnectionWrapper,
                                   ByVal resourceID As String,
                                   ByVal actionRows As IEnumerable(Of UIResourceNameValueRow))

    If String.IsNullOrEmpty(resourceID) Then Exit Sub
    If actionRows Is Nothing Then Exit Sub

    Dim row As UIResourceNameValueRow

    For Each row In actionRows

      Select Case row.PendingAction

        Case ResourceNameValueAction.Add
          SaveNameValue_Add(conn, resourceID, row)

        Case ResourceNameValueAction.Update
          SaveNameValue_Update(conn, resourceID, row)

        Case ResourceNameValueAction.Delete
          SaveNameValue_Delete(conn, resourceID, row)

        Case ResourceNameValueAction.None
          ' no-op

      End Select

    Next

  End Sub

  ' ==========================================================================================
  ' Routine: SaveNameValue_Add
  ' Purpose:
  '   Insert new Resource Name/Value data for a single UIResourceNameValueRow.
  ' Parameters:
  '   conn       - SQLite connection wrapper.
  '   resourceID - ResourceID for the attributes.
  '   row        - UIResourceNameValueRow with PendingAction = Add.
  ' Returns:
  '   None
  ' Notes:
  '   - Text and SingleSelectList use tblResourceNameValue.
  '   - MultiSelectList uses tblResourceNameValueListItem (one row per ListItemID).
  ' ==========================================================================================
  Private Sub SaveNameValue_Add(ByVal conn As SQLiteConnectionWrapper,
                              ByVal resourceID As String,
                              ByVal row As UIResourceNameValueRow)

    Select Case row.ValueType

      Case ResourceListItemValueType.Text
        Dim rec As New RecordResourceNameValue()
        rec.ResourceID = resourceID
        rec.ResourceListItemID = row.ResourceListItemID
        rec.ResourceListItemValue = row.ResourceListItemValue
        rec.IsNew = True
        RecordSaver.SaveRecord(conn, rec)

      Case ResourceListItemValueType.SingleSelectList
        Dim rec As New RecordResourceNameValue()
        rec.ResourceID = resourceID
        rec.ResourceListItemID = row.ResourceListItemID
        rec.ListItemID = row.SelectedListItemID
        rec.IsNew = True
        RecordSaver.SaveRecord(conn, rec)

      Case ResourceListItemValueType.MultiSelectList
        If row.SelectedListItemIDs Is Nothing Then Exit Sub

        Dim id As String
        For Each id In row.SelectedListItemIDs
          Dim rec As New RecordResourceNameValueListItem()
          rec.ResourceID = resourceID
          rec.ResourceListItemID = row.ResourceListItemID
          rec.ListItemID = id
          rec.IsNew = True
          RecordSaver.SaveRecord(conn, rec)
        Next

    End Select

  End Sub

  ' ==========================================================================================
  ' Routine: SaveNameValue_Update
  ' Purpose:
  '   Update existing Resource Name/Value data for a single UIResourceNameValueRow.
  ' Parameters:
  '   conn       - SQLite connection wrapper.
  '   resourceID - ResourceID for the attributes.
  '   row        - UIResourceNameValueRow with PendingAction = Update.
  ' Returns:
  '   None
  ' Notes:
  '   - Text and SingleSelectList update tblResourceNameValue via LoadResourceNameValueRecord.
  '   - MultiSelectList deletes and reinserts rows in tblResourceNameValueListItem.
  ' ==========================================================================================
  Private Sub SaveNameValue_Update(ByVal conn As SQLiteConnectionWrapper,
                                 ByVal resourceID As String,
                                 ByVal row As UIResourceNameValueRow)

    Select Case row.ValueType

      Case ResourceListItemValueType.Text
        Dim rec As RecordResourceNameValue =
        LoadResourceNameValueRecord(conn, resourceID, row.ResourceListItemID)
        If rec Is Nothing Then Exit Sub
        rec.ResourceListItemValue = row.ResourceListItemValue
        rec.IsDirty = True
        RecordSaver.SaveRecord(conn, rec)

      Case ResourceListItemValueType.SingleSelectList
        Dim rec As RecordResourceNameValue =
        LoadResourceNameValueRecord(conn, resourceID, row.ResourceListItemID)
        If rec Is Nothing Then Exit Sub
        rec.ListItemID = row.SelectedListItemID
        rec.IsDirty = True
        RecordSaver.SaveRecord(conn, rec)

      Case ResourceListItemValueType.MultiSelectList
        ' Replace all existing multi-select rows with the new selection set.
        DeleteMultiSelectValues(conn, resourceID, row.ResourceListItemID)

        If row.SelectedListItemIDs Is Nothing Then Exit Sub

        Dim id As String
        For Each id In row.SelectedListItemIDs
          Dim rec As New RecordResourceNameValueListItem()
          rec.ResourceID = resourceID
          rec.ResourceListItemID = row.ResourceListItemID
          rec.ListItemID = id
          rec.IsNew = True
          RecordSaver.SaveRecord(conn, rec)
        Next

    End Select

  End Sub

  ' ==========================================================================================
  ' Routine: SaveNameValue_Delete
  ' Purpose:
  '   Soft-delete Resource Name/Value data for a single UIResourceNameValueRow.
  ' Parameters:
  '   conn       - SQLite connection wrapper.
  '   resourceID - ResourceID for the attributes.
  '   row        - UIResourceNameValueRow with PendingAction = Delete.
  ' Returns:
  '   None
  ' Notes:
  '   - Text and SingleSelectList mark tblResourceNameValue row as deleted.
  '   - MultiSelectList marks all tblResourceNameValueListItem rows as deleted.
  ' ==========================================================================================
  Private Sub SaveNameValue_Delete(ByVal conn As SQLiteConnectionWrapper,
                                 ByVal resourceID As String,
                                 ByVal row As UIResourceNameValueRow)

    Select Case row.ValueType

      Case ResourceListItemValueType.MultiSelectList
        DeleteMultiSelectValues(conn, resourceID, row.ResourceListItemID)

      Case Else
        Dim rec As RecordResourceNameValue =
        LoadResourceNameValueRecord(conn, resourceID, row.ResourceListItemID)
        If rec Is Nothing Then Exit Sub
        rec.IsDeleted = True
        RecordSaver.SaveRecord(conn, rec)

    End Select

  End Sub

  ' ==========================================================================================
  ' Routine: LoadResourceNameValueRecord
  ' Purpose:
  '   Load a single RecordResourceNameValue by (ResourceID, ResourceListItemID).
  ' Parameters:
  '   conn               - SQLite connection wrapper.
  '   resourceID         - ResourceID.
  '   resourceListItemID - ResourceListItemID.
  ' Returns:
  '   RecordResourceNameValue or Nothing if not found.
  ' Notes:
  '   - Uses RecordLoader.LoadRecordsByFields only.
  ' ==========================================================================================
  Private Function LoadResourceNameValueRecord(ByVal conn As SQLiteConnectionWrapper,
                                               ByVal resourceID As String,
                                               ByVal resourceListItemID As String) _
                                               As RecordResourceNameValue

    Dim fieldNames() As String = {"ResourceID", "ResourceListItemID"}
    Dim fieldValues() As Object = {resourceID, resourceListItemID}
    Dim list As List(Of RecordResourceNameValue)

    list = RecordLoader.LoadRecordsByFields(Of RecordResourceNameValue)(
             conn, fieldNames, fieldValues)

    If list Is Nothing OrElse list.Count = 0 Then Return Nothing
    Return list(0)

  End Function

  ' ==========================================================================================
  ' Routine: DeleteAllResourceNameValues
  ' Purpose:
  '   Soft-delete all Resource Name/Value attributes for a given resource.
  '   Used when the resource itself is deleted.
  ' Parameters:
  '   conn       - SQLite connection wrapper.
  '   resourceID - ResourceID for which to delete attributes.
  ' Returns:
  '   None
  ' Notes:
  '   - Uses RecordLoader and RecordSaver only.
  ' ==========================================================================================
  Private Sub DeleteAllResourceNameValues(ByVal conn As SQLiteConnectionWrapper,
                                          ByVal resourceID As String)

    Dim fieldNames() As String = {"ResourceID"}
    Dim fieldValues() As Object = {resourceID}
    Dim existing As List(Of RecordResourceNameValue)
    Dim rec As RecordResourceNameValue
    Dim toSave As New List(Of ISQLiteRecord)()

    If String.IsNullOrEmpty(resourceID) Then Exit Sub

    existing = RecordLoader.LoadRecordsByFields(Of RecordResourceNameValue)(
                 conn, fieldNames, fieldValues)

    If existing Is Nothing Then Exit Sub

    For Each rec In existing
      rec.IsDeleted = True
      toSave.Add(rec)
    Next

    If toSave.Count > 0 Then
      RecordSaver.SaveRecords(conn, toSave)
    End If

  End Sub

  ' ==========================================================================================
  ' Routine: ResetActionResource
  ' Purpose:
  '   Reset the action payload and pending action to an idle state.
  ' Parameters:
  '   model - UIModelResources to reset.
  ' Returns:
  '   None
  ' Notes:
  '   - Not used by LoadResourceModel, but available for explicit UI reset if needed.
  ' ==========================================================================================
  Private Sub ResetActionResource(ByVal model As UIModelResources)

    model.ActionResource = New UIResourceRow()
    model.ActionResourceNameValues = New SortableBindingList(Of UIResourceNameValueRow)()
    model.PendingAction = ResourceAction.None

  End Sub

  ' ==========================================================================================
  ' Routine: CloneResourceRow
  ' Purpose:
  '   Deep-copy a UIResourceRow so ActionResource can diverge from the snapshot without drift.
  ' Parameters:
  '   src - source UIResourceRow.
  ' Returns:
  '   UIResourceRow clone.
  ' Notes:
  '   - Copies IDs and editable fields only (no names).
  ' ==========================================================================================
  Private Function CloneResourceRow(ByVal src As UIResourceRow) As UIResourceRow

    If src Is Nothing Then Return New UIResourceRow()

    Dim dst As New UIResourceRow()

    dst.ResourceID = src.ResourceID
    dst.PreferredName = src.PreferredName
    dst.SalutationID = src.SalutationID
    dst.GenderID = src.GenderID
    dst.FirstName = src.FirstName
    dst.LastName = src.LastName

    dst.Email = src.Email
    dst.Phone = src.Phone

    dst.StartDate = src.StartDate
    dst.EndDate = src.EndDate

    dst.Notes = src.Notes

    Return dst

  End Function

  ' ==========================================================================================
  ' Routine: CloneResourceNameValueRow
  ' Purpose:
  '   Deep-copy a UIResourceNameValueRow so ActionResourceNameValues can diverge from
  '   the snapshot without drift.
  ' Parameters:
  '   src - source UIResourceNameValueRow.
  ' Returns:
  '   UIResourceNameValueRow clone.
  ' Notes:
  '   - Copies ListItems reference as-is; they are lookup lists, not edited by the user.
  ' ==========================================================================================
  Private Function CloneResourceNameValueRow(ByVal src As UIResourceNameValueRow) _
                                             As UIResourceNameValueRow

    If src Is Nothing Then Return New UIResourceNameValueRow()

    Dim dst As New UIResourceNameValueRow()

    dst.ResourceListItemID = src.ResourceListItemID
    dst.ResourceListItemName = src.ResourceListItemName
    dst.ListItemTypeID = src.ListItemTypeID
    dst.SelectedListItemID = src.SelectedListItemID
    dst.ResourceListItemValue = src.ResourceListItemValue
    dst.ValueType = src.ValueType
    dst.SelectedListItemIDs = If(src.SelectedListItemIDs Is Nothing,
                       New List(Of String)(),
                       New List(Of String)(src.SelectedListItemIDs))
    ' Lookup list can be shared; it is not mutated by the UI.
    dst.ListItems = If(src.ListItems Is Nothing,
                       New List(Of UIListItemRow)(),
                       New List(Of UIListItemRow)(src.ListItems))

    Return dst

  End Function


  ' ==========================================================================================
  ' Routine: ResolveRoleFunctionName
  ' Purpose:
  '   Resolve a RoleFunctionID to its ListItemName using the RoleFunctions lookup list.
  ' Parameters:
  '   roleFunctions  - List(Of UIListItemRow) lookup for "Role / Function".
  '   roleFunctionID - GUID string to resolve.
  ' Returns:
  '   String - RoleFunctionName if found; otherwise empty string.
  ' Notes:
  '   - Uses in-memory lookup only; no DB access.
  ' ==========================================================================================
  Private Function ResolveRoleFunctionName(ByVal roleFunctions As List(Of UIListItemRow),
                                           ByVal roleFunctionID As String) As String

    If roleFunctions Is Nothing OrElse String.IsNullOrEmpty(roleFunctionID) Then Return ""

    Dim li As UIListItemRow
    For Each li In roleFunctions
      If String.Equals(li.ListItemID, roleFunctionID, StringComparison.OrdinalIgnoreCase) Then
        Return li.ListItemName
      End If
    Next

    Return ""

  End Function

  ' ==========================================================================================
  ' Routine: LoadMultiSelectValues
  ' Purpose:
  '   Load all ListItemIDs for a multi-select ResourceListItem from tblResourceNameValueListItem.
  ' Parameters:
  '   conn               - SQLite connection wrapper.
  '   resourceID         - ResourceID.
  '   resourceListItemID - ResourceListItemID.
  ' Returns:
  '   List(Of String) - ListItemIDs for the multi-select attribute.
  ' Notes:
  '   - Filters out deleted rows.
  ' ==========================================================================================
  Private Function LoadMultiSelectValues(ByVal conn As SQLiteConnectionWrapper,
                                       ByVal resourceID As String,
                                       ByVal resourceListItemID As String) As List(Of String)

    Dim recs As List(Of RecordResourceNameValueListItem)

    recs = RecordLoader.LoadRecordsByFields(Of RecordResourceNameValueListItem)(
            conn,
            {"ResourceID", "ResourceListItemID"},
            {resourceID, resourceListItemID})

    If recs Is Nothing Then Return New List(Of String)()

    Dim result As New List(Of String)()
    Dim r As RecordResourceNameValueListItem
    For Each r In recs
      If r IsNot Nothing AndAlso Not r.IsDeleted AndAlso
       Not String.IsNullOrEmpty(r.ListItemID) Then
        result.Add(r.ListItemID)
      End If
    Next

    Return result

  End Function

  ' ==========================================================================================
  ' Routine: DeleteMultiSelectValues
  ' Purpose:
  '   Soft-delete all multi-select values for a given (ResourceID, ResourceListItemID)
  '   from tblResourceNameValueListItem.
  ' Parameters:
  '   conn               - SQLite connection wrapper.
  '   resourceID         - ResourceID.
  '   resourceListItemID - ResourceListItemID.
  ' Returns:
  '   None
  ' Notes:
  '   - Uses RecordLoader and RecordSaver only.
  ' ==========================================================================================
  Private Sub DeleteMultiSelectValues(ByVal conn As SQLiteConnectionWrapper,
                                    ByVal resourceID As String,
                                    ByVal resourceListItemID As String)

    Dim recs As List(Of RecordResourceNameValueListItem)

    recs = RecordLoader.LoadRecordsByFields(Of RecordResourceNameValueListItem)(
            conn,
            {"ResourceID", "ResourceListItemID"},
            {resourceID, resourceListItemID})

    If recs Is Nothing Then Exit Sub

    Dim rec As RecordResourceNameValueListItem
    For Each rec In recs
      rec.IsDeleted = True
      RecordSaver.SaveRecord(conn, rec)
    Next

  End Sub

End Module
