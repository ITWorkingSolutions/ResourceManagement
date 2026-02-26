Option Explicit On
Imports System.Text.Json
Imports ExcelDna.Integration
Imports Microsoft
Imports Microsoft.Office.Core
Imports Microsoft.Office.Interop

Friend Module UILoaderSaverExcelRuleDesigner

  ' ==========================================================================================
  ' Routine: LoadExcelRuleDesignerModel
  ' Purpose:
  '   Load or refresh the UIModelExcelRuleDesigner model using a single entry point.
  '   - First load (model Is Nothing): full initialisation of metadata, lists, and working sets.
  '   - Subsequent loads: refresh rule list and load detail for SelectedRule if present.
  ' Parameters:
  '   model - ByRef UIModelExcelRuleDesigner. Nothing on first load; existing model on refresh.
  ' Returns:
  '   None
  ' Notes:
  '   - UI stability is preserved. The form retains lookup lists, metadata, and working copies.
  '   - No SQL is written here; all DB access is via RecordLoader and RecordSaver.
  ' ==========================================================================================
  Friend Sub LoadExcelRuleDesignerModel(ByRef model As UIModelExcelRuleDesigner)

    Dim isFirstLoad As Boolean = (model Is Nothing)
    Dim conn As SQLiteConnectionWrapper = Nothing

    Try
      conn = OpenDatabase(AddInContext.Current.Config.ActiveDbPath)

      If isFirstLoad Then

        ' --- First load: initialise full model ---
        model = New UIModelExcelRuleDesigner()

        ' --- Load Apply instances ---
        LoadApplyInstances(model)

        ' --- Load metadata and helper ---
        ' 1. Load the raw map (unchanged)
        Dim baseMap As ExcelRuleViewMap = ExcelRuleViewMapLoader.LoadExcelRuleViewMap()

        ' 2. Load custom fields from DB
        Dim customFields As List(Of RecordResourceListItem) =
          RecordLoader.LoadRecords(Of RecordResourceListItem)(conn)

        ' 3. Build the UI-safe map
        Dim builder As New UIExcelRuleDesignerViewMapBuilder(baseMap)
        Dim map As ExcelRuleViewMap = builder.BuildUiViewMap(customFields)

        ' --- Load list types and list items (storage) and map to UI DTOs (no Record* on model) ---
        Try
          Dim listTypes = RecordLoader.LoadRecords(Of RecordListItemType)(conn)
          Dim listItems = RecordLoader.LoadRecords(Of RecordListItem)(conn)

          ' Map list types to DTOs
          If listTypes IsNot Nothing Then
            model.ListTypes = listTypes.Select(Function(lt) New ListTypeDescriptor With {
                                                  .Id = lt.ListItemTypeID,
                                                  .Name = If(String.IsNullOrEmpty(lt.ListItemTypeName), lt.ListItemTypeID, lt.ListItemTypeName),
                                                  .IsSystem = (lt.IsSystemType <> 0)
                                                }).ToList()
          Else
            model.ListTypes = New List(Of ListTypeDescriptor)()
          End If

          ' Group list items by ListItemTypeID
          Dim dict As New Dictionary(Of String, List(Of ListItemDescriptor))(StringComparer.OrdinalIgnoreCase)
          If listItems IsNot Nothing Then
            For Each li In listItems
              Dim key = If(li.ListItemTypeID, String.Empty).Trim()
              If String.IsNullOrEmpty(key) Then Continue For
              If Not dict.ContainsKey(key) Then dict(key) = New List(Of ListItemDescriptor)()
              dict(key).Add(New ListItemDescriptor With {
                              .Id = li.ListItemID,
                              .Name = li.ListItemName
                            })
            Next
          End If
          ''model.ListItemsByType = dict

        Catch ex As Exception
          ErrorHandler.UnHandleError(ex)
          model.ListTypes = New List(Of ListTypeDescriptor)()
          ''model.ListItemsByType = New Dictionary(Of String, List(Of ListItemDescriptor))(StringComparer.OrdinalIgnoreCase)
        End Try

        model.ViewMapHelper = New ExcelRuleViewMapHelper(map)

        ' --- Load rule list for left panel ---
        model.Rules = LoadAllRules(conn)

        ' --- Initialise empty detail + working copies ---
        model.SelectedRule = Nothing
        model.RuleDetail = New UIExcelRuleDesignerRuleRowDetail()
        model.ActionRule = New UIExcelRuleDesignerRuleRow()
        model.ActionRuleDetail = New UIExcelRuleDesignerRuleRowDetail()

        model.PendingAction = ExcelRuleDesignerAction.None

        ' --- Populate base lookup lists (views only) ---
        PopulateBaseLookupLists(model)

        Exit Sub

      End If

      ' --- Refresh path: keep UI stable, refresh only what is safe ---
      model.Rules = LoadAllRules(conn)

      ' --- Load Apply instances ---
      LoadApplyInstances(model)

      ' --- Load detail only if UI has selected a rule ---
      Dim selectedRuleID As String = model.SelectedRule?.RuleID
      If model.SelectedRule IsNot Nothing AndAlso
         Not String.IsNullOrEmpty(model.SelectedRule.RuleID) AndAlso
         model.Rules.Any(Function(r) r.RuleID = selectedRuleID) Then

        Dim pk() As Object = {model.SelectedRule.RuleID}

        ' --- Load DB record for selected rule ---
        Dim rec As RecordExcelRule =
          RecordLoader.LoadRecord(Of RecordExcelRule)(conn, pk)

        If rec Is Nothing OrElse rec.IsDeleted Then
          Throw New InvalidOperationException(
            $"RuleID '{model.SelectedRule.RuleID}' not found.")
        End If

        ' --- Deserialize JSON into detail snapshot ---
        'model.RuleDetail = New UIExcelRuleDesignerRuleRowDetail()
        model.RuleDetail = DeserializeRuleDetail(rec.DefinitionJson)
        model.RuleDetail.RuleID = model.SelectedRule.RuleID
        model.RuleDetail.RuleName = model.SelectedRule.RuleName
        model.RuleDetail.RuleType = model.SelectedRule.RuleType

        ' --- Update the rule field names in selected values and filters as the end user my have changed the names --- 
        For Each sv In model.RuleDetail.SelectedValues
          If Not String.IsNullOrEmpty(sv.FieldID) Then
            Dim f = model.ViewMapHelper.GetField(sv.View, sv.Field, sv.FieldID)
            If f IsNot Nothing Then sv.Field = f.Name
          End If
        Next

        For Each flt In model.RuleDetail.Filters
          If Not String.IsNullOrEmpty(flt.FieldID) Then
            Dim f = model.ViewMapHelper.GetField(flt.View, flt.Field, flt.FieldID)
            If f IsNot Nothing Then flt.Field = f.Name
          End If
        Next

        ' --- Ensure every filter has a stable FilterID ---
        For Each f In model.RuleDetail.Filters
          If String.IsNullOrEmpty(f.FilterID) Then
            f.FilterID = Guid.NewGuid().ToString()
          End If
        Next

      End If

      ' --- Ensure lookup lists exist (first load only) ---
      If model.AvailableViews Is Nothing OrElse model.AvailableViews.Count = 0 Then
        PopulateBaseLookupLists(model)
      End If

    Catch ex As Exception
      ErrorHandler.UnHandleError(ex)
      Throw
    Finally
      If conn IsNot Nothing Then conn.Close()
    End Try

  End Sub

  ' ==========================================================================================
  ' Routine: SavePendingRuleAction
  ' Purpose:
  '   Commit the pending Add / Update / Delete action encoded in the model.
  '   Uses ActionRule and ActionRuleDetail as the working copies.
  ' Parameters:
  '   model - UIModelExcelRuleDesigner containing PendingAction and working copies.
  ' Returns:
  '   None
  ' Notes:
  '   - Caller is responsible for refreshing the model after save.
  '   - No inference is performed; PendingAction must be explicitly set by the UI.
  ' ==========================================================================================
  Friend Sub SavePendingRuleAction(ByRef model As UIModelExcelRuleDesigner)

    Dim conn As SQLiteConnectionWrapper = Nothing

    If model Is Nothing Then Throw New ArgumentNullException(NameOf(model))

    Try
      conn = OpenDatabase(AddInContext.Current.Config.ActiveDbPath)

      ' --- Normal execution ---
      If model.PendingAction = ExcelRuleDesignerAction.None Then Exit Sub


      Select Case model.PendingAction

        Case ExcelRuleDesignerAction.Add
          ' --- Insert new rule ---
          Dim rec As RecordExcelRule = MapNewValuesToRecord(model)
          RecordSaver.SaveRecord(conn, rec)
          model.ActionRule.RuleID = rec.RuleID

        Case ExcelRuleDesignerAction.Update
          ' --- Update existing rule ---
          If model.ActionRule Is Nothing OrElse
             String.IsNullOrEmpty(model.ActionRule.RuleID) Then
            Throw New InvalidOperationException("No rule selected for update.")
          End If

          Dim rec As RecordExcelRule = MapUpdateValuesToRecord(model)
          RecordSaver.SaveRecord(conn, rec)

        Case ExcelRuleDesignerAction.Delete
          ' --- Soft-delete rule ---
          If model.ActionRule Is Nothing OrElse
             String.IsNullOrEmpty(model.ActionRule.RuleID) Then
            Throw New InvalidOperationException("No rule selected for delete.")
          End If

          Dim rec As New RecordExcelRule()
          rec.RuleID = model.ActionRule.RuleID
          rec.IsDeleted = True
          RecordSaver.SaveRecord(conn, rec)

        Case Else
          ' --- No action requested ---
          Return

      End Select

      ' --- Reset pending action ---
      model.PendingAction = ExcelRuleDesignerAction.None

    Catch ex As Exception
      ErrorHandler.UnHandleError(ex)
      Throw
    Finally
      If conn IsNot Nothing Then conn.Close()
    End Try

  End Sub

  ' ==========================================================================================
  ' Routine: LoadAllRules
  ' Purpose:
  '   Load all non-deleted Excel rules from the database and map them to list rows.
  ' Parameters:
  '   conn - SQLiteConnectionWrapper used for DB access.
  ' Returns:
  '   SortableBindingList(Of UIExcelRuleDesignerRuleRow)
  ' Notes:
  '   - Only list-level fields (RuleID, RuleName, RuleType) are mapped here.
  ' ==========================================================================================
  Private Function LoadAllRules(conn As SQLiteConnectionWrapper) _
                               As SortableBindingList(Of UIExcelRuleDesignerRuleRow)

    Dim list As New SortableBindingList(Of UIExcelRuleDesignerRuleRow)()
    Dim recs As List(Of RecordExcelRule)

    ' --- Load all rule records ---
    recs = RecordLoader.LoadRecords(Of RecordExcelRule)(conn)
    If recs Is Nothing Then Return list

    ' --- Map to UI list rows ---
    For Each r In recs
      If r IsNot Nothing AndAlso Not r.IsDeleted Then
        list.Add(MapRecordToListRow(r))
      End If
    Next

    Return list

  End Function

  ' ==========================================================================================
  ' Routine: MapRecordToListRow
  ' Purpose:
  '   Convert a RecordExcelRule into a UIExcelRuleDesignerRuleRow for list display.
  ' Parameters:
  '   r - RecordExcelRule loaded from the database.
  ' Returns:
  '   UIExcelRuleDesignerRuleRow
  ' Notes:
  '   - Only list-level fields are mapped.
  ' ==========================================================================================
  Private Function MapRecordToListRow(r As RecordExcelRule) As UIExcelRuleDesignerRuleRow
    Dim row As New UIExcelRuleDesignerRuleRow()

    ' --- Map list-level fields ---
    row.RuleID = r.RuleID
    row.RuleName = r.RuleName
    row.RuleType = r.RuleType

    Return row
  End Function

  ' ==========================================================================================
  ' Routine: MapUpdateValuesToRecord
  ' Purpose:
  '   Build a RecordExcelRule for update using ActionRule and ActionRuleDetail.
  ' Parameters:
  '   model - UIModelExcelRuleDesigner containing working copies.
  ' Returns:
  '   RecordExcelRule ready for RecordSaver.SaveRecord.
  ' Notes:
  '   - RuleID is preserved from ActionRule.
  ' ==========================================================================================
  Private Function MapUpdateValuesToRecord(model As UIModelExcelRuleDesigner) _
                                           As RecordExcelRule

    Dim rec As New RecordExcelRule()
    Dim a As UIExcelRuleDesignerRuleRow = model.ActionRule
    Dim d As UIExcelRuleDesignerRuleRowDetail = model.ActionRuleDetail

    ' --- Ensure new filters have FilterID ---
    For Each f In d.Filters
      If String.IsNullOrEmpty(f.FilterID) Then
        f.FilterID = Guid.NewGuid().ToString()
      End If
    Next

    ' --- Primary key + update flag ---
    rec.RuleID = a.RuleID
    rec.IsDirty = True

    ' --- Map editable fields ---
    rec.RuleName = a.RuleName
    rec.RuleType = a.RuleType

    ' --- Serialize detail JSON ---
    rec.DefinitionJson = SerializeRuleDetail(d)

    Return rec

  End Function

  ' ==========================================================================================
  ' Routine: MapNewValuesToRecord
  ' Purpose:
  '   Build a new RecordExcelRule for insertion using ActionRule and ActionRuleDetail.
  ' Parameters:
  '   model - UIModelExcelRuleDesigner containing working copies.
  ' Returns:
  '   RecordExcelRule ready for RecordSaver.SaveRecord.
  ' Notes:
  '   - Generates a new GUID for RuleID.
  ' ==========================================================================================
  Private Function MapNewValuesToRecord(model As UIModelExcelRuleDesigner) _
                                        As RecordExcelRule

    Dim rec As New RecordExcelRule()
    Dim a As UIExcelRuleDesignerRuleRow = model.ActionRule
    Dim d As UIExcelRuleDesignerRuleRowDetail = model.ActionRuleDetail

    ' --- Ensure new filters have FilterID ---
    For Each f In d.Filters
      If String.IsNullOrEmpty(f.FilterID) Then
        f.FilterID = Guid.NewGuid().ToString()
      End If
    Next

    ' --- Primary key + new flag ---
    rec.IsNew = True
    rec.RuleID = Guid.NewGuid().ToString()


    ' --- Map editable fields ---
    rec.RuleName = a.RuleName
    rec.RuleType = a.RuleType

    ' --- Serialize detail JSON ---
    rec.DefinitionJson = SerializeRuleDetail(d)

    Return rec

  End Function

  ' ==========================================================================================
  ' Routine: SerializeRuleDetail
  ' Purpose:
  '   Convert a UIExcelRuleDesignerRuleRowDetail into a JSON string for storage.
  ' Parameters:
  '   detail - UIExcelRuleDesignerRuleRowDetail to serialize.
  ' Returns:
  '   JSON string.
  ' Notes:
  '   - Uses camelCase naming to match existing JSON conventions.
  ' ==========================================================================================
  Private Function SerializeRuleDetail(detail As UIExcelRuleDesignerRuleRowDetail) As String
    ' --- Configure serializer ---
    Dim options As New JsonSerializerOptions With {
      .WriteIndented = False,
      .PropertyNamingPolicy = JsonNamingPolicy.CamelCase
    }

    ' --- Serialize to JSON ---
    Return JsonSerializer.Serialize(detail, options)

  End Function

  ' ==========================================================================================
  ' Routine: DeserializeRuleDetail
  ' Purpose:
  '   Convert a JSON string into a UIExcelRuleDesignerRuleRowDetail object.
  ' Parameters:
  '   json - JSON string from RecordExcelRule.DefinitionJson.
  ' Returns:
  '   UIExcelRuleDesignerRuleRowDetail
  ' Notes:
  '   - Returns an empty detail object when json is blank.
  ' ==========================================================================================
  Private Function DeserializeRuleDetail(json As String) As UIExcelRuleDesignerRuleRowDetail

    If String.IsNullOrEmpty(json) Then
      ' --- No JSON stored yet ---
      Return New UIExcelRuleDesignerRuleRowDetail()
    End If

    ' --- Configure deserializer ---
    Dim options As New JsonSerializerOptions With {
      .PropertyNameCaseInsensitive = True
    }

    ' --- Deserialize JSON ---
    Return JsonSerializer.Deserialize(Of UIExcelRuleDesignerRuleRowDetail)(json, options)

  End Function

  ' ==========================================================================================
  ' Routine: PopulateBaseLookupLists
  ' Purpose:
  '   Populate lookup lists that depend only on ViewMap metadata.
  ' Parameters:
  '   model - UIModelExcelRuleDesigner to populate.
  ' Returns:
  '   None
  ' Notes:
  '   - AvailableFields and AvailableOperators are context-dependent and left empty.
  ' ==========================================================================================
  Private Sub PopulateBaseLookupLists(model As UIModelExcelRuleDesigner)

    Dim h = model.ViewMapHelper

    ' --- Views come directly from metadata ---
    model.AvailableViews = h.GetAllViewNames()
    ' --- Fields/operators are populated dynamically by UI ---
    model.AvailableFields = New List(Of String)()
    ' --- Populate global operator list ---
    model.AvailableOperators = New List(Of String) From {"=", ">", "<", "<>", ">=", "<="}
    ' --- Populate global open parentheses list ---
    model.AvailableOpenParentheses = New List(Of String) From {"", "(", "((", "((("}
    ' --- Populate global close parentheses list ---
    model.AvailableCloseParentheses = New List(Of String) From {"", ")", "))", ")))"}
    ' --- Populate global boolean operator list ---
    model.AvailableBooleanOperators = New List(Of String) From {"", "AND", "OR"}
  End Sub


  ' ==========================================================================================
  ' Routine: CloneListRow
  ' Purpose:
  '   Create a working copy of a UIExcelRuleDesignerRuleRow for editing.
  ' Parameters:
  '   src - UIExcelRuleDesignerRuleRow to clone.
  ' Returns:
  '   UIExcelRuleDesignerRuleRow
  ' Notes:
  '   - Only list-level fields are cloned.
  ' ==========================================================================================
  Private Function CloneListRow(src As UIExcelRuleDesignerRuleRow) _
                               As UIExcelRuleDesignerRuleRow

    If src Is Nothing Then Return New UIExcelRuleDesignerRuleRow()

    ' --- Shallow clone of list-level fields ---
    Dim dst As New UIExcelRuleDesignerRuleRow()
    dst.RuleID = src.RuleID
    dst.RuleName = src.RuleName
    dst.RuleType = src.RuleType

    Return dst

  End Function

  ' ==========================================================================================
  ' Routine: AddListTypeViewsToMap
  ' Purpose:
  '   Materialise synthetic per-list views into the UI view map using the
  '   existing RecordLoader (avoids raw ADO calls and keeps project conventions).
  ' Notes:
  '   - Each synthetic view is named "vwDim_List_{ListTypeID}" (internal).
  '   - DisplayName is the user-facing ListTypeName.
  ' ==========================================================================================
  Private Sub AddListTypeViewsToMap(conn As SQLiteConnectionWrapper, uiMap As ExcelRuleViewMap)
    If conn Is Nothing OrElse uiMap Is Nothing Then Exit Sub

    Try
      ' Use RecordLoader to fetch list types (follows existing project patterns)
      Dim listTypes As List(Of RecordListItemType) = Nothing
      Try
        listTypes = RecordLoader.LoadRecords(Of RecordListItemType)(conn)
      Catch ex As Exception
        ErrorHandler.UnHandleError(ex)
        Return
      End Try

      If listTypes Is Nothing OrElse listTypes.Count = 0 Then Return

      For Each lt In listTypes
        If lt Is Nothing Then Continue For
        Dim id = If(lt.ListItemTypeID, String.Empty).Trim()
        If String.IsNullOrEmpty(id) Then Continue For

        Dim viewName As String = $"vwDim_List_{id}"

        ' Skip duplicates
        If uiMap.Views.Any(Function(v) v.Name.Equals(viewName, StringComparison.OrdinalIgnoreCase)) Then
          Continue For
        End If

        Dim displayName As String = If(String.IsNullOrEmpty(lt.ListItemTypeName), id, lt.ListItemTypeName)

        Dim view As New ExcelRuleViewMapView With {
          .Name = viewName,
          .DisplayName = displayName,
          .Fields = New List(Of ExcelRuleViewMapField)(),
          .Relations = New List(Of ExcelRuleViewMapRelation)()
        }

        uiMap.Views.Add(view)
      Next

    Catch ex As Exception
      ErrorHandler.UnHandleError(ex)
    End Try
  End Sub
  ' ==========================================================================================
  ' Routine: CloneDetail
  ' Purpose:
  '   Create a deep working copy of a UIExcelRuleDesignerRuleRowDetail.
  ' Parameters:
  '   src - UIExcelRuleDesignerRuleRowDetail to clone.
  ' Returns:
  '   UIExcelRuleDesignerRuleRowDetail
  ' Notes:
  '   - Uses JSON round-trip for deterministic deep cloning.
  ' ==========================================================================================
  Private Function CloneDetail(src As UIExcelRuleDesignerRuleRowDetail) _
                               As UIExcelRuleDesignerRuleRowDetail

    If src Is Nothing Then Return New UIExcelRuleDesignerRuleRowDetail()

    ' --- Deep clone via JSON round-trip ---
    Dim json = SerializeRuleDetail(src)
    Return DeserializeRuleDetail(json)

  End Function
#Region "Apply Routines"
  ' ==========================================================================================
  ' Routine: LoadApplyInstances
  ' Purpose:
  '   Populate model.ApplyInstances by calling ExcelCellRuleStore.GetAllApplyInstances and
  '   mapping each storage-layer ExcelApplyInstance into a UIExcelApplyInstance.
  '
  ' Parameters:
  '   model - UIModelExcelRuleDesigner to populate.
  '
  ' Returns:
  '   None.
  '
  ' Notes:
  '   - UILoaderSaver must not touch XML; ExcelCellRuleStore hides all storage details.
  '   - ApplyName is UI-only and not stored in XML.
  ' ==========================================================================================
  Friend Sub LoadApplyInstances(model As UIModelExcelRuleDesigner)

    Const RoutineName As String = "LoadApplyInstances"

    Dim wb As Excel.Workbook = Nothing
    Dim storageList As List(Of ExcelApplyInstance) = Nothing
    Dim storageItem As ExcelApplyInstance = Nothing
    Dim uiItem As UIExcelRuleDesignerApplyInstance = Nothing

    Try
      ' --- Normal execution ---
      wb = CType(ExcelDnaUtil.Application, Excel.Application).ActiveWorkbook
      If wb Is Nothing Then Exit Sub

      ' Load storage-layer apply instances
      storageList = ExcelCellRuleStore.GetAllApplyInstances(wb)

      ' Reset UI list
      model.ApplyInstances = New List(Of UIExcelRuleDesignerApplyInstance)

      ' Map storage-layer → UI model
      For Each storageItem In storageList

        uiItem = New UIExcelRuleDesignerApplyInstance
        uiItem.ApplyID = storageItem.ApplyID
        uiItem.RuleID = storageItem.RuleID
        uiItem.ListSelectType = storageItem.ListSelectType

        ' Map parameters
        uiItem.Parameters = New List(Of UIExcelRuleDesignerApplyParameter)
        For Each p In storageItem.Parameters
          Dim uiParam As New UIExcelRuleDesignerApplyParameter With {
            .FilterID = p.Name,
            .RefType = p.RefType,
            .RefValue = p.RefValue,
            .LiteralValue = p.LiteralValue
          }
          uiItem.Parameters.Add(uiParam)
        Next

        ' ApplyName is UI-only
        uiItem.ApplyName = storageItem.ApplyName

        model.ApplyInstances.Add(uiItem)

      Next

    Catch ex As Exception
      ErrorHandler.UnHandleError(ex, RoutineName)

    Finally
      ' --- Cleanup ---
      wb = Nothing
      storageList = Nothing
      storageItem = Nothing
      uiItem = Nothing

    End Try

  End Sub

  ' ==========================================================================================
  ' Routine: SavePendingApplyAction
  ' Purpose:
  '   Commit the pending Add / Update / Delete action for Apply instances.
  '   Uses model.PendingApplyAction and model.ActionApply as the working copy.
  '
  ' Parameters:
  '   model - UIModelExcelRuleDesignerRule containing PendingApplyAction and ActionApply.
  '
  ' Returns:
  '   None.
  '
  ' Notes:
  '   - This is the ONLY entry point for committing Apply changes.
  '   - UI must set PendingApplyAction and ActionApply before calling this routine.
  '   - UILoaderSaver must not touch XML; ExcelCellRuleStore handles all storage.
  '   - Cell membership is NOT modified here; only rule metadata is updated.
  ' ==========================================================================================
  Friend Sub SavePendingApplyAction(ByRef model As UIModelExcelRuleDesigner,
                                   target As Excel.Range)

    Dim wb As Excel.Workbook = Nothing
    Dim storage As ExcelApplyInstance = Nothing

    Try
      If model Is Nothing Then Exit Sub
      If model.PendingAction = ExcelRuleDesignerAction.None Then Exit Sub
      If model.ActionApply Is Nothing Then Exit Sub

      wb = CType(ExcelDnaUtil.Application, Excel.Application).ActiveWorkbook
      If wb Is Nothing Then Exit Sub

      storage = New ExcelApplyInstance()
      storage.ApplyID = model.ActionApply.ApplyID
      storage.ApplyName = model.ActionApply.ApplyName
      storage.RuleID = model.ActionApply.RuleID
      storage.ListSelectType = model.ActionApply.ListSelectType
      storage.Parameters = New List(Of RuleParameter)

      For Each uiParam In model.ActionApply.Parameters
        storage.Parameters.Add(
          New RuleParameter With {
            .Name = uiParam.FilterID,
            .RefType = uiParam.RefType,
            .RefValue = uiParam.RefValue,
            .LiteralValue = uiParam.LiteralValue
          })
      Next

      Select Case model.PendingAction
        Case ExcelRuleDesignerAction.Add
          storage.IsNew = True
          storage.IsDirty = True
          storage.IsDeleted = False

        Case ExcelRuleDesignerAction.Update
          storage.IsNew = False
          storage.IsDirty = True
          storage.IsDeleted = False

        Case ExcelRuleDesignerAction.Delete
          storage.IsNew = False
          storage.IsDirty = True
          storage.IsDeleted = True
      End Select

      ExcelCellRuleStore.SaveApplyInstance(wb, storage, target)
      model.PendingAction = ExcelRuleDesignerAction.None

    Catch ex As Exception
      ErrorHandler.UnHandleError(ex)

    Finally
      wb = Nothing
      storage = Nothing
    End Try

  End Sub

#End Region
End Module