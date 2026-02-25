Imports System.ComponentModel

Module UILoaderSaverClosure

  ' ==========================================================================================
  ' Routine: LoadClosuresModel
  ' Purpose: Load closure records from the active database and return a UIModelClosures.
  ' Parameters:
  ' Returns:
  '   UIModelClosures - model containing the Closures collection and reset action state.
  ' Notes:
  '   - Uses OpenDatabase and RecordLoader only (no direct SQL).
  '   - Maps RecordClosure -> UIClosureRow via MapRecordToUIClosure.
  ' ==========================================================================================
  Friend Function LoadClosuresModel() As UIModelClosures
    Dim conn As SQLiteConnectionWrapper = Nothing

    Try
      Dim model As New UIModelClosures()
      model.Closures = New SortableBindingList(Of UIClosureRow)()

      ' Open a validated DB connection using the project's OpenDatabase routine
      conn = OpenDatabase(AddInContext.Current.Config.ActiveDbPath)

      ' Load domain records via generic loader
      Dim recs As List(Of RecordClosure) = RecordLoader.LoadRecords(Of RecordClosure)(conn)

      ' Map domain records to UI rows
      For Each r In recs
        model.Closures.Add(MapRecordToUIClosure(r))
      Next

      ' No selected/default values set here (UI chooses selection)
      ResetActionClosure(model)

      Return model

    Catch ex As Exception
      ErrorHandler.UnHandleError(ex)
      Throw
    Finally
      If conn IsNot Nothing Then conn.Close()
    End Try
  End Function

  ' ==========================================================================================
  ' Routine: SavePendingClosureAction
  ' Purpose: Perform the pending action encoded in the UI model (Add / Update / Delete).
  ' Parameters:
  '   model - UIModelClosures containing PendingAction and related data
  ' Returns:
  ' Notes:
  '   - Uses OpenDatabase and RecordSaver.
  '   - Refreshes the model after a successful save so the UI reflects canonical DB state.
  ' ==========================================================================================
  Friend Sub SavePendingClosureAction(ByRef model As UIModelClosures)
    If model Is Nothing Then Throw New ArgumentNullException(NameOf(model))

    Dim conn As SQLiteConnectionWrapper = Nothing

    Try
      conn = OpenDatabase(AddInContext.Current.Config.ActiveDbPath)

      Select Case model.PendingAction
        Case ClosureAction.Add
          Dim newRec As RecordClosure = MapNewValuesToRecord(model)
          RecordSaver.SaveRecord(conn, newRec)

        Case ClosureAction.Update
          If model.ActionClosure.ClosureID Is Nothing Then
            Throw New InvalidOperationException("No closure selected for update.")
          End If
          ' Use the selected ClosureID and the UI's New* values for the update
          Dim updRec As RecordClosure = MapUpdateValuesToRecord(model)
          RecordSaver.SaveRecord(conn, updRec)

        Case ClosureAction.Delete
          If model.ActionClosure.ClosureID Is Nothing Then
            Throw New InvalidOperationException("No closure selected for delete.")
          End If
          ' Only the PK is required for delete
          Dim delRec As New RecordClosure()
          delRec.ClosureID = model.ActionClosure.ClosureID
          delRec.IsDeleted = True
          RecordSaver.SaveRecord(conn, delRec)

        Case Else
          ' Nothing to do
      End Select

      ' Refresh the UI collection so the UI sees canonical data (IDs, etc.)
      Dim refreshed As UIModelClosures = LoadClosuresModel()
      model = refreshed

    Catch ex As Exception
      ErrorHandler.UnHandleError(ex)
      Throw
    Finally
      If conn IsNot Nothing Then conn.Close()
    End Try
  End Sub

  ' ==========================================================================================
  ' Routine: MapRecordToUIClosure
  ' Purpose: Map a RecordClosure domain object to a UIClosureRow for display.
  ' Parameters:
  '   r - RecordClosure instance loaded from the database
  ' Returns:
  '   UIClosureRow - mapped UI row
  ' Notes:
  '   - Adjust parsing if RecordClosure fields use different types/formats.
  ' ==========================================================================================
  Private Function MapRecordToUIClosure(r As RecordClosure) As UIClosureRow
    Dim row As New UIClosureRow()

    row.ClosureID = r.ClosureID
    row.ClosureName = r.ClosureName
    row.StartDate = Date.ParseExact(r.StartDate.ToString(), "yyyy-MM-dd", Nothing)
    row.EndDate = Date.ParseExact(r.EndDate.ToString(), "yyyy-MM-dd", Nothing)

    Return row
  End Function

  ' ==========================================================================================
  ' Routine: MapUpdateValuesToRecord
  ' Purpose: Build a RecordClosure for update from the UIModelClosures.ActionClosure payload.
  ' Parameters:
  '   model - UIModelClosures containing ActionClosure with updated values and ClosureID
  ' Returns:
  '   RecordClosure - ready for RecordSaver.SaveRecord (IsDirty = True)
  ' Notes:
  '   - Dates are formatted as "yyyy-MM-dd" strings for storage.
  ' ==========================================================================================
  Private Function MapUpdateValuesToRecord(model As UIModelClosures) As RecordClosure
    Dim r As New RecordClosure()

    ' Ensure primary key is preserved
    r.ClosureID = model.ActionClosure.ClosureID
    r.IsDirty = True
    r.ClosureName = model.ActionClosure.ClosureName
    r.StartDate = model.ActionClosure.StartDate.ToString("yyyy-MM-dd")
    r.EndDate = model.ActionClosure.EndDate.ToString("yyyy-MM-dd")

    Return r
  End Function

  ' ==========================================================================================
  ' Routine: MapNewValuesToRecord
  ' Purpose: Build a RecordClosure for insertion from the UIModelClosures.ActionClosure payload.
  ' Parameters:
  '   model - UIModelClosures containing ActionClosure with new values
  ' Returns:
  '   RecordClosure - ready for RecordSaver.SaveRecord (IsNew = True)
  ' Notes:
  '   - Generates a GUID string for ClosureID.
  '   - Dates are formatted as "yyyy-MM-dd" strings for storage.
  ' ==========================================================================================
  Private Function MapNewValuesToRecord(model As UIModelClosures) As RecordClosure
    Dim r As New RecordClosure()
    ' Generate PK using GUID string for new records
    r.ClosureID = Guid.NewGuid().ToString()
    r.IsNew = True
    r.ClosureName = model.ActionClosure.ClosureName
    r.StartDate = model.ActionClosure.StartDate.ToString("yyyy-MM-dd")
    r.EndDate = model.ActionClosure.EndDate.ToString("yyyy-MM-dd")

    Return r
  End Function

  ' ==========================================================================================
  ' Routine: ResetActionClosure
  ' Purpose: Private helper to reset the action payload and pending action to an idle state.
  ' Parameters:
  '   model - UIModelClosures to reset
  ' Returns: 
  ' Notes:
  '   - Ensures the UI has a clean ActionClosure object and PendingAction = None.
  ' ==========================================================================================
  Private Sub ResetActionClosure(model As UIModelClosures)
    model.ActionClosure = New UIClosureRow()
    model.PendingAction = ClosureAction.None
  End Sub

End Module
