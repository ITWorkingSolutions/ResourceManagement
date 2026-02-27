Imports System.ComponentModel
Imports System.IO
Imports System.Linq
Imports System.Runtime.InteropServices
Imports System.Windows.Forms

Module UILoaderSaverDatabaseManagement

  ' ===========================================================================
  ' Routine: LoadDatabaseManagementModel
  ' Purpose:
  '   - Create and return a UIModelDatabaseManagement populated from the current
  '     AddInContext configuration.
  '   - Ensures the Paths collection is always a non-null BindingList(Of String).
  ' Returns:
  '   UIModelDatabaseManagement
  ' Notes:
  '   - This routine does only read+shape configuration; it does not perform any
  '     validation against the filesystem or attempt to open databases.
  ' ===========================================================================
  Friend Function LoadDatabaseManagementModel() As UIModelDatabaseManagement
    Try
      Dim cfg = AddInContext.Current.Config
      Dim list As List(Of String) = Nothing
      If cfg IsNot Nothing AndAlso cfg.KnownDbPaths IsNot Nothing Then
        list = cfg.KnownDbPaths.ToList()
      Else
        list = New List(Of String)()
      End If

      Dim model As New UIModelDatabaseManagement With {
        .Paths = New BindingList(Of String)(list),
        .ActivePath = If(cfg IsNot Nothing, cfg.ActiveDbPath, Nothing)
      }
      Return model
    Catch ex As Exception
      ErrorHandler.UnHandleError(ex)
      Throw
    End Try
  End Function

  Friend Sub SaveDatabaseManagementModel(ByRef model As UIModelDatabaseManagement)

    Try
      Select Case model.PendingAction
        Case DatabaseManagementAction.AddNew : HandleAdd(model)
        Case DatabaseManagementAction.AddExisting : HandleAdd(model)
        Case DatabaseManagementAction.Update : HandleUpdate(model)
        Case DatabaseManagementAction.Delete : HandleDelete(model)
        Case DatabaseManagementAction.Activate : HandleActivate(model)
        Case Else
          Throw New UserFriendlyException("No action selected.")
      End Select
      ' === Clear action and inputs ===
      model.PendingAction = DatabaseManagementAction.None
      model.NewPath = String.Empty
      model.SelectedPath = String.Empty
    Catch ex As UserFriendlyException
      ' Pass through unchanged
      Throw
    Catch ex As Exception
      ' Wrap technical errors
      Throw New UserFriendlyException(
            "Unable to complete the requested operation." & vbCrLf & ex.Message)
    Finally
      model.PendingAction = DatabaseManagementAction.None
    End Try

  End Sub

  ' ===========================================================================
  ' Routine: HandleAdd
  ' Purpose:
  '   - Add the provided model.newPath to KnownDbPaths (if not already present) and
  '     make the database
  ' Parameters:
  '   model       - UIModelDatabaseManagement
  ' Notes:
  '   - Path normalization is performed to ensure consistent casing and format.
  '   - This routine persists configuration via DatabaseConfigManager.Save.
  '   - The database file is created if it does not already exist.
  '   - model.Paths is updated to include the new path if it was not already present.
  ' ===========================================================================
  Private Sub HandleAdd(model As UIModelDatabaseManagement)
    Dim conn As SQLiteConnectionWrapper = Nothing

    Try
      Dim newPath As String = model.NewPath

      ' === Make sure required parameters are present ===
      If String.IsNullOrWhiteSpace(newPath) Then
        Throw New ArgumentNullException(NameOf(newPath))
      End If

      ' === Normalize paths ===
      newPath = NormalizePathSafe(newPath)

      ' === Get current config ===
      Dim cfg = AddInContext.Current.Config
      If cfg Is Nothing Then
        Throw New InvalidOperationException("AddInContext.Current.Config is not available.")
      End If

      If model.PendingAction = DatabaseManagementAction.AddNew Then
        ' === Create the database using the full lifecycle ===
        Try
          conn = OpenDatabase(newPath)
        Catch ex As Exception
          ErrorHandler.UnHandleError(ex)
          Exit Sub
        Finally
          If conn IsNot Nothing Then
            conn.Close()
          End If
        End Try
      ElseIf model.PendingAction = DatabaseManagementAction.AddExisting Then
        ' === Validate existing database ===
        Try
          conn = OpenDatabase(newPath)
        Catch ex As Exception
          MessageBoxHelper.Show(UIMessageOwner.ExcelOwner, "The selected file is not a valid existing database.",
                          "Error",
                          MessageBoxButtons.OK,
                          MessageBoxIcon.Warning)
          'If conn IsNot Nothing Then conn.Close()
          Exit Sub
        Finally
          If conn IsNot Nothing Then
            conn.Close()
          End If

        End Try
      End If

      ' === Create KnownDbPaths list if missing ===
      If cfg.KnownDbPaths Is Nothing Then
        cfg.KnownDbPaths = New List(Of String)()
      End If

      ' === Add if missing to config (case-insensitive) ===
      If Not cfg.KnownDbPaths.Any(Function(p) String.Equals(p, newPath, StringComparison.OrdinalIgnoreCase)) Then
        cfg.KnownDbPaths.Add(newPath)
      End If

      ' === Add if missing to model (case-insensitive) ===
      If Not model.Paths.Contains(newPath, StringComparer.OrdinalIgnoreCase) Then
        model.Paths.Add(newPath)
      End If

      ' === Save config ===
      DatabaseConfigManager.Save(cfg)
    Catch ex As Exception
      ErrorHandler.UnHandleError(ex)
      Throw
    End Try
  End Sub
  ' ===========================================================================
  ' Routine: HandleUpdate
  ' Purpose:
  '   - Update an existing database path (selectedPath) to a new path (newPath)
  ' Parameters:
  '   model       - UIModelDatabaseManagement
  ' Notes:
  '   - Path normalization is performed to ensure consistent casing and format.
  '   - The newPath must already exist as a valid SQLite database file.
  '   - This routine persists configuration via DatabaseConfigManager.Save.
  ' ===========================================================================
  Private Sub HandleUpdate(model As UIModelDatabaseManagement)
    Dim conn As SQLiteConnectionWrapper = Nothing

    Try
      Dim index As Integer
      Dim newPath As String = model.NewPath
      Dim selectedPath As String = model.SelectedPath

      ' === Make sure required parameters are present ===
      If String.IsNullOrWhiteSpace(newPath) Then
        Throw New ArgumentNullException(NameOf(newPath))
      End If
      If String.IsNullOrWhiteSpace(selectedPath) Then
        Throw New ArgumentNullException(NameOf(selectedPath))
      End If

      ' === Normalize paths ===
      newPath = NormalizePathSafe(newPath)
      selectedPath = NormalizePathSafe(selectedPath)

      ' === Get current config ===
      Dim cfg = AddInContext.Current.Config
      If cfg Is Nothing Then
        Throw New InvalidOperationException("AddInContext.Current.Config is not available.")
      End If

      ' === New path must exist when updating (Update cannot create new DBs) ===
      If Not IO.File.Exists(newPath) Then
        MessageBoxHelper.Show(UIMessageOwner.ExcelOwner, "The selected file does not exist. Use Add instead of Update.",
                          "Error",
                          MessageBoxButtons.OK,
                          MessageBoxIcon.Warning)
        Exit Sub
      End If

      ' === Validate that the new file is a valid existing SQLite DB ===
      Try
        conn = OpenDatabase(newPath)   ' Will throw if invalid or new
      Catch ex As UserFriendlyException
        MessageBoxHelper.Show(UIMessageOwner.ExcelOwner, "The selected file is not a valid existing database." &
                              vbCrLf & "Use Add instead of Update.",
                              "Invalid Database",
                              MessageBoxButtons.OK,
                              MessageBoxIcon.Warning)
        Exit Sub
      Catch ex As Exception
        ErrorHandler.UnHandleError(ex)
        Exit Sub
      Finally
        If conn IsNot Nothing Then conn.Close()
      End Try

      ' === Create KnownDbPaths list if missing ===
      If cfg.KnownDbPaths Is Nothing Then
        cfg.KnownDbPaths = New List(Of String)()
      End If

      ' === Update model ===
      index = model.Paths.IndexOf(selectedPath)
      If index >= 0 Then
        model.Paths(index) = newPath
      Else
        ' === If selectedPath wasn't found, add newPath to the list (upsert behaviour) ===
        If Not model.Paths.Contains(newPath, StringComparer.OrdinalIgnoreCase) Then
          model.Paths.Add(newPath)
        End If
      End If

      ' === Update config ===
      index = cfg.KnownDbPaths.FindIndex(Function(p) String.Equals(p, selectedPath, StringComparison.OrdinalIgnoreCase))
      If index >= 0 Then
        cfg.KnownDbPaths(index) = newPath
      Else
        ' === If selectedPath wasn't found, add newPath to the list (upsert behaviour) ===
        If Not cfg.KnownDbPaths.Any(Function(p) String.Equals(p, newPath, StringComparison.OrdinalIgnoreCase)) Then
          cfg.KnownDbPaths.Add(newPath)
        End If
      End If

      ' === Update active path if it referenced the old path ===
      If String.Equals(model.ActivePath, selectedPath, StringComparison.OrdinalIgnoreCase) Then
        model.ActivePath = newPath
        cfg.ActiveDbPath = newPath
      End If

      ' === Save config ===
      DatabaseConfigManager.Save(cfg)
    Catch ex As Exception
      ErrorHandler.UnHandleError(ex)
      Throw
    End Try
  End Sub

  ' ===========================================================================
  ' Routine: HandleDelete
  ' Purpose:
  '   - Delete the selectedPath from KnownDbPaths.
  ' Parameters:
  '   model       - UIModelDatabaseManagement
  ' Notes:
  '   - Path normalization is performed to ensure consistent casing and format.
  '   - This routine persists configuration via DatabaseConfigManager.Save.
  ' ===========================================================================
  Friend Sub HandleDelete(model As UIModelDatabaseManagement)
    Try
      Dim selectedPath As String = model.SelectedPath
      ' === Make sure required parameters are present ===
      If String.IsNullOrWhiteSpace(selectedPath) Then
        Throw New ArgumentNullException(NameOf(selectedPath))
      End If
      ' === Normalize path ===
      selectedPath = NormalizePathSafe(selectedPath)
      ' === Get current config ===
      Dim cfg = AddInContext.Current.Config
      If cfg Is Nothing Then
        Throw New InvalidOperationException("AddInContext.Current.Config is not available.")
      End If
      ' === Create KnownDbPaths list if missing ===
      If cfg.KnownDbPaths Is Nothing Then
        cfg.KnownDbPaths = New List(Of String)()
      End If
      ' === Remove from config ===
      cfg.KnownDbPaths = cfg.KnownDbPaths.
                          Where(Function(p) Not String.Equals(p, selectedPath, StringComparison.OrdinalIgnoreCase)).
                          ToList()
      ' === Remove from model ===
      Dim index As Integer = model.Paths.IndexOf(selectedPath)
      If index >= 0 Then
        model.Paths.RemoveAt(index)
      End If
      ' === Clear active path if it referenced the deleted path ===
      If String.Equals(model.ActivePath, selectedPath, StringComparison.OrdinalIgnoreCase) Then
        model.ActivePath = Nothing
        cfg.ActiveDbPath = Nothing
      End If
      ' === Save config ===
      DatabaseConfigManager.Save(cfg)
    Catch ex As Exception
      ErrorHandler.UnHandleError(ex)
      Throw
    End Try
  End Sub

  ' ===========================================================================
  ' Routine: HandleActivate
  ' Purpose:
  '   - Set the selectedPath as the active database path in configuration.
  ' Parameters:
  '   model       - UIModelDatabaseManagement
  ' Notes:
  '   - Path normalization is performed to ensure consistent casing and format.
  '   - This routine persists configuration via DatabaseConfigManager.Save.
  ' ===========================================================================
  Friend Sub HandleActivate(model As UIModelDatabaseManagement)
    Try
      Dim selectedPath As String = model.SelectedPath
      ' === Make sure required parameters are present ===
      If String.IsNullOrWhiteSpace(selectedPath) Then
        Throw New ArgumentNullException(NameOf(selectedPath))
      End If
      ' === Normalize path ===
      selectedPath = NormalizePathSafe(selectedPath)
      ' === Get current config ===
      Dim cfg = AddInContext.Current.Config
      If cfg Is Nothing Then
        Throw New InvalidOperationException("AddInContext.Current.Config is not available.")
      End If
      ' === Create KnownDbPaths list if missing ===
      If cfg.KnownDbPaths Is Nothing Then
        cfg.KnownDbPaths = New List(Of String)()
      End If
      ' === Ensure the selected path exists in KnownDbPaths ===
      If cfg.KnownDbPaths.Any(Function(p) String.Equals(p, selectedPath, StringComparison.OrdinalIgnoreCase)) Then
        ' === Set active path in config and model ===
        cfg.ActiveDbPath = selectedPath
      Else
        Throw New UserFriendlyException("The selected path is not in the list of known database paths.")
      End If
      ' === Ensure the selected path exists in the model Paths collection ===
      Dim index As Integer = model.Paths.IndexOf(selectedPath)
      If index >= 0 Then
        ' === Set active path in model ===
        model.ActivePath = selectedPath
      Else
        Throw New UserFriendlyException("The selected path is not in the list of known database paths.")
      End If
      ' === Save config ===
      DatabaseConfigManager.Save(cfg)
    Catch ex As Exception
      ErrorHandler.UnHandleError(ex)
      Throw
    End Try
  End Sub

  ' Helper: List.FindIndex extension for VB projects that don't have it on List(Of T)
  <System.Runtime.CompilerServices.Extension()>
  Private Function FindIndex(Of T)(list As List(Of T), predicate As Func(Of T, Boolean)) As Integer
    For i As Integer = 0 To list.Count - 1
      If predicate(list(i)) Then Return i
    Next
    Return -1
  End Function

  Private Function NormalizePathSafe(path As String) As String
    If String.IsNullOrWhiteSpace(path) Then
      Throw New ArgumentNullException(NameOf(path))
    End If

    Try
      ' GetFullPath is fine for already-absolute paths (it will normalize them).
      Return IO.Path.GetFullPath(path)
    Catch ex As ArgumentException
      Throw New ArgumentException("The path is malformed or contains invalid characters.", NameOf(path), ex)
    Catch ex As PathTooLongException
      Throw ' propagate so caller can decide how to handle
    Catch ex As NotSupportedException
      Throw
    End Try
  End Function

End Module
