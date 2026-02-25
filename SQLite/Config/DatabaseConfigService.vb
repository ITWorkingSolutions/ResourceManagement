' ============================================================================================
'  Class: DatabaseConfigService
'  Purpose:
'       Provides operations for managing known database paths and the active database path.
'
'  Notes:
'       - Wraps DatabaseConfigManager.Load/Save.
'       - Ensures no duplicates.
'       - Ensures active DB is always valid.
' ============================================================================================
Friend Class DatabaseConfigService

  ' ========================================================================================
  '  Routine: AddDatabasePath
  '  Purpose:
  '       Adds a new database path to the known list and optionally sets it active.
  '
  '  Parameters:
  '       dbPath        - Full path to the database file.
  '       makeActive    - If True, sets this path as the active database.
  '
  '  Returns:
  '       None
  ' ========================================================================================
  Friend Shared Sub AddDatabasePath(ByVal dbPath As String, ByVal makeActive As Boolean)

    ' === Variable declarations ===
    Dim config As DatabaseConfig = Nothing

    Try
      config = DatabaseConfigManager.Load()

      ' --- Add if not already present (case-insensitive) ---
      If Not config.KnownDbPaths.Any(Function(p) String.Equals(p, dbPath, StringComparison.OrdinalIgnoreCase)) Then
        config.KnownDbPaths.Add(dbPath)
      End If

      ' --- Set active if requested ---
      If makeActive Then
        config.ActiveDbPath = dbPath
      End If

      DatabaseConfigManager.Save(config)

    Catch ex As Exception
      ErrorHandler.UnHandleError(ex, "DatabaseConfigService.AddDatabasePath")

    Finally
      config = Nothing
    End Try

  End Sub

  ' ========================================================================================
  '  Routine: RemoveDatabasePath
  '  Purpose:
  '       Removes a database path from the known list. If it is the active database,
  '       clears the active path or switches to another known path.
  '
  '  Parameters:
  '       dbPath - Full path to remove.
  '
  '  Returns:
  '       None
  ' ========================================================================================
  Friend Shared Sub RemoveDatabasePath(ByVal dbPath As String)

    ' === Variable declarations ===
    Dim config As DatabaseConfig = Nothing

    Try
      config = DatabaseConfigManager.Load()

      ' --- Remove path (case-insensitive) ---
      config.KnownDbPaths =
          config.KnownDbPaths.
              Where(Function(p) Not String.Equals(p, dbPath, StringComparison.OrdinalIgnoreCase)).
              ToList()

      ' --- If this was the active DB, clear or switch ---
      If String.Equals(config.ActiveDbPath, dbPath, StringComparison.OrdinalIgnoreCase) Then

        If config.KnownDbPaths.Count > 0 Then
          ' Switch to first known DB
          config.ActiveDbPath = config.KnownDbPaths(0)
        Else
          ' No known DBs left
          config.ActiveDbPath = Nothing
        End If

      End If

      DatabaseConfigManager.Save(config)

    Catch ex As Exception
      ErrorHandler.UnHandleError(ex, "DatabaseConfigService.RemoveDatabasePath")

    Finally
      config = Nothing
    End Try

  End Sub

  ' ========================================================================================
  '  Routine: SetActiveDatabase
  '  Purpose:
  '       Sets the active database path and ensures it is in the known list.
  '
  '  Parameters:
  '       dbPath - Full path to set as active.
  '
  '  Returns:
  '       None
  ' ========================================================================================
  Friend Shared Sub SetActiveDatabase(ByVal dbPath As String)

    ' === Variable declarations ===
    Dim config As DatabaseConfig = Nothing

    Try
      config = DatabaseConfigManager.Load()

      ' --- Ensure path is known (case-insensitive) ---
      If Not config.KnownDbPaths.Any(Function(p) String.Equals(p, dbPath, StringComparison.OrdinalIgnoreCase)) Then
        config.KnownDbPaths.Add(dbPath)
      End If

      ' --- Set active ---
      config.ActiveDbPath = dbPath

      DatabaseConfigManager.Save(config)

    Catch ex As Exception
      ErrorHandler.UnHandleError(ex, "DatabaseConfigService.SetActiveDatabase")

    Finally
      config = Nothing
    End Try

  End Sub

End Class