Imports System.IO
Module DatabaseOpen

  ' ============================================================
  '  Routine: OpenDatabase
  '  Purpose:
  '       Opens an existing SQLite database or creates a new one.
  '       Ensures schema compatibility before returning the connection.
  '
  '  Parameters:
  '       dbPath - The full path to the SQLite database file.
  '
  '  Returns:
  '       SQLiteConnectionWrapper - A fully validated, ready-to-use connection.
  '
  '  Notes:
  '       - dbPath MUST be provided by AddInContext.Current.Config.ActiveDbPath.
  '       - If no active database is configured, this routine throws.
  ' ============================================================
  Friend Function OpenDatabase(ByVal dbPath As String) As SQLiteConnectionWrapper

    ' === Variable declarations ===
    Dim conn As SQLiteConnectionWrapper
    Dim existedBefore As Boolean = File.Exists(dbPath)

    Try
      ' ------------------------------------------------------------
      '  Validate database path
      ' ------------------------------------------------------------
      If String.IsNullOrWhiteSpace(dbPath) Then
        Throw New InvalidOperationException(
        "No active database is configured. Please select or create a database in Settings.")
      End If

      conn = New SQLiteConnectionWrapper()
      conn.Open(dbPath)

      ' ------------------------------------------------------------
      '  Early validation: ensure file is a real SQLite DB
      ' ------------------------------------------------------------
      Try
        Dim cmdWrapper = conn.CreateCommand("PRAGMA schema_version;")
        cmdWrapper.InnerCommand.ExecuteScalar()
      Catch
        If existedBefore Then
          Throw New InvalidOperationException(
                    "The selected file exists but is not a valid SQLite database.")
        Else
          Throw New InvalidOperationException(
                    "The selected path is not a valid SQLite database file.")
        End If
      End Try

      ' ------------------------------------------------------------
      '  Determine if database is new (no user tables)
      ' ------------------------------------------------------------
      Dim isNew As Boolean = IsNewDatabase(conn)

      If isNew Then
        ' Create schema using latest manifest
        DatabaseCreator.CreateNewDatabase(conn)
      Else
        ' Validate schema version (Major/Minor must match)
        DatabaseSchemaVersioning.EnsureDatabaseSchemaCompatible(conn)
      End If

      Return conn

    Catch ex As Exception
      ErrorHandler.UnHandleError(ex)
      Throw
    End Try

  End Function

  ' ============================================================
  '  Routine: IsNewDatabase
  '  Purpose:
  '       Determines whether the SQLite database contains any
  '       user-defined tables. A database is considered NEW if it
  '       contains no tables other than SQLite's internal tables.
  '
  '  Parameters:
  '       conn - An open SQLiteConnectionWrapper instance.
  '
  '  Returns:
  '       Boolean - True if the database has no user tables.
  '                 False if at least one user table exists.
  '
  '  Notes:
  '       - Internal SQLite tables (sqlite_*) are ignored.
  '       - This routine must be called immediately after opening
  '         the database file, before any schema-dependent logic.
  ' ============================================================
  Friend Function IsNewDatabase(ByVal conn As SQLiteConnectionWrapper) As Boolean

    If conn Is Nothing Then
      Throw New ArgumentNullException(NameOf(conn), "Connection wrapper cannot be null.")
    End If

    ' === Variable declarations ===
    Dim sql As String
    Dim reader As SQLiteDataRowReader = Nothing
    Dim hasTables As Boolean

    Try
      ' ------------------------------------------------------------
      '  Query sqlite_master for user tables.
      '  A database is considered NEW if it contains no user tables.
      ' ------------------------------------------------------------
      sql = "SELECT name FROM sqlite_master " &
            "WHERE type = 'table' AND name NOT LIKE 'sqlite_%';"

      reader = conn.OpenDataSet(sql)

      ' If reader.Read returns True, at least one table exists
      hasTables = reader.Read()

      ' New database = no tables found
      Return (Not hasTables)

    Catch ex As Exception
      ErrorHandler.UnHandleError(ex)
      Throw

    Finally
      If reader IsNot Nothing Then
        reader.Dispose()
        reader = Nothing
      End If

      sql = Nothing
      hasTables = Nothing
    End Try

  End Function
End Module
