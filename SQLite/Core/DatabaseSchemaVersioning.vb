Option Explicit On
Imports Microsoft.Data.Sqlite

Module DatabaseSchemaVersioning

  ' ============================================================
  '  Purpose:
  '    - Read SchemaVersion from tblMetadata.
  '    - Compare it to the add-in's supported schema version.
  '    - Enforce compatibility for existing databases.
  '
  '  Notes:
  '    - Assumes tblMetadata exists in databases created by DatabaseCreator.
  '    - Assumes a row with Name = 'SchemaVersion' is present.
  ' ============================================================


  ' ------------------------------------------------------------
  '  GetDatabaseSchemaVersion
  ' ------------------------------------------------------------
  ''' <summary>
  ''' Gets the schema version string stored in tblMetadata.
  ''' </summary>
  ''' <param name="conn">Open SQLiteConnectionWrapper.</param>
  ''' <returns>
  '''   The schema version string (for example, "1.0" or "1.2.3").
  '''   Returns Nothing if the row is missing.
  ''' </returns>
  Friend Function GetDatabaseSchemaVersion(ByVal conn As SQLiteConnectionWrapper) As String

    If conn Is Nothing Then
      Throw New ArgumentNullException(NameOf(conn), "Connection wrapper cannot be null.")
    End If

    ' === Variable declarations ===
    Dim sql As String
    Dim cmd As SqliteCommand = Nothing
    Dim result As Object
    Dim version As String

    Try
      sql = "SELECT Value FROM tblMetadata WHERE Name = 'SchemaVersion';"

      cmd = New SqliteCommand(sql, conn.InnerConnection)
      result = cmd.ExecuteScalar()

      If (result IsNot Nothing) AndAlso (Not Convert.IsDBNull(result)) Then
        version = Convert.ToString(result)
      Else
        version = Nothing
      End If

      Return version

    Catch ex As Exception
      ErrorHandler.UnHandleError(ex)
      Return Nothing

    Finally
      If cmd IsNot Nothing Then
        cmd.Dispose()
        cmd = Nothing
      End If

      sql = Nothing
      result = Nothing
      version = Nothing
    End Try

  End Function


  ' ------------------------------------------------------------
  '  EnsureDatabaseSchemaCompatible
  ' ------------------------------------------------------------
  ''' <summary>
  ''' Ensures that the database schema version stored in tblMetadata
  ''' is compatible with the add-in's supported schema version.
  ''' Compatibility rule:
  '''   - Major and Minor must match exactly (structural).
  '''   - Patch differences are allowed (non-structural).
  ''' </summary>
  ''' <param name="conn">Open SQLiteConnectionWrapper with an existing schema.</param>
  Friend Sub EnsureDatabaseSchemaCompatible(ByVal conn As SQLiteConnectionWrapper)

      If conn Is Nothing Then
        Throw New ArgumentNullException(NameOf(conn), "Connection wrapper cannot be null.")
      End If

      ' === Variable declarations ===
      Dim dbVersionString As String
      Dim dbVersion As Version
      Dim supportedVersionString As String
      Dim supportedVersion As Version

      Try
      ' ------------------------------------------------------------
      '  1. Get supported schema version from resources
      ' ------------------------------------------------------------
      supportedVersionString = My.Resources.SupportedSchemaVersion

      If String.IsNullOrWhiteSpace(supportedVersionString) Then
          Throw New InvalidOperationException(
          "Supported schema version (My.Resources.SupportSchemaVersion) is not configured.")
        End If

        ' ------------------------------------------------------------
        '  2. Read version from database
        ' ------------------------------------------------------------
        dbVersionString = GetDatabaseSchemaVersion(conn)

        If String.IsNullOrWhiteSpace(dbVersionString) Then
          Throw New InvalidOperationException(
          "The database does not contain a SchemaVersion entry in tblMetadata.")
        End If

        ' ------------------------------------------------------------
        '  3. Parse both versions
        ' ------------------------------------------------------------
        dbVersion = Version.Parse(dbVersionString)
        supportedVersion = Version.Parse(supportedVersionString)

        ' ------------------------------------------------------------
        '  4. Compatibility rule:
        '       - Major must match
        '       - Minor must match
        '       - Patch differences are allowed
        ' ------------------------------------------------------------
        If (dbVersion.Major <> supportedVersion.Major) OrElse
         (dbVersion.Minor <> supportedVersion.Minor) Then

          Throw New InvalidOperationException(
          "Database schema version " & dbVersion.ToString() &
          " is not compatible with supported schema version " &
          supportedVersion.ToString() & ".")
        End If

        ' Patch differences are allowed → compatible

      Catch ex As Exception
        ErrorHandler.UnHandleError(ex)
        Throw

      Finally
        dbVersionString = Nothing
        dbVersion = Nothing
        supportedVersionString = Nothing
        supportedVersion = Nothing
      End Try

    End Sub

  End Module
