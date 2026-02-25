Option Explicit On
Imports System.Linq.Expressions
Imports System.Reflection
Imports System.Windows.Forms

Module DatabaseCreator

  ' ============================================================
  '  Purpose:
  '    Create a brand-new SQLite database using the schema manifest.
  '
  '    Lifecycle:
  '      1. Load embedded JSON schema manifest.
  '      2. Deserialize to SchemaManifest / SchemaTable objects.
  '      3. Resolve table creation order (topological sort).
  '      4. Generate CREATE TABLE SQL for each table.
  '      5. Execute SQL statements in dependency-safe order.
  '      6. Insert SchemaVersion row into tblMetadata.
  '
  '    This is the ONLY supported way to create a new database.
  ' ============================================================

  ''' <summary>
  ''' Create a new SQLite database schema in the target connection.
  ''' The connection should point to an empty or newly created file,
  ''' with PRAGMA foreign_keys = ON already configured by the wrapper.
  ''' </summary>
  ''' <param name="conn">Open SQLiteConnectionWrapper pointing at the target database.</param>
  Friend Sub CreateNewDatabase(conn As SQLiteConnectionWrapper)

    ' === Variable declarations ===
    Dim manifest As SchemaManifest = Nothing
    Dim orderedTables As List(Of SchemaTable) = Nothing
    Dim tbl As SchemaTable = Nothing
    Dim sql As String = Nothing
    Dim recMetadata As RecordMetadata = Nothing

    Try
      ' === Prompt for default mode BEFORE creating anything ===
      Dim defaultMode As String = Nothing
      Using dlg As New SelectDefaultMode()
        Dim result = dlg.ShowDialog()
        If result <> DialogResult.OK Then Exit Sub
        defaultMode = dlg.SelectedMode   ' "Available" or "Unavailable"
      End Using

      ' 1. Load schema manifest (from embedded resource)
      manifest = SchemaLoader.LoadSchemaManifest()

      ' 2. Resolve table creation order (parent tables first)
      orderedTables =
        SchemaDependencyResolver.ResolveTableCreationOrder(manifest)

      ' 3. Generate and execute CREATE TABLE statements
      For Each tbl In orderedTables
        sql = SchemaSqlBuilder.GenerateCreateTableSQL(tbl)
        conn.Execute(sql)
      Next

      ' 4. Create views
      ViewCreator.CreateViews(conn, manifest)

      ' 5. Insert system ListItemTypes
      For Each sysType As RecordListItemType In ListItemTypeSystemCatalog.AllSystemTypes()
        RecordSaver.SaveRecord(conn, sysType)
      Next

      ' 6. Insert schema version record into tblMetadata
      recMetadata = New RecordMetadata With {
        .Name = "SchemaVersion",
        .Value = manifest.Version,   ' latest version from manifest
        .IsNew = True
      }
      RecordSaver.SaveRecord(conn, recMetadata)

      ' 7. Insert default mode
      Dim recDefaultMode As New RecordMetadata With {
            .Name = "DefaultMode",
            .Value = defaultMode,
            .IsNew = True
        }
      RecordSaver.SaveRecord(conn, recDefaultMode)

    Catch ex As Exception
      ErrorHandler.UnHandleError(ex)

    Finally
      manifest = Nothing
      orderedTables = Nothing
      tbl = Nothing
      sql = Nothing
      recMetadata = Nothing
    End Try

  End Sub

End Module