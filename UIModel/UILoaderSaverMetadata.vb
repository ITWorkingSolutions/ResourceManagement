Option Explicit On

' ==========================================================================================
' Module: UILoaderSaverMetadata
' Purpose:
'   Provide UI-safe access to application metadata stored in tblMetadata.
'
'   - NO direct SQL. All database interaction is via:
'       * OpenDatabase
'       * RecordLoader
'       * RecordSaver
'
'   - Supports:
'       * LoadMetadataValue(name)
'       * SaveMetadataValue(name, value)
'
'   - Metadata is global and may be used by any UI form (Resource, Availability, etc.).
' ==========================================================================================
Friend Module UILoaderSaverMetadata

  ' ==========================================================================================
  ' Routine: LoadMetadataValue
  ' Purpose:
  '   Load a single metadata value by name.
  '
  ' Parameters:
  '   name - Metadata key (e.g., "DefaultMode")
  '
  ' Returns:
  '   String - Metadata value, or "" if not found or deleted.
  '
  ' Notes:
  '   - Uses RecordLoader only.
  '   - Safe for use from any UI form.
  ' ==========================================================================================
  Friend Function LoadMetadataValue(ByVal name As String) As String

    Dim conn As SQLiteConnectionWrapper = Nothing

    Try
      conn = OpenDatabase(AddInContext.Current.Config.ActiveDbPath)

      Dim rec As RecordMetadata =
        RecordLoader.LoadRecord(Of RecordMetadata)(conn, {name})

      If rec Is Nothing OrElse rec.IsDeleted Then
        Return ""
      End If

      Return rec.Value

    Catch ex As Exception
      ErrorHandler.UnHandleError(ex)
      Return ""

    Finally
      If conn IsNot Nothing Then conn.Close()
    End Try

  End Function


  ' ==========================================================================================
  ' Routine: SaveMetadataValue
  ' Purpose:
  '   Insert or update a metadata value.
  '
  ' Parameters:
  '   name  - Metadata key
  '   value - Metadata value
  '
  ' Returns:
  '   None
  '
  ' Notes:
  '   - Uses RecordLoader + RecordSaver.
  '   - Creates new record if not found.
  ' ==========================================================================================
  Friend Sub SaveMetadataValue(ByVal name As String, ByVal value As String)

    Dim conn As SQLiteConnectionWrapper = Nothing

    Try
      conn = OpenDatabase(AddInContext.Current.Config.ActiveDbPath)

      Dim rec As RecordMetadata =
        RecordLoader.LoadRecord(Of RecordMetadata)(conn, {name})

      If rec Is Nothing Then
        rec = New RecordMetadata()
        rec.Name = name
        rec.Value = value
        rec.IsNew = True
      Else
        rec.Value = value
        rec.IsDirty = True
      End If

      RecordSaver.SaveRecord(conn, rec)

    Catch ex As Exception
      ErrorHandler.UnHandleError(ex)
      Throw

    Finally
      If conn IsNot Nothing Then conn.Close()
    End Try

  End Sub

End Module
