Option Explicit On
Imports System.Reflection
Imports ResourceManagement.My



' ============================================================
'  RecordLoader
'  Purpose:
'     Provides generic routines for loading one or more records
'     from the database using the ISQLiteRecord contract.
' ============================================================
Friend Module RecordLoader

  ' ============================================================
  '  LoadRecords
  '  Purpose:
  '     Loads ALL rows from the table associated with T.
  '     Returns a List(Of T) with lifecycle flags reset.
  '
  '  Parameters:
  '     conn - SQLiteConnectionWrapper used to execute commands.
  '
  '  Returns:
  '     List(Of T) - All records from the table.
  ' ============================================================
  Friend Function LoadRecords(Of T As {ISQLiteRecord, New})(ByVal conn As SQLiteConnectionWrapper) As List(Of T)

    ' ------------------------------------------------------------
    '  Declarations
    ' ------------------------------------------------------------
    Dim proto As New T()
    Dim tableName As String
    Dim fieldNames() As String
    Dim sql As String
    Dim cmd As SQLiteCommandWrapper
    Dim ds As SQLiteDataRowReader
    Dim results As New List(Of T)
    Dim rec As T
    Dim row As SQLiteDataRow
    Dim field As String
    Dim value As Object

    Try
      ' ------------------------------------------------------------
      '  Build SQL
      ' ------------------------------------------------------------
      tableName = proto.TableName()
      fieldNames = proto.FieldNames()

      sql = "SELECT " & String.Join(", ", fieldNames) & " FROM " & tableName

      ' ------------------------------------------------------------
      '  Execute query
      ' ------------------------------------------------------------
      cmd = conn.CreateCommand(sql)
      ds = cmd.OpenDataSet()

      ' ------------------------------------------------------------
      '  Read rows
      ' ------------------------------------------------------------
      While ds.Read()
        row = ds.Row
        rec = New T()

        For Each field In rec.FieldNames()
          value = row.GetValue(field)
          Dim prop = rec.GetType().GetProperty(field,
            BindingFlags.Instance Or BindingFlags.Public Or BindingFlags.NonPublic)
          prop.SetValue(rec, value)
        Next

        rec.IsNew = False
        rec.IsDirty = False
        rec.IsDeleted = False

        results.Add(rec)
      End While

      Return results

    Catch ex As Exception
      Throw New Exception("RecordLoader.LoadRecords failed: " & ex.Message, ex)
    End Try

  End Function

  ' ============================================================================================
  '  Routine: LoadRecordsByFields
  '  Purpose:
  '       Loads all records of type T where one or more fields match specified values.
  '       Uses the ISQLiteRecord contract to determine table name and field list.
  '       Returns a List(Of T) with lifecycle flags reset.
  '
  '  Parameters:
  '       conn        - SQLiteConnectionWrapper used to execute commands.
  '       fieldNames  - Array of field names to filter on. Must exist in T.FieldNames().
  '       fieldValues - Array of values to match for the corresponding fields.
  '
  '  Returns:
  '       List(Of T) - All matching records, or an empty list if none found.
  '
  '  Notes:
  '       - fieldNames and fieldValues must be the same length.
  '       - WHERE clause is built as: f1 = @f1 AND f2 = @f2 AND ...
  '       - All records returned have IsNew = False, IsDirty = False, IsDeleted = False.
  '       - This is the multi-row analogue of LoadRecord (which uses primary keys).
  '  Examples: 
  '         All availability rows for a resource
  '         Dim avCore = LoadRecordsByFields(Of RecordResourceAvailability)(
  '                         conn,
  '                         New String() {"ResourceID"},
  '                         New Object() {resourceID})
  '          All weekly pattern rows For a given availability
  '          Dim weekly = LoadRecordsByFields(Of RecordPatternWeekly)(
  '                         conn,
  '                         New String() {"AvailabilityID"},
  '                         New Object() {availabilityID})
  '           If you ever need more complex filters:
  '          e.g., ResourceID + Mode
  '          Dim filtered = LoadRecordsByFields(Of RecordResourceAvailability)(
  '                           conn,
  '                           New String() {"ResourceID", "Mode"},
  '                           New Object() {resourceID, "Available"})
  ' ============================================================================================
  Friend Function LoadRecordsByFields(Of T As {ISQLiteRecord, New})(
        ByVal conn As SQLiteConnectionWrapper,
        ByVal fieldNames() As String,
        ByVal fieldValues() As Object) As List(Of T)

    ' ------------------------------------------------------------
    '  Declarations
    ' ------------------------------------------------------------
    Dim proto As New T()
    Dim tableName As String
    Dim allFieldNames() As String
    Dim sql As String
    Dim whereParts As New List(Of String)
    Dim cmd As SQLiteCommandWrapper
    Dim ds As SQLiteDataRowReader
    Dim results As New List(Of T)
    Dim rec As T
    Dim row As SQLiteDataRow
    Dim i As Long
    Dim f As String
    Dim value As Object

    Try
      ' ------------------------------------------------------------
      '  Validate input
      ' ------------------------------------------------------------
      If fieldNames Is Nothing OrElse fieldValues Is Nothing Then
        Throw New Exception("fieldNames and fieldValues must not be Nothing.")
      End If

      If fieldNames.Length <> fieldValues.Length Then
        Throw New Exception("fieldNames and fieldValues must have the same length.")
      End If

      allFieldNames = proto.FieldNames()

      ' Ensure all filter fields exist on the record
      For Each f In fieldNames
        If Not allFieldNames.Contains(f) Then
          Throw New Exception("Field '" & f & "' does not exist in record type " &
                                    GetType(T).Name & ".")
        End If
      Next

      ' ------------------------------------------------------------
      '  Build SQL
      ' ------------------------------------------------------------
      tableName = proto.TableName()

      For Each f In fieldNames
        whereParts.Add(f & " = @" & f)
      Next

      sql = "SELECT " & String.Join(", ", allFieldNames) &
              " FROM " & tableName

      If whereParts.Count > 0 Then
        sql &= " WHERE " & String.Join(" AND ", whereParts)
      End If

      ' ------------------------------------------------------------
      '  Prepare command
      ' ------------------------------------------------------------
      cmd = conn.CreateCommand(sql)
      cmd.AddParameters(fieldNames)   ' Pre-create parameters by name

      For i = 0 To fieldNames.Length - 1
        cmd.SetParameter(fieldNames(i), fieldValues(i))
      Next

      ' ------------------------------------------------------------
      '  Execute and read rows
      ' ------------------------------------------------------------
      ds = cmd.OpenDataSet()

      While ds.Read()
        row = ds.Row
        rec = New T()

        For Each f In rec.FieldNames()
          value = row.GetValue(f)
          Dim prop = rec.GetType().GetProperty(f,
            BindingFlags.Instance Or BindingFlags.Public Or BindingFlags.NonPublic)
          prop.SetValue(rec, value)
        Next

        rec.IsNew = False
        rec.IsDirty = False
        rec.IsDeleted = False

        results.Add(rec)
      End While

      Return results

    Catch ex As Exception
      Throw New Exception("RecordLoader.LoadRecordsByFields failed: " & ex.Message, ex)
    End Try

  End Function

  ' ============================================================
  '  LoadRecord
  '  Purpose:
  '     Loads a single record identified by its primary key(s).
  '     Returns Nothing if no matching row exists.
  '
  '  Parameters:
  '     conn     - SQLiteConnectionWrapper used to execute commands.
  '     pkValues - Array of primary key values in correct order.
  '
  '  Returns:
  '     T - The loaded record, or Nothing if not found.
  ' ============================================================
  Friend Function LoadRecord(Of T As {ISQLiteRecord, New})(ByVal conn As SQLiteConnectionWrapper,
                                                           ByVal pkValues As Object()) As T

    ' ------------------------------------------------------------
    '  Declarations
    ' ------------------------------------------------------------
    Dim proto As New T()
    Dim tableName As String
    Dim fieldNames() As String
    Dim pkNames() As String
    Dim sql As String
    Dim whereParts As New List(Of String)
    Dim cmd As SQLiteCommandWrapper
    Dim ds As SQLiteDataRowReader
    Dim rec As T
    Dim row As SQLiteDataRow
    Dim i As Long
    Dim field As String
    Dim value As Object

    Try
      ' ------------------------------------------------------------
      '  Build SQL
      ' ------------------------------------------------------------
      tableName = proto.TableName()
      fieldNames = proto.FieldNames()
      pkNames = proto.PrimaryKeyNames()

      For Each field In pkNames
        whereParts.Add(field & " = @" & field)
      Next

      sql = "SELECT " & String.Join(", ", fieldNames) &
            " FROM " & tableName &
            " WHERE " & String.Join(" AND ", whereParts)

      ' ------------------------------------------------------------
      '  Prepare command
      ' ------------------------------------------------------------
      cmd = conn.CreateCommand(sql)
      cmd.AddParameters(pkNames) ' Pre-create PK parameters

      ' Bind PK parameters
      For i = 0 To pkNames.Length - 1
        cmd.SetParameter(pkNames(i), pkValues(i))
      Next

      ' ------------------------------------------------------------
      '  Execute and read
      ' ------------------------------------------------------------
      ds = cmd.OpenDataSet()

      If Not ds.Read() Then
        Return Nothing
      End If

      row = ds.Row
      rec = New T()

      For Each field In rec.FieldNames()
        value = row.GetValue(field)
        Dim prop = rec.GetType().GetProperty(field,
            BindingFlags.Instance Or BindingFlags.Public Or BindingFlags.NonPublic)
        prop.SetValue(rec, value)
        'rec.GetType().GetProperty(field).SetValue(rec, value)
      Next

      rec.IsNew = False
      rec.IsDirty = False
      rec.IsDeleted = False

      Return rec

    Catch ex As Exception
      Throw New Exception("RecordLoader.LoadRecord failed: " & ex.Message, ex)
    End Try

  End Function

End Module