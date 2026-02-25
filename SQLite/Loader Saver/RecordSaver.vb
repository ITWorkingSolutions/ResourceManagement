Option Explicit On
Imports System.Reflection



' ============================================================
'  RecordSaver
'  Purpose:
'     Provides generic routines for saving records using the
'     ISQLiteRecord contract and SQLiteConnectionWrapper.
'     Handles INSERT, UPDATE, DELETE based on lifecycle flags.
' ============================================================
Friend Module RecordSaver

  ' ============================================================
  '  SaveRecord
  '  Purpose:
  '     Saves a single record based on lifecycle flags.
  '
  '  Parameters:
  '     conn - SQLiteConnectionWrapper used to execute commands.
  '     rec  - The record to save.
  '
  '  Returns:
  '     Nothing.
  ' ============================================================
  Friend Sub SaveRecord(ByVal conn As SQLiteConnectionWrapper,
                        ByVal rec As ISQLiteRecord)

    Try
      ' ------------------------------------------------------------
      '  Deleted → DELETE
      ' ------------------------------------------------------------
      If rec.IsDeleted AndAlso Not rec.IsNew Then
        DeleteRecord(conn, rec)
        Return
      End If

      ' ------------------------------------------------------------
      '  New → INSERT
      ' ------------------------------------------------------------
      If rec.IsNew AndAlso Not rec.IsDeleted Then
        InsertRecord(conn, rec)
        rec.IsNew = False
        rec.IsDirty = False
        Return
      End If

      ' ------------------------------------------------------------
      '  Dirty → UPDATE
      ' ------------------------------------------------------------
      If rec.IsDirty AndAlso Not rec.IsNew AndAlso Not rec.IsDeleted Then
        UpdateRecord(conn, rec)
        rec.IsDirty = False
        Return
      End If

    Catch ex As Exception
      Throw New Exception("RecordSaver.SaveRecord failed: " & ex.Message, ex)
    End Try

  End Sub



  ' ============================================================
  '  SaveRecords
  '  Purpose:
  '     Saves a collection of records inside a single transaction.
  '
  '  Parameters:
  '     conn    - SQLiteConnectionWrapper used to execute commands.
  '     records - Collection of ISQLiteRecord instances to save.
  '
  '  Returns:
  '     Nothing.
  ' ============================================================
  Friend Sub SaveRecords(ByVal conn As SQLiteConnectionWrapper,
                         ByVal records As IEnumerable(Of ISQLiteRecord))

    Dim rec As ISQLiteRecord

    Try
      ' ------------------------------------------------------------
      '  Begin transaction
      ' ------------------------------------------------------------
      conn.BeginTransaction()

      ' ------------------------------------------------------------
      '  Save each record
      ' ------------------------------------------------------------
      For Each rec In records
        SaveRecord(conn, rec)
      Next

      ' ------------------------------------------------------------
      '  Commit
      ' ------------------------------------------------------------
      conn.Commit()

    Catch ex As Exception
      ' ------------------------------------------------------------
      '  Rollback on error
      ' ------------------------------------------------------------
      conn.Rollback()
      Throw New Exception("RecordSaver.SaveRecords failed: " & ex.Message, ex)
    End Try

  End Sub



  ' ============================================================
  '  InsertRecord
  '  Purpose:
  '     Inserts a new record into the database.
  '
  '  Parameters:
  '     conn - SQLiteConnectionWrapper used to execute commands.
  '     rec  - The record to insert.
  '
  '  Returns:
  '     Nothing.
  ' ============================================================
  Private Sub InsertRecord(ByVal conn As SQLiteConnectionWrapper,
                           ByVal rec As ISQLiteRecord)

    ' ------------------------------------------------------------
    '  Declarations
    ' ------------------------------------------------------------
    Dim tableName As String
    Dim fields() As String
    Dim field As String
    Dim sql As String
    Dim cmd As SQLiteCommandWrapper
    Dim value As Object
    Dim fieldList As String
    Dim paramList As String

    Try
      ' ------------------------------------------------------------
      '  Build SQL
      ' ------------------------------------------------------------
      tableName = rec.TableName()
      fields = rec.FieldNames()

      fieldList = String.Join(", ", fields)
      paramList = BuildParamList(fields)

      sql = "INSERT INTO " & tableName &
            " (" & fieldList & ") VALUES (" & paramList & ")"

      cmd = conn.CreateCommand(sql)
      cmd.AddParameters(fields) ' make sure the parameters exist in the command

      ' ------------------------------------------------------------
      '  Bind parameters
      ' ------------------------------------------------------------
      For Each field In fields
        Dim prop = rec.GetType().GetProperty(field,
          BindingFlags.Instance Or BindingFlags.Public Or BindingFlags.NonPublic)
        value = prop.GetValue(rec)
        'value = rec.GetType().GetProperty(field).GetValue(rec)
        cmd.SetParameter(field, If(value Is Nothing, DBNull.Value, value))
      Next

      ' ------------------------------------------------------------
      '  Execute
      ' ------------------------------------------------------------
      cmd.Execute()

    Catch ex As Exception
      Throw New Exception("RecordSaver.InsertRecord failed: " & ex.Message, ex)
    End Try

  End Sub



  ' ============================================================
  '  UpdateRecord
  '  Purpose:
  '     Updates an existing record in the database.
  '
  '  Parameters:
  '     conn - SQLiteConnectionWrapper used to execute commands.
  '     rec  - The record to update.
  '
  '  Returns:
  '     Nothing.
  ' ============================================================
  Private Sub UpdateRecord(ByVal conn As SQLiteConnectionWrapper,
                           ByVal rec As ISQLiteRecord)

    ' ------------------------------------------------------------
    '  Declarations
    ' ------------------------------------------------------------
    Dim tableName As String
    Dim fields() As String
    Dim pkNames() As String
    Dim sql As String
    Dim cmd As SQLiteCommandWrapper
    Dim field As String
    Dim pk As String
    Dim value As Object
    Dim setParts As New List(Of String)
    Dim whereParts As New List(Of String)

    Try
      ' ------------------------------------------------------------
      '  Build SQL
      ' ------------------------------------------------------------
      tableName = rec.TableName()
      fields = rec.FieldNames()
      pkNames = rec.PrimaryKeyNames()

      ' SET clause
      For Each field In fields
        If Not pkNames.Contains(field) Then
          setParts.Add(field & " = @" & field)
        End If
      Next

      ' WHERE clause
      For Each pk In pkNames
        whereParts.Add(pk & " = @PK_" & pk)
      Next

      sql = "UPDATE " & tableName &
            " SET " & String.Join(", ", setParts) &
            " WHERE " & String.Join(" AND ", whereParts)

      cmd = conn.CreateCommand(sql)
      ' Add parameters for SET clause
      cmd.AddParameters(fields)

      ' Add parameters for WHERE clause
      Dim pkParamNames = pkNames.Select(Function(pkName) "PK_" & pkName).ToArray()
      cmd.AddParameters(pkParamNames)

      ' ------------------------------------------------------------
      '  Bind non-PK fields
      ' ------------------------------------------------------------
      For Each field In fields
        Dim prop = rec.GetType().GetProperty(field,
          BindingFlags.Instance Or BindingFlags.Public Or BindingFlags.NonPublic)
        value = prop.GetValue(rec)
        'value = rec.GetType().GetProperty(field).GetValue(rec)
        cmd.SetParameter(field, If(value Is Nothing, DBNull.Value, value))
      Next

      ' ------------------------------------------------------------
      '  Bind PK fields
      ' ------------------------------------------------------------
      For Each pk In pkNames
        Dim prop = rec.GetType().GetProperty(pk,
          BindingFlags.Instance Or BindingFlags.Public Or BindingFlags.NonPublic)
        value = prop.GetValue(rec)
        'value = rec.GetType().GetProperty(pk).GetValue(rec)
        cmd.SetParameter("PK_" & pk, If(value Is Nothing, DBNull.Value, value))
      Next

      ' ------------------------------------------------------------
      '  Execute
      ' ------------------------------------------------------------
      cmd.Execute()

    Catch ex As Exception
      Throw New Exception("RecordSaver.UpdateRecord failed: " & ex.Message, ex)
    End Try

  End Sub



  ' ============================================================
  '  DeleteRecord
  '  Purpose:
  '     Deletes a record from the database.
  '
  '  Parameters:
  '     conn - SQLiteConnectionWrapper used to execute commands.
  '     rec  - The record to delete.
  '
  '  Returns:
  '     Nothing.
  ' ============================================================
  Private Sub DeleteRecord(ByVal conn As SQLiteConnectionWrapper,
                           ByVal rec As ISQLiteRecord)

    ' ------------------------------------------------------------
    '  Declarations
    ' ------------------------------------------------------------
    Dim tableName As String
    Dim pkNames() As String
    Dim sql As String
    Dim cmd As SQLiteCommandWrapper
    Dim pk As String
    Dim value As Object
    Dim whereParts As New List(Of String)

    Try
      ' ------------------------------------------------------------
      '  Build SQL
      ' ------------------------------------------------------------
      tableName = rec.TableName()
      pkNames = rec.PrimaryKeyNames()

      For Each pk In pkNames
        whereParts.Add(pk & " = @" & pk)
      Next

      sql = "DELETE FROM " & tableName &
            " WHERE " & String.Join(" AND ", whereParts)

      cmd = conn.CreateCommand(sql)

      cmd.AddParameters(pkNames) ' make sure the parameters exist in the command

      ' ------------------------------------------------------------
      '  Bind PK fields
      ' ------------------------------------------------------------
      For Each pk In pkNames
        Dim prop = rec.GetType().GetProperty(pk,
          BindingFlags.Instance Or BindingFlags.Public Or BindingFlags.NonPublic)
        value = prop.GetValue(rec)
        'value = rec.GetType().GetProperty(pk).GetValue(rec)
        cmd.SetParameter(pk, If(value Is Nothing, DBNull.Value, value))
      Next

      ' ------------------------------------------------------------
      '  Execute
      ' ------------------------------------------------------------
      cmd.Execute()

    Catch ex As Exception
      Throw New Exception("RecordSaver.DeleteRecord failed: " & ex.Message, ex)
    End Try

  End Sub



  ' ============================================================
  '  BuildParamList
  '  Purpose:
  '     Builds a comma-separated list of @param names.
  ' ============================================================
  Private Function BuildParamList(ByVal fields() As String) As String

    Dim i As Integer
    Dim parts(fields.Length - 1) As String

    For i = 0 To fields.Length - 1
      parts(i) = "@" & fields(i)
    Next

    Return String.Join(", ", parts)

  End Function

End Module