Option Explicit On
Imports Microsoft.Data.Sqlite

Friend Class SQLiteConnectionWrapper

  Private _conn As SqliteConnection
  Private _isOpen As Boolean
  Private _path As String

  ' ------------------------------------------------------------
  '  Open
  ' ------------------------------------------------------------
  Friend Sub Open(databasePath As String)
    If String.IsNullOrWhiteSpace(databasePath) Then
      Throw New ArgumentException("Database path cannot be empty.", NameOf(databasePath))
    End If

    _path = databasePath

    ' SQLite connection string
    Dim cs As String = $"Data Source={databasePath};Foreign Keys=True;"

    _conn = New SqliteConnection(cs)

    Try
      _conn.Open()
      _isOpen = True
    Catch ex As Exception
      Throw New InvalidOperationException(
          $"Failed to open SQLite database: {databasePath}",
          ex
      )
    End Try
  End Sub

  ' ------------------------------------------------------------
  '  Close
  ' ------------------------------------------------------------
  Friend Sub Close()
    If _isOpen Then
      Try
        SqliteConnection.ClearPool(_conn)
      Catch
      End Try

      Try
        _conn.Close()
        _conn.Dispose()
      Catch
      End Try

      _conn = Nothing
      _isOpen = False
    End If
  End Sub

  ' ------------------------------------------------------------
  '  Properties
  ' ------------------------------------------------------------
  Friend ReadOnly Property IsOpen As Boolean
    Get
      Return _isOpen
    End Get
  End Property

  Friend ReadOnly Property DatabasePath As String
    Get
      Return _path
    End Get
  End Property

  ' ------------------------------------------------------------
  '  Connection
  ' ------------------------------------------------------------
  Friend ReadOnly Property InnerConnection As SqliteConnection
    Get
      Return _conn
    End Get
  End Property

  ' ------------------------------------------------------------
  '  Execute (non-query)
  ' ------------------------------------------------------------
  Friend Sub Execute(sqlText As String)
    Using cmd As New SqliteCommand(sqlText, _conn)
      cmd.ExecuteNonQuery()
    End Using
  End Sub

  ' ------------------------------------------------------------
  '  CreateCommand
  ' ------------------------------------------------------------
  Friend Function CreateCommand(sqlText As String) As SQLiteCommandWrapper
    Dim cmd As New SqliteCommand(sqlText, _conn)

    ' If there is a transaction associated with the connection then assign it to the command
    If _transaction IsNot Nothing Then
      cmd.Transaction = _transaction
    End If

    ' Pre-create parameters if needed (optional)
    ' cmd.Parameters.AddWithValue("@param", DBNull.Value)

    Return New SQLiteCommandWrapper(cmd, sqlText)
  End Function

  ' ------------------------------------------------------------
  '  OpenDataSet (SELECT)
  ' ------------------------------------------------------------
  Friend Function OpenDataSet(sqlText As String) As SQLiteDataRowReader
    Dim cmd As New SqliteCommand(sqlText, _conn)
    Dim reader As SqliteDataReader = cmd.ExecuteReader()
    Return New SQLiteDataRowReader(reader, cmd)
  End Function

  ' ------------------------------------------------------------
  '  Transactions
  ' ------------------------------------------------------------
  Private _transaction As SqliteTransaction

  Friend Sub BeginTransaction()
    If _transaction IsNot Nothing Then
      Throw New InvalidOperationException("Transaction already in progress.")
    End If

    _transaction = _conn.BeginTransaction()
  End Sub

  Friend Sub Commit()
    If _transaction Is Nothing Then
      Throw New InvalidOperationException("No active transaction to commit.")
    End If

    _transaction.Commit()
    _transaction.Dispose()
    _transaction = Nothing
  End Sub

  Friend Sub Rollback()
    If _transaction Is Nothing Then
      Throw New InvalidOperationException("No active transaction to roll back.")
    End If

    _transaction.Rollback()
    _transaction.Dispose()
    _transaction = Nothing
  End Sub

End Class
