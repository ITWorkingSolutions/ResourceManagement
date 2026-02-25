Option Explicit On
Imports Microsoft.Data.Sqlite

Friend Class SQLiteTransactionWrapper

  Private ReadOnly _conn As SQLiteConnectionWrapper
  Private _transaction As SqliteTransaction
  Private _active As Boolean

  ' ------------------------------------------------------------
  ' Constructor
  ' Called by SQLiteConnectionWrapper.BeginTransaction()
  ' ------------------------------------------------------------
  Friend Sub New(conn As SQLiteConnectionWrapper)
    If conn Is Nothing Then
      Throw New ArgumentNullException(NameOf(conn))
    End If

    _conn = conn
    _transaction = conn.InnerConnection.BeginTransaction()
    _active = True
  End Sub

  ' ------------------------------------------------------------
  ' Commit
  ' ------------------------------------------------------------
  Friend Sub Commit()
    If Not _active Then Exit Sub

    _transaction.Commit()
    _transaction.Dispose()
    _transaction = Nothing
    _active = False
  End Sub

  ' ------------------------------------------------------------
  ' Rollback
  ' ------------------------------------------------------------
  Friend Sub Rollback()
    If Not _active Then Exit Sub

    _transaction.Rollback()
    _transaction.Dispose()
    _transaction = Nothing
    _active = False
  End Sub

  ' ------------------------------------------------------------
  ' Active
  ' ------------------------------------------------------------
  Friend ReadOnly Property Active As Boolean
    Get
      Return _active
    End Get
  End Property

End Class
