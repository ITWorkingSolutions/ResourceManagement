Option Explicit On
Imports Microsoft.Data.Sqlite

Friend Class SQLiteCommandWrapper

  Private ReadOnly _cmd As SqliteCommand
  Private ReadOnly _sql As String

  ' ------------------------------------------------------------
  ' Constructor
  ' ------------------------------------------------------------
  Friend Sub New(innerCmd As SqliteCommand, sqlText As String)
    _cmd = innerCmd
    _sql = sqlText
  End Sub

  ' ------------------------------------------------------------
  ' SQL text (read-only)
  ' ------------------------------------------------------------
  Friend ReadOnly Property SqlText As String
    Get
      Return _sql
    End Get
  End Property

  ' ------------------------------------------------------------
  ' AddParameters
  ' Adds named parameters to the command before binding values.
  ' Uses AddWithValue with DBNull.Value placeholder so SetParameter
  ' can simply update .Value later.
  ' ------------------------------------------------------------
  Friend Sub AddParameters(fieldNames As IEnumerable(Of String))
    If fieldNames Is Nothing Then Return

    For Each name In fieldNames
      Dim fullName As String = "@" & name
      If Not _cmd.Parameters.Contains(fullName) Then
        ' Add a parameter placeholder; use DBNull.Value until a real value is set.
        _cmd.Parameters.AddWithValue(fullName, DBNull.Value)
      End If
    Next
  End Sub

  ' ------------------------------------------------------------
  ' SetParameter
  '
  ' Example:
  '   cmd.SetParameter("oid", 123)
  '
  ' Internally:
  '   binds @oid = 123
  ' ------------------------------------------------------------
  Friend Sub SetParameter(paramName As String, value As Object)

    Dim fullName As String = "@" & paramName

    If Not _cmd.Parameters.Contains(fullName) Then
      Throw New ArgumentException(
          $"Parameter not found: {fullName}",
          NameOf(paramName)
      )
    End If

    _cmd.Parameters(fullName).Value = If(value Is Nothing, DBNull.Value, value)
  End Sub

  ' ------------------------------------------------------------
  ' Execute (non-query)
  ' ------------------------------------------------------------
  Friend Sub Execute()
    _cmd.ExecuteNonQuery()
  End Sub

  ' ------------------------------------------------------------
  ' Execute SELECT and return a reader wrapper
  ' ------------------------------------------------------------
  Friend Function OpenDataSet() As SQLiteDataRowReader
    Dim reader As SqliteDataReader = _cmd.ExecuteReader()
    Return New SQLiteDataRowReader(reader)
  End Function

  ' ------------------------------------------------------------
  ' Expose underlying command (if needed)
  ' ------------------------------------------------------------
  Friend ReadOnly Property InnerCommand As SqliteCommand
    Get
      Return _cmd
    End Get
  End Property

End Class
