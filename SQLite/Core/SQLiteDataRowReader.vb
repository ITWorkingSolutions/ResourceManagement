Option Explicit On
Imports Microsoft.Data.Sqlite

Friend Class SQLiteDataRowReader
  Implements IDisposable

  Private ReadOnly _reader As SqliteDataReader
  Private ReadOnly _command As SqliteCommand
  Private _currentRow As SQLiteDataRow

  Friend Sub New(reader As SqliteDataReader, command As SqliteCommand)
    _reader = reader
    _command = command
  End Sub

  ' ------------------------------------------------------------
  ' Move to next row
  ' ------------------------------------------------------------
  Friend Function Read() As Boolean
    If _reader.Read() Then
      _currentRow = New SQLiteDataRow()

      For i = 0 To _reader.FieldCount - 1
        Dim name = _reader.GetName(i)
        Dim value As Object =
            If(_reader.IsDBNull(i), Nothing, _reader.GetValue(i))

        _currentRow.AddField(name, value)
      Next

      Return True
    End If

    Return False
  End Function

  ' ------------------------------------------------------------
  ' Current row
  ' ------------------------------------------------------------
  Friend ReadOnly Property Row As SQLiteDataRow
    Get
      Return _currentRow
    End Get
  End Property

  ' ------------------------------------------------------------
  ' End-of-data indicator
  ' ------------------------------------------------------------
  Friend ReadOnly Property EOF As Boolean
    Get
      Return _reader Is Nothing OrElse _reader.IsClosed OrElse Not _reader.HasRows
    End Get
  End Property

  ' ------------------------------------------------------------
  ' Cleanup
  ' ------------------------------------------------------------
  Friend Sub Dispose() Implements IDisposable.Dispose
    _reader?.Close()
    _reader?.Dispose()
    _command?.Dispose()
  End Sub


End Class
