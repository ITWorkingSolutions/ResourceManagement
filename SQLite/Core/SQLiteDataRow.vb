Option Explicit On
Imports System.Collections.Generic

Friend Class SQLiteDataRow

  Private ReadOnly _values As Dictionary(Of String, Object)

  ' ------------------------------------------------------------
  ' Constructor
  ' ------------------------------------------------------------
  Friend Sub New()
    ' Case-insensitive key comparison (equivalent to TextCompare)
    _values = New Dictionary(Of String, Object)(
        StringComparer.OrdinalIgnoreCase
    )
  End Sub

  ' ------------------------------------------------------------
  ' AddField (internal use by row reader)
  ' ------------------------------------------------------------
  Friend Sub AddField(columnName As String, value As Object)
    _values(columnName) = value
  End Sub

  ' ------------------------------------------------------------
  ' GetValue
  ' Returns Nothing if field does not exist (equivalent to Null)
  ' ------------------------------------------------------------
  Friend Function GetValue(columnName As String) As Object
    If _values.ContainsKey(columnName) Then
      Return _values(columnName)
    Else
      Return Nothing
    End If
  End Function

  ' ------------------------------------------------------------
  ' HasField
  ' ------------------------------------------------------------
  Friend Function HasField(columnName As String) As Boolean
    Return _values.ContainsKey(columnName)
  End Function

  ' ------------------------------------------------------------
  ' FieldNames
  ' Returns the list of column names in insertion order
  ' ------------------------------------------------------------
  Friend ReadOnly Property FieldNames As IReadOnlyList(Of String)
    Get
      Return _values.Keys.ToList()
    End Get
  End Property

  ' ------------------------------------------------------------
  ' Default property (indexer)
  ' Allows row("ColumnName") syntax
  ' ------------------------------------------------------------
  Default Friend ReadOnly Property Item(columnName As String) As Object
    Get
      Return GetValue(columnName)
    End Get
  End Property
End Class
