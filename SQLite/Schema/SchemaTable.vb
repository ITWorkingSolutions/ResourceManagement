Option Explicit On

Friend Class SchemaTable

  ' ------------------------------------------------------------
  '  Properties
  ' ------------------------------------------------------------
  Friend Property Name As String
  Friend Property ClassName As String

  ' Strongly typed lists
  Friend Property Fields As List(Of SchemaField)
  Friend Property ForeignKeys As List(Of SchemaForeignKey)

  ' ------------------------------------------------------------
  '  Constructor
  ' ------------------------------------------------------------
  Friend Sub New()
    Fields = New List(Of SchemaField)()
    ForeignKeys = New List(Of SchemaForeignKey)()
  End Sub

  ' ------------------------------------------------------------
  '  Add a field
  ' ------------------------------------------------------------
  Friend Sub AddField(fld As SchemaField)
    If fld Is Nothing Then Exit Sub
    Fields.Add(fld)
  End Sub

  ' ------------------------------------------------------------
  '  Lookup a field by name
  ' ------------------------------------------------------------
  Friend Function GetField(name As String) As SchemaField
    If String.IsNullOrWhiteSpace(name) Then Return Nothing

    Return Fields.
        FirstOrDefault(Function(f) f.Name.Equals(name, StringComparison.OrdinalIgnoreCase))
  End Function

  ' ------------------------------------------------------------
  '  Add a foreign key
  ' ------------------------------------------------------------
  Friend Sub AddForeignKey(fk As SchemaForeignKey)
    If fk Is Nothing Then Exit Sub
    ForeignKeys.Add(fk)
  End Sub

  ' ------------------------------------------------------------
  '  Get all primary key fields
  ' ------------------------------------------------------------
  Friend Function PrimaryKeyFields() As List(Of SchemaField)
    Return Fields.
        Where(Function(f) f.IsPrimaryKey).
        ToList()
  End Function

End Class
