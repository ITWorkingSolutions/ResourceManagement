Option Explicit On

Public Class SchemaTable

  ' ------------------------------------------------------------
  '  Properties
  ' ------------------------------------------------------------
  Public Property Name As String
  Public Property ClassName As String

  ' Strongly typed lists
  Public Property Fields As List(Of SchemaField)
  Public Property ForeignKeys As List(Of SchemaForeignKey)

  ' ------------------------------------------------------------
  '  Constructor
  ' ------------------------------------------------------------
  Public Sub New()
    Fields = New List(Of SchemaField)()
    ForeignKeys = New List(Of SchemaForeignKey)()
  End Sub

  ' ------------------------------------------------------------
  '  Add a field
  ' ------------------------------------------------------------
  Public Sub AddField(fld As SchemaField)
    If fld Is Nothing Then Exit Sub
    Fields.Add(fld)
  End Sub

  ' ------------------------------------------------------------
  '  Lookup a field by name
  ' ------------------------------------------------------------
  Public Function GetField(name As String) As SchemaField
    If String.IsNullOrWhiteSpace(name) Then Return Nothing

    Return Fields.
        FirstOrDefault(Function(f) f.Name.Equals(name, StringComparison.OrdinalIgnoreCase))
  End Function

  ' ------------------------------------------------------------
  '  Add a foreign key
  ' ------------------------------------------------------------
  Public Sub AddForeignKey(fk As SchemaForeignKey)
    If fk Is Nothing Then Exit Sub
    ForeignKeys.Add(fk)
  End Sub

  ' ------------------------------------------------------------
  '  Get all primary key fields
  ' ------------------------------------------------------------
  Public Function PrimaryKeyFields() As List(Of SchemaField)
    Return Fields.
        Where(Function(f) f.IsPrimaryKey).
        ToList()
  End Function

End Class
