Option Explicit On

Friend Class SchemaManifestRoot
  Friend Property Versions As List(Of SchemaManifest)
  Friend Property Views As List(Of SchemaView)
End Class

Friend Class SchemaManifest

  ' ------------------------------------------------------------
  '  Properties
  ' ------------------------------------------------------------
  Friend Property Version As String
  ' Strongly typed list of SchemaTable
  Friend Property Tables As List(Of SchemaTable)
  ' Strongly typed list of SchemaView
  Friend Property Views As List(Of SchemaView)

  ' ------------------------------------------------------------
  '  Constructor
  ' ------------------------------------------------------------
  Friend Sub New()
    Tables = New List(Of SchemaTable)()
    Views = New List(Of SchemaView)()
  End Sub

  ' ------------------------------------------------------------
  '  Add a table to the manifest
  ' ------------------------------------------------------------
  Friend Sub AddTable(tbl As SchemaTable)
    If tbl Is Nothing Then Exit Sub
    Tables.Add(tbl)
  End Sub

  ' ------------------------------------------------------------
  '  Lookup by table name
  ' ------------------------------------------------------------
  Friend Function GetTable(name As String) As SchemaTable
    If String.IsNullOrWhiteSpace(name) Then Return Nothing

    Return Tables.
        FirstOrDefault(Function(t) t.Name.Equals(name, StringComparison.OrdinalIgnoreCase))
  End Function

End Class
