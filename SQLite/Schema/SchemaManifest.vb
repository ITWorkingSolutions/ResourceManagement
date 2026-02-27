Option Explicit On

Public Class SchemaManifestRoot
  Public Property Versions As List(Of SchemaManifest)
End Class

Public Class SchemaManifest

  ' ------------------------------------------------------------
  '  Properties
  ' ------------------------------------------------------------
  Public Property Version As String
  ' Strongly typed list of SchemaTable
  Public Property Tables As List(Of SchemaTable)
  ' Strongly typed list of SchemaView
  Public Property Views As List(Of SchemaView)

  ' ------------------------------------------------------------
  '  Constructor
  ' ------------------------------------------------------------
  Public Sub New()
    Tables = New List(Of SchemaTable)()
    Views = New List(Of SchemaView)()
  End Sub

  ' ------------------------------------------------------------
  '  Add a table to the manifest
  ' ------------------------------------------------------------
  Public Sub AddTable(tbl As SchemaTable)
    If tbl Is Nothing Then Exit Sub
    Tables.Add(tbl)
  End Sub

  ' ------------------------------------------------------------
  '  Lookup by table name
  ' ------------------------------------------------------------
  Public Function GetTable(name As String) As SchemaTable
    If String.IsNullOrWhiteSpace(name) Then Return Nothing

    Return Tables.
        FirstOrDefault(Function(t) t.Name.Equals(name, StringComparison.OrdinalIgnoreCase))
  End Function

End Class
