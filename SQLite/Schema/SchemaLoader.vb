Option Explicit On

Imports System.IO
Imports System.Reflection
Imports System.Text.Json

Module SchemaLoader

  ' ============================================================
  '  LoadSchemaManifest
  '
  '  Loads schema_manifest.json embedded in the DLL and
  '  deserializes it into a strongly typed SchemaManifest.
  '
  '  This is the ONLY source of truth for schema metadata.
  ' ============================================================
  Friend Function LoadSchemaManifest() As SchemaManifest
    Dim jsonText As String = LoadEmbeddedSchemaJson()
    Return DeserializeSchema(jsonText)
  End Function


  ' ------------------------------------------------------------
  '  LoadEmbeddedSchemaJson
  '
  '  Reads schema_manifest.json from embedded resources.
  ' ------------------------------------------------------------
  Private Function LoadEmbeddedSchemaJson() As String
    Dim asm = Assembly.GetExecutingAssembly()

    ' Find the embedded resource name
    Dim resourceName As String =
        asm.GetManifestResourceNames().
            First(Function(n) n.EndsWith("SchemaManifest.json",
                    StringComparison.OrdinalIgnoreCase))

    Using stream = asm.GetManifestResourceStream(resourceName)
      Using reader As New StreamReader(stream)
        Return reader.ReadToEnd()
      End Using
    End Using
  End Function

  ' ------------------------------------------------------------
  '  LoadEmbeddedTextResource
  '
  '  Reads the sql from embedded resources.
  ' ------------------------------------------------------------
  Friend Function LoadEmbeddedTextResource(resourceName As String) As String
    Dim asm = Assembly.GetExecutingAssembly()
    Using stream = asm.GetManifestResourceStream(resourceName)
      If stream Is Nothing Then
        Throw New FileNotFoundException($"Embedded resource not found: {resourceName}")
      End If
      Using reader As New StreamReader(stream)
        Return reader.ReadToEnd()
      End Using
    End Using
  End Function

  ' ------------------------------------------------------------
  '  DeserializeSchema
  '
  '  Converts JSON text into a SchemaManifestRoot, then selects
  '  the latest schema version (Major.Minor.Patch).
  ' ------------------------------------------------------------
  Private Function DeserializeSchema(jsonText As String) As SchemaManifest

    ' === Variable declarations ===
    Dim options As JsonSerializerOptions
    Dim root As SchemaManifestRoot
    Dim latest As SchemaManifest

    Try
      options = New JsonSerializerOptions With {
      .PropertyNameCaseInsensitive = True
    }

      root = JsonSerializer.Deserialize(Of SchemaManifestRoot)(jsonText, options)

      If root Is Nothing OrElse root.Versions Is Nothing OrElse root.Versions.Count = 0 Then
        Throw New InvalidDataException("Schema manifest contains no versions.")
      End If

      ' ------------------------------------------------------------
      '  Select the latest version using semantic version comparison
      ' ------------------------------------------------------------
      latest = root.Versions _
      .OrderBy(Function(v) Version.Parse(v.Version)) _
      .Last()

      ' ------------------------------------------------------------
      '  Ensure lists are initialized
      ' ------------------------------------------------------------
      If latest.Tables Is Nothing Then
        latest.Tables = New List(Of SchemaTable)()
      End If

      For Each tbl In latest.Tables
        If tbl.Fields Is Nothing Then tbl.Fields = New List(Of SchemaField)()
        If tbl.ForeignKeys Is Nothing Then tbl.ForeignKeys = New List(Of SchemaForeignKey)()
      Next

      ' ------------------------------------------------------------
      '  Enforce schema invariants - primary key fields are NOT NULL
      ' ------------------------------------------------------------
      For Each tbl In latest.Tables
        For Each fld In tbl.Fields
          If fld.IsPrimaryKey Then
            fld.IsNullable = False
          End If
        Next
      Next

      Return latest

    Catch ex As Exception
      Throw New InvalidDataException("Schema manifest could not be deserialized.", ex)

    Finally
      options = Nothing
      root = Nothing
      latest = Nothing
    End Try

  End Function

End Module
