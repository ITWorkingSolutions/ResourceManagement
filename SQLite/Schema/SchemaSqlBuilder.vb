Option Explicit On
Imports System.Text

Module SchemaSqlBuilder

  ' ============================================================
  '  Purpose:
  '    Convert the in-memory schema representation (SchemaManifest,
  '    SchemaTable, SchemaField, SchemaForeignKey) into SQLite
  '    CREATE TABLE statements.
  '
  '    This module contains NO persistence logic and NO database
  '    operations. It is purely responsible for generating SQL text.
  '
  '    All SQL generation is deterministic and based solely on the
  '    manifest. No assumptions or inference are allowed.
  ' ============================================================


  ' ------------------------------------------------------------
  '  GenerateCreateTableSQL
  ' ------------------------------------------------------------
  Friend Function GenerateCreateTableSQL(tbl As SchemaTable) As String
    Dim sb As New StringBuilder()

    sb.AppendLine($"CREATE TABLE {tbl.Name} (")

    Dim lines As New List(Of String)()

    ' --------------------------------------------------------
    ' Emit field definitions
    ' --------------------------------------------------------
    For Each fld In tbl.Fields
      Dim sqliteType As String = MapFieldTypeToSQLite(fld.Type)

      Dim fieldLine As New StringBuilder()
      fieldLine.Append($"    {fld.Name} {sqliteType}")
      If fld.IsNullable Then
        fieldLine.Append(" NULL")
      Else
        fieldLine.Append(" NOT NULL")
      End If

      If Not String.IsNullOrWhiteSpace(fld.DefaultValue) Then
        fieldLine.Append($" DEFAULT {fld.DefaultValue}")
      End If

      lines.Add(fieldLine.ToString())
    Next

    ' --------------------------------------------------------
    ' Emit PRIMARY KEY clause (supports composite PKs)
    ' --------------------------------------------------------
    Dim pkFields = tbl.PrimaryKeyFields()

    If pkFields.Count > 0 Then
      Dim pkNames = pkFields.Select(Function(f) f.Name)
      Dim pkLine = $"    PRIMARY KEY ({String.Join(", ", pkNames)})"
      lines.Add(pkLine)
    End If

    ' --------------------------------------------------------
    ' Emit FOREIGN KEY clauses
    ' --------------------------------------------------------
    For Each fk In tbl.ForeignKeys
      Dim fkLine As New StringBuilder()

      ' Join FK fields as [Field1], [Field2], ...
      Dim fkFields As String =
    String.Join(", ", fk.Field.Select(Function(f) $"[{f}]"))

      ' Join referenced fields the same way
      Dim refFields As String =
    String.Join(", ", fk.ReferencesField.Select(Function(f) $"[{f}]"))

      fkLine.Append($"    FOREIGN KEY ({fkFields}) REFERENCES {fk.ReferencesTable} ({refFields})")

      If Not String.IsNullOrWhiteSpace(fk.OnDelete) Then
        fkLine.Append($" ON DELETE {fk.OnDelete}")
      End If

      lines.Add(fkLine.ToString())
    Next
    'For Each fk In tbl.ForeignKeys
    '  Dim fkLine As New StringBuilder()

    '  fkLine.Append($"    FOREIGN KEY({fk.Field}) REFERENCES {fk.ReferencesTable}({fk.ReferencesField})")

    '  If Not String.IsNullOrWhiteSpace(fk.OnDelete) Then
    '    fkLine.Append($" ON DELETE {fk.OnDelete}")
    '  End If

    '  lines.Add(fkLine.ToString())
    'Next

    ' --------------------------------------------------------
    ' Combine all lines into final SQL
    ' --------------------------------------------------------
    For i = 0 To lines.Count - 1
      sb.Append(lines(i))

      If i < lines.Count - 1 Then
        sb.AppendLine(",")
      Else
        sb.AppendLine()
      End If
    Next

    sb.Append(");")

    Return sb.ToString()
  End Function


  ' ------------------------------------------------------------
  '  MapFieldTypeToSQLite
  ' ------------------------------------------------------------
  Private Function MapFieldTypeToSQLite(logicalType As String) As String
    Select Case logicalType.Trim().ToUpperInvariant()
      Case "TEXT"
        Return "TEXT"
      Case "LONG"
        Return "INTEGER"
      Case "BOOLEAN"
        Return "INTEGER"
      Case "BLOB"
        Return "BLOB"
      Case Else
        Throw New ArgumentException(
            $"Unknown logical field type: {logicalType}",
            NameOf(logicalType)
        )
    End Select
  End Function

End Module
