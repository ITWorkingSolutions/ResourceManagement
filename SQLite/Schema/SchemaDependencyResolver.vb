Option Explicit On
Imports System.Collections.Generic
Imports System.Linq

Module SchemaDependencyResolver

  ' ============================================================
  '  ResolveTableCreationOrder
  '
  '  Performs a topological sort over the schema's foreign key
  '  graph to determine the correct table creation order.
  '
  '  SQLite requires parent tables to exist before child tables.
  ' ============================================================
  Friend Function ResolveTableCreationOrder(manifest As SchemaManifest) As List(Of SchemaTable)

    Dim result As New List(Of SchemaTable)()

    ' --------------------------------------------------------
    ' Build dependency counts:
    ' dependencyCount(tableName) = number of parent tables
    ' --------------------------------------------------------
    Dim dependencyCount As New Dictionary(Of String, Integer)(StringComparer.OrdinalIgnoreCase)

    For Each tbl In manifest.Tables
      dependencyCount(tbl.Name) = CountParentDependencies(tbl)
    Next

    ' --------------------------------------------------------
    ' Queue of tables with zero dependencies
    ' --------------------------------------------------------
    Dim ready As New Queue(Of SchemaTable)(
        manifest.Tables.Where(Function(t) dependencyCount(t.Name) = 0)
    )

    ' --------------------------------------------------------
    ' Process tables in dependency order
    ' --------------------------------------------------------
    While ready.Count > 0

      Dim current As SchemaTable = ready.Dequeue()
      result.Add(current)

      ' Reduce dependency counts of tables that depend on this one
      For Each child In manifest.Tables
        If TableDependsOn(child, current.Name) Then

          dependencyCount(child.Name) -= 1

          ' If child now has zero dependencies, enqueue it
          If dependencyCount(child.Name) = 0 Then
            ready.Enqueue(child)
          End If
        End If
      Next
    End While

    ' --------------------------------------------------------
    ' Detect cycles (unresolved dependencies)
    ' --------------------------------------------------------
    Dim unresolved As Integer = 0

    For Each v In dependencyCount.Values
      If v > 0 Then unresolved += 1
    Next

    If unresolved > 0 Then
      Throw New InvalidOperationException(
          "Circular dependency detected in schema. Table creation order cannot be resolved."
      )
    End If

    Return result
  End Function

  ' ------------------------------------------------------------
  '  CountParentDependencies
  '
  '  Returns the number of foreign keys in the table.
  ' ------------------------------------------------------------
  Private Function CountParentDependencies(tbl As SchemaTable) As Integer
    Return tbl.ForeignKeys.Count
  End Function

  ' ------------------------------------------------------------
  '  TableDependsOn
  '
  '  Returns True if tbl has a foreign key referencing parentName.
  ' ------------------------------------------------------------
  Private Function TableDependsOn(tbl As SchemaTable, parentName As String) As Boolean
    Return tbl.ForeignKeys.Any(Function(fk) _
        fk.ReferencesTable.Equals(parentName, StringComparison.OrdinalIgnoreCase))
  End Function

End Module