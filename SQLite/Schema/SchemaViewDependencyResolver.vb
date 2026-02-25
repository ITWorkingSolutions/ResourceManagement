Option Explicit On
Imports System.Collections.Generic
Imports System.Linq

Module SchemaViewDependencyResolver
  ' ============================================================
  ' Routine: ResolveViewCreationOrder
  '
  ' Purpose: Determines the correct order In which To create views,
  '          based On which tables Or other views they depend On.
  '
  ' Note: SQLite requires that all dependencies exist before
  ' ============================================================
  Friend Function ResolveViewCreationOrder(manifest As SchemaManifest) As List(Of SchemaView)

    Dim result As New List(Of SchemaView)()

    ' Identify all view names
    Dim viewNames As New HashSet(Of String)(
            manifest.Views.Select(Function(v) v.Name),
            StringComparer.OrdinalIgnoreCase
        )

    ' Build dependency count: only count dependencies on other views
    Dim dependencyCount As New Dictionary(Of String, Integer)(StringComparer.OrdinalIgnoreCase)

    For Each vw In manifest.Views
      'dependencyCount(vw.Name) =
      '          vw.DependsOn.Count(Function(d) viewNames.Contains(d))
      dependencyCount(vw.Name) =
        System.Linq.Enumerable.Count(
            vw.DependsOn,
            Function(d) viewNames.Contains(d)
        )
    Next

    ' Queue of views with zero view-dependencies
    Dim ready As New Queue(Of SchemaView)(
            manifest.Views.Where(Function(v) dependencyCount(v.Name) = 0)
        )

    ' Process in dependency order
    While ready.Count > 0

      Dim current As SchemaView = ready.Dequeue()
      result.Add(current)

      ' Reduce dependency counts of views that depend on this one
      For Each child In manifest.Views
        If child.DependsOn.Any(Function(d) d.Equals(current.Name, StringComparison.OrdinalIgnoreCase)) Then

          dependencyCount(child.Name) -= 1

          If dependencyCount(child.Name) = 0 Then
            ready.Enqueue(child)
          End If
        End If
      Next
    End While

    ' Detect cycles
    If dependencyCount.Values.Any(Function(v) v > 0) Then
      Throw New InvalidOperationException(
                "Circular dependency detected in view definitions."
            )
    End If

    Return result
  End Function

End Module
