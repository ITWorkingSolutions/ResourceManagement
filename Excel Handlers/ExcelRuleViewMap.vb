Imports System.Runtime.CompilerServices
Imports System.Text

Friend Class ExcelRuleViewMap
  Public Property Views As List(Of ExcelRuleViewMapView)

End Class

Friend Class ExcelRuleViewMapView
  Public Property Name As String
  Public Property DisplayName As String
  Public Property Fields As List(Of ExcelRuleViewMapField)
  Public Property Relations As List(Of ExcelRuleViewMapRelation)
  Public Property MergedFieldTarget As Boolean   ' Allows conceptual fields to anchor here
  Public Property IsConceptual As Boolean        ' Optional: marks conceptual views
End Class

Friend Class ExcelRuleViewMapField
  Public Property Name As String
  Public Property DisplayName As String
  Public Property DataType As String      ' "Text", "Number", "Boolean"
  Public Property Role As String          ' "Key", "Attribute", "Lookup", etc.
  Public Property FieldID As String   ' Null for physical fields, set for custom fields
  Public Property SourceField As String    ' Holds the field name from which the Field physically comes from
  Public Property SourceView As String    ' Holds the view name from which the Field physically comes from
  ' --- Time-bounded semantics ---
  Public Property IsTimeBounded As Boolean
  Public Property PartnerField As String
  Public Property RequiresPartner As Boolean

  ' --- Operator constraints ---
  Public Property AllowedOperators As List(Of String)

  ' --- Allowed fixed values (e.g. Availability) ---
  Public Property AllowedValues As List(Of String)

  ' --- Slicing support (only for some time-bounded fields) ---
  Public Property SupportsSlicing As Boolean
  Public Property SlicingDescriptions As Dictionary(Of String, String)
  Public Property SlicingOptions As List(Of String)
  Public Property DefaultSlicing As String

  ' --- Handler indirection for domain-specific logic ---
  Public Property TimeBoundHandler As String

  ' --- UI placement for conceptual fields ---
  Public Property MergedIntoView As String
End Class

Friend Class ExcelRuleViewMapRelation
  Public Property ToView As String
  Public Property Join As List(Of ExcelRuleViewMapJoinField)
End Class

Friend Class ExcelRuleViewMapJoinField
  Public Property FromField As String
  Public Property ToField As String
End Class

'Friend Class ValueQueryInfo
'  Public Property Sql As String
'  Public Property KeyColumn As String
'  Public Property DisplayColumn As String
'  Public Property IsListView As Boolean

'  Public Sub New(sql As String, keyCol As String, displayCol As String, isList As Boolean)
'    Me.Sql = sql
'    Me.KeyColumn = keyCol
'    Me.DisplayColumn = displayCol
'    Me.IsListView = isList
'  End Sub
'End Class

Friend Class SelectFields
  Friend Property sourceView As String
  Friend Property sourceField As String
  Friend Property fieldID As String
  Friend Property listTypeID As String ' used by vwDim_List
End Class
Friend Class ExcelRuleViewMapHelper

  Private ReadOnly _map As ExcelRuleViewMap

  Friend Sub New(map As ExcelRuleViewMap)
    If map Is Nothing Then Throw New ArgumentNullException(NameOf(map))
    _map = map
  End Sub

  ' ============================================================
  '  View lookup
  ' ============================================================
  Friend Function GetView(viewName As String) As ExcelRuleViewMapView
    Return _map.Views.
            FirstOrDefault(Function(v) v.Name.Equals(viewName, StringComparison.OrdinalIgnoreCase))
  End Function

  Friend Function GetViewDisplayName(viewName As String) As String
    Dim v = GetView(viewName)
    Return If(v Is Nothing, Nothing, v.DisplayName)
  End Function

  ' ============================================================
  '  Field lookup
  ' ============================================================
  Friend Function GetField(viewName As String, fieldName As String) As ExcelRuleViewMapField
    Dim v = GetView(viewName)
    If v Is Nothing Then Return Nothing

    Return v.Fields.
            FirstOrDefault(Function(f) f.Name.Equals(fieldName, StringComparison.OrdinalIgnoreCase))
  End Function

  Friend Function GetFieldDisplayName(viewName As String, fieldName As String) As String
    Dim f = GetField(viewName, fieldName)
    Return If(f Is Nothing, Nothing, f.DisplayName)
  End Function

  Friend Function GetFields(viewName As String) As List(Of ExcelRuleViewMapField)
    Dim v = GetView(viewName)
    Return If(v Is Nothing, New List(Of ExcelRuleViewMapField), v.Fields)
  End Function

  Friend Function GetFieldsByRole(viewName As String, role As String) As List(Of ExcelRuleViewMapField)
    Dim v = GetView(viewName)
    If v Is Nothing Then Return New List(Of ExcelRuleViewMapField)

    Return v.Fields.
            Where(Function(f) f.Role.Equals(role, StringComparison.OrdinalIgnoreCase)).
            ToList()
  End Function

  ' ============================================================
  '  Relationship lookup
  ' ============================================================

  ' --- One-way relations FROM a view
  Friend Function GetRelationsFrom(viewName As String) As List(Of ExcelRuleViewMapRelation)
    Dim v As ExcelRuleViewMapView = GetView(viewName)
    If v Is Nothing Then Return New List(Of ExcelRuleViewMapRelation)
    Return v.Relations
  End Function

  ' --- Reverse relations TO a view
  Friend Function GetRelationsTo(viewName As String) As List(Of ExcelRuleViewMapRelation)
    Dim list As New List(Of ExcelRuleViewMapRelation)

    For Each v In _map.Views
      For Each r In v.Relations
        If r.ToView.Equals(viewName, StringComparison.OrdinalIgnoreCase) Then
          list.Add(New ExcelRuleViewMapRelation With {
                        .ToView = v.Name,
                        .Join = r.Join.Select(
                            Function(j) New ExcelRuleViewMapJoinField With {
                                .FromField = j.ToField,   ' reverse direction
                                .ToField = j.FromField
                            }).ToList()
                    })
        End If
      Next
    Next

    Return list
  End Function

  ' ============================================================
  ' Multi-hop join path search (BFS)
  ' ============================================================
  Friend Function FindJoinPath(fromView As String, toView As String) As List(Of ExcelRuleViewMapRelation)

    Dim visited As New HashSet(Of String)(StringComparer.OrdinalIgnoreCase)
    Dim queue As New Queue(Of (View As String, Path As List(Of ExcelRuleViewMapRelation)))

    queue.Enqueue((fromView, New List(Of ExcelRuleViewMapRelation)))

    While queue.Count > 0
      Dim current = queue.Dequeue()

      If visited.Contains(current.View) Then Continue While
      visited.Add(current.View)

      ' Found target
      If current.View.Equals(toView, StringComparison.OrdinalIgnoreCase) Then
        Return current.Path
      End If

      ' Explore forward relations
      For Each rel In GetRelationsFrom(current.View)
        Dim nextPath = New List(Of ExcelRuleViewMapRelation)(current.Path)
        nextPath.Add(rel)
        queue.Enqueue((rel.ToView, nextPath))
      Next

      ' Explore reverse relations
      For Each rel In GetRelationsTo(current.View)
        Dim nextPath = New List(Of ExcelRuleViewMapRelation)(current.Path)
        nextPath.Add(rel)
        queue.Enqueue((rel.ToView, nextPath))
      Next
    End While

    Return Nothing
  End Function

  ' ============================================================
  ' Get join fields for a direct relationship
  ' ============================================================
  Friend Function GetJoinFields(fromView As String, toView As String) As List(Of ExcelRuleViewMapJoinField)

    ' Try forward
    Dim v = GetView(fromView)
    If v IsNot Nothing Then
      Dim rel = v.Relations.FirstOrDefault(
                Function(r) r.ToView.Equals(toView, StringComparison.OrdinalIgnoreCase))
      If rel IsNot Nothing Then Return rel.Join
    End If

    ' Try reverse
    For Each vv In _map.Views
      For Each r In vv.Relations
        If r.ToView.Equals(fromView, StringComparison.OrdinalIgnoreCase) Then
          If vv.Name.Equals(toView, StringComparison.OrdinalIgnoreCase) Then
            Return r.Join.Select(
                            Function(j) New ExcelRuleViewMapJoinField With {
                                .FromField = j.ToField,
                                .ToField = j.FromField
                            }).ToList()
          End If
        End If
      Next
    Next

    Return Nothing
  End Function

  'Friend Function CanJoin(fromView As String, toView As String) As Boolean
  '  Dim v = GetView(fromView)
  '  If v Is Nothing Then Return False

  '  Return v.Relations.Any(Function(r) r.ToView.Equals(toView, StringComparison.OrdinalIgnoreCase))
  'End Function

  'Friend Function GetJoinField(fromView As String, toView As String) As String
  '  Dim v = GetView(fromView)
  '  If v Is Nothing Then Return Nothing

  '  Dim rel = v.Relations.
  '          FirstOrDefault(Function(r) r.ToView.Equals(toView, StringComparison.OrdinalIgnoreCase))

  '  Return If(rel Is Nothing, Nothing, rel.Via)
  'End Function


  ' ============================================================
  '  Operator rules based on DataType
  ' ============================================================
  Friend Function GetOperatorsForDataType(dataType As String) As List(Of String)
    Select Case dataType.Trim().ToLowerInvariant()
      Case "text"
        Return New List(Of String) From {"=", "<>", "Contains", "StartsWith", "EndsWith"}

      Case "number"
        Return New List(Of String) From {"=", "<>", "<", "<=", ">", ">="}

      Case "boolean"
        Return New List(Of String) From {"=", "<>"}

      Case Else
        Return New List(Of String)
    End Select
  End Function


  ' ============================================================
  '  Convenience helpers
  ' ============================================================
  Friend Function GetAllViewNames() As List(Of String)
    Return _map.Views.Select(Function(v) v.Name).ToList()
  End Function

  Friend Function GetAllViewDisplayNames() As List(Of String)
    Return _map.Views.Select(Function(v) v.DisplayName).ToList()
  End Function

  ' ============================================================
  '  Class used in BuildRuleQuery
  ' ============================================================
  Friend Class TimeBoundFilterContext
    Public Property Filter As UIExcelRuleDesignerRuleFilter
    Public Property Value As Object
    Public Property FieldMap As ExcelRuleViewMapField
  End Class

  ' ==========================================================================================
  ' Routine:    BuildRuleQuery
  ' Purpose:
  '   Build the Base SQL query AND the Base QuerySpec, then invoke any handler routines
  '   (e.g., BuildAvailabilityQuery) to add additional queries. Returns:
  '
  '       sqlQueries("Base")         → Base SQL
  '       sqlQueries("Availability") → Availability SQL (if needed)
  '
  '       querySpecs("Base")         → Base QuerySpec
  '       querySpecs("Availability") → Availability QuerySpec (if needed)
  '
  '   The Base query is NEVER mutated by handlers. Internal surrogate keys (__Key_*) are
  '   ALWAYS added to the Base SELECT list so handlers can stitch results.
  '
  ' Returns:
  '   (sqlQueries, querySpecs)
  ' ==========================================================================================
  Friend Sub BuildRuleQuery(
      anchorView As String,
      selectFields As List(Of SelectFields),
      filters As List(Of UIExcelRuleDesignerRuleFilter),
      args() As Object,
      ByRef sqlQueries As Dictionary(Of String, String),
      ByRef querySpecs As Dictionary(Of String, QuerySpec)
  )

    sqlQueries = New Dictionary(Of String, String)(StringComparer.OrdinalIgnoreCase)
    querySpecs = New Dictionary(Of String, QuerySpec)(StringComparer.OrdinalIgnoreCase)

    Dim sqlFrom As New System.Text.StringBuilder()
    Dim selectParts As New List(Of String)()
    Dim whereParts As New List(Of String)()
    Dim aliasMap As New Dictionary(Of String, String)(StringComparer.OrdinalIgnoreCase)
    Dim aliasCounter As Integer = 1
    Dim argIndex As Integer = 0
    Dim vm As New ExcelRuleViewMapHelper(ExcelRuleViewMapLoader.LoadExcelRuleViewMap())

    ' Group time-bounded filters by handler name
    Dim timeBoundGroups As New Dictionary(Of String, List(Of TimeBoundFilterContext))(StringComparer.OrdinalIgnoreCase)

    Try
      ' ==================================================================================
      ' 0. Validate
      ' ==================================================================================
      If String.IsNullOrEmpty(anchorView) Then
        Throw New ArgumentException("anchorView cannot be null or empty.")
      End If

      If selectFields Is Nothing OrElse selectFields.Count = 0 Then
        Throw New ArgumentException("selectFields cannot be empty.")
      End If

      ' ==================================================================================
      ' 1. FROM anchorView
      ' ==================================================================================
      sqlFrom.AppendLine("FROM [" & anchorView & "] AS V0")
      aliasMap(anchorView) = "V0"

      ' ==================================================================================
      ' 2. Build JOIN graph for selectFields + filters
      ' ==================================================================================
      Dim requiredViews As New HashSet(Of String)(StringComparer.OrdinalIgnoreCase)

      For Each sf In selectFields
        Dim viewMap = vm.GetView(sf.sourceView)
        If viewMap IsNot Nothing AndAlso viewMap.IsConceptual Then Continue For
        requiredViews.Add(sf.sourceView)
      Next

      For Each f In filters
        Dim viewMap = vm.GetView(f.SourceView)
        If viewMap IsNot Nothing AndAlso viewMap.IsConceptual Then Continue For
        requiredViews.Add(f.SourceView)
      Next

      For Each viewName In requiredViews

        If aliasMap.ContainsKey(viewName) Then Continue For

        Dim path = FindJoinPath(anchorView, viewName)
        If path Is Nothing Then
          Throw New InvalidOperationException(
              "No join path found between " & anchorView & " and " & viewName & ".")
        End If

        Dim currentView = anchorView
        Dim currentAlias = aliasMap(currentView)

        For Each rel In path

          Dim nextView = rel.ToView

          If aliasMap.ContainsKey(nextView) Then
            currentView = nextView
            currentAlias = aliasMap(nextView)
            Continue For
          End If

          Dim nextAlias = "V" & aliasCounter
          aliasCounter += 1

          sqlFrom.AppendLine("JOIN [" & nextView & "] AS " & nextAlias & " ON")

          Dim joinParts As New List(Of String)
          For Each jf In rel.Join
            joinParts.Add(currentAlias & ".[" & jf.FromField & "] = " &
                          nextAlias & ".[" & jf.ToField & "]")
          Next

          sqlFrom.AppendLine("  " & String.Join(" AND ", joinParts))

          aliasMap(nextView) = nextAlias
          currentView = nextView
          currentAlias = nextAlias
        Next
      Next

      ' ==================================================================================
      ' 3. Build SELECT list (user fields)
      ' ==================================================================================
      For Each sf In selectFields

        Dim viewMap = vm.GetView(sf.sourceView)
        If viewMap IsNot Nothing AndAlso viewMap.IsConceptual Then Continue For

        Dim aliasName = aliasMap(sf.sourceView)
        Dim col = sf.sourceField

        If Not String.IsNullOrEmpty(sf.fieldID) Then
          Dim expr As String =
              "CASE WHEN " & aliasName & ".ValueName IS NOT NULL " &
              "THEN " & aliasName & ".ValueName " &
              "ELSE " & aliasName & ".ValueText END AS [" & col & "]"

          selectParts.Add(expr)
        Else
          selectParts.Add(aliasName & ".[" & col & "] AS [" & col & "]")
        End If
      Next

      ' ==================================================================================
      ' 3b. Add internal surrogate keys (__Key_*)
      ' ==================================================================================
      For Each viewName In aliasMap.Keys

        Dim viewMap = vm.GetView(viewName)
        If viewMap Is Nothing OrElse viewMap.Fields Is Nothing Then Continue For

        Dim aliasName = aliasMap(viewName)

        For Each fld In viewMap.Fields
          If String.Equals(fld.Role, "Key", StringComparison.OrdinalIgnoreCase) Then

            Dim internalAlias As String =
                "__Key_" & viewName & "_" & fld.Name

            selectParts.Add(aliasName & ".[" & fld.Name & "] AS [" & internalAlias & "]")
          End If
        Next
      Next

      ' ==================================================================================
      ' 4. Build WHERE clause (non-time-bounded filters)
      ' ==================================================================================
      For Each f In filters

        If String.IsNullOrEmpty(f.FieldOperator) Then Continue For

        Dim value As Object = Nothing

        If String.Equals(f.ValueBinding, ValueBinding.Rule.ToString(), StringComparison.OrdinalIgnoreCase) Then
          value = f.ValueBinding
        Else
          If args IsNot Nothing AndAlso argIndex < args.Length Then
            value = args(argIndex)
          End If
          argIndex += 1
        End If

        Dim viewMap = vm.GetView(f.SourceView)
        Dim fieldMap = vm.GetField(f.SourceView, f.SourceField)

        ' TIME-BOUNDED → handler
        If fieldMap IsNot Nothing AndAlso fieldMap.IsTimeBounded Then

          Dim handlerName = fieldMap.TimeBoundHandler

          If Not timeBoundGroups.ContainsKey(handlerName) Then
            timeBoundGroups(handlerName) = New List(Of TimeBoundFilterContext)
          End If

          timeBoundGroups(handlerName).Add(New TimeBoundFilterContext With {
              .Filter = f,
              .Value = value,
              .FieldMap = fieldMap
          })

          Continue For
        End If

        ' Conceptual → skip
        If viewMap IsNot Nothing AndAlso viewMap.IsConceptual Then Continue For

        ' Normal predicate
        Dim aliasName = aliasMap(f.SourceView)
        Dim predicate As String

        Dim s As String = If(value Is Nothing, Nothing, value.ToString().Trim())

        If Not String.IsNullOrEmpty(s) AndAlso
           (s.Equals("true", StringComparison.OrdinalIgnoreCase) OrElse
            s.Equals("false", StringComparison.OrdinalIgnoreCase)) Then

          Dim normalized = s.ToLowerInvariant()

          predicate =
              "LOWER(" & aliasName & ".[" & f.SourceField & "]) " &
              f.FieldOperator & " '" & normalized.Replace("'", "''") & "'"

        Else
          Dim literal As String = FormatSqlLiteral(value)

          predicate =
              aliasName & ".[" & f.SourceField & "] " &
              f.FieldOperator & " " & literal
        End If

        If whereParts.Count > 0 AndAlso Not String.IsNullOrEmpty(f.BooleanOperator) Then
          Dim boolOp As String = f.BooleanOperator.Trim().ToUpperInvariant()
          If boolOp <> "AND" AndAlso boolOp <> "OR" Then boolOp = "AND"
          predicate = boolOp & " " & predicate
        End If

        whereParts.Add(predicate)
      Next

      ' ==================================================================================
      ' 5. SPECIAL CASE: vwDim_List ListTypeID enforcement
      ' ==================================================================================
      For Each sf In selectFields
        If sf.sourceView.Equals("vwDim_List", StringComparison.OrdinalIgnoreCase) Then

          If Not String.IsNullOrEmpty(sf.listTypeID) Then

            Dim aliasName = aliasMap("vwDim_List")
            Dim safeListTypeId = sf.listTypeID.Replace("'", "''")

            Dim predicate As String =
                aliasName & ".[ListTypeID] = '" & safeListTypeId & "'"

            whereParts.Insert(0, predicate)
          End If
        End If
      Next

      ' ==================================================================================
      ' SPECIAL CASE: vwFact_ResourceCustomField requires FieldID filtering
      ' ==================================================================================
      For Each sf In selectFields
        If sf.sourceView.Equals("vwFact_ResourceCustomField", StringComparison.OrdinalIgnoreCase) Then

          If Not String.IsNullOrEmpty(sf.fieldID) Then
            Dim aliasName = aliasMap("vwFact_ResourceCustomField")
            Dim safeFieldId = sf.fieldID.Replace("'", "''")

            Dim predicate As String =
          aliasName & ".[FieldID] = '" & safeFieldId & "'"

            ' Make it a hard invariant, same as you did for vwDim_List
            whereParts.Insert(0, predicate)
          End If

        End If
      Next

      ' ==================================================================================
      ' 6. Build Base QuerySpec
      ' ==================================================================================
      Dim baseSpec As New QuerySpec With {
          .AliasMap = aliasMap,
          .JoinGraph = New System.Text.StringBuilder(sqlFrom.ToString()),
          .WhereParts = New List(Of String)(whereParts),
          .SelectParts = New List(Of String)(selectParts),
          .AnchorView = anchorView,
          .AliasCounter = aliasCounter
      }

      querySpecs("Base") = baseSpec

      ' ==================================================================================
      ' 7. Build Base SQL
      ' ==================================================================================
      Dim sbBase As New System.Text.StringBuilder()

      sbBase.AppendLine("SELECT DISTINCT " & String.Join(", ", selectParts))
      sbBase.Append(sqlFrom.ToString())

      If whereParts.Count > 0 Then
        sbBase.AppendLine("WHERE " & String.Join(" ", whereParts))
      End If

      sqlQueries("Base") = sbBase.ToString()

      ' ==================================================================================
      ' 8. Invoke handlers (Availability, etc.)
      ' ==================================================================================
      For Each kvp In timeBoundGroups
        Dim handlerName = kvp.Key
        Dim contexts = kvp.Value

        Select Case handlerName
          Case "Availability"
            BuildAvailabilityQuery(contexts, vm, querySpecs, sqlQueries)

          Case Else
            Throw New InvalidOperationException("Unknown timeBoundHandler: " & handlerName)
        End Select
      Next

    Catch ex As Exception
      ErrorHandler.UnHandleError(ex)
      sqlQueries("Base") = String.Empty

    End Try

  End Sub

End Class