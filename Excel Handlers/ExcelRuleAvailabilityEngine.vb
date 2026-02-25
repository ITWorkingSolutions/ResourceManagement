Imports ResourceManagement.ExcelRuleViewMapHelper

Module ExcelRuleAvailabilityEngine

  ' ==========================================================================================
  ' Routine: BuildAvailabilityQuery
  ' Purpose:
  '   Build a SEPARATE Availability SQL query using the Base QuerySpec stored in querySpecs("Base").
  '
  '   This routine:
  '     - CLONES the base FROM/JOIN graph (never mutates it)
  '     - SELECTs internal surrogate keys (__Key_*)
  '     - SELECTs raw pattern fields from vwFact_ResourceAvailabilityPattern
  '     - LEFT JOINs the pattern table
  '     - WHERE includes:
  '           * base non-time-bounded predicates
  '           * coarse overlap predicate for AvailabilityFromDate/ToDate
  '
  '   The resulting Availability SQL is stored in sqlQueries("Availability") and its structure
  '   is stored in querySpecs("Availability").
  '
  ' Parameters:
  '   contexts    - AvailabilityFromDate / AvailabilityToDate contexts.
  '   vm          - View map helper.
  '   querySpecs  - BYREF dictionary of QuerySpec objects ("Base", "Availability", etc.)
  '   sqlQueries  - BYREF dictionary of SQL strings ("Base", "Availability", etc.)
  '
  ' Returns:
  '   None (Availability SQL and QuerySpec are stored in the dictionaries).
  '
  ' Notes:
  '   - Internal keys are ALWAYS aliased as "__Key_{View}_{Field}".
  '   - User-visible fields are NEVER included in the Availability query.
  '   - Base query is NEVER mutated.
  ' ==========================================================================================
  Friend Sub BuildAvailabilityQuery(
    contexts As List(Of TimeBoundFilterContext),
    vm As ExcelRuleViewMapHelper,
    ByRef querySpecs As Dictionary(Of String, QuerySpec),
    ByRef sqlQueries As Dictionary(Of String, String)
)

    Try
      ' ----------------------------------------------------------------------
      ' 1. Get Base QuerySpec
      ' ----------------------------------------------------------------------
      If Not querySpecs.ContainsKey("Base") Then Exit Sub

      Dim baseSpec As QuerySpec = querySpecs("Base")

      ' ----------------------------------------------------------------------
      ' 2. Clone BaseSpec for Availability
      ' ----------------------------------------------------------------------
      Dim availSpec As New QuerySpec With {
            .AliasMap = New Dictionary(Of String, String)(baseSpec.AliasMap, StringComparer.OrdinalIgnoreCase),
            .JoinGraph = New System.Text.StringBuilder(baseSpec.JoinGraph.ToString()),
            .WhereParts = New List(Of String)(baseSpec.WhereParts),
            .SelectParts = New List(Of String)(),   ' handler builds its own SELECT
            .AnchorView = baseSpec.AnchorView,
            .AliasCounter = baseSpec.AliasCounter
        }

      ' ----------------------------------------------------------------------
      ' 3. SELECT internal keys (__Key_*)
      ' ----------------------------------------------------------------------
      For Each viewName In availSpec.AliasMap.Keys
        Dim viewMap = vm.GetView(viewName)
        If viewMap Is Nothing OrElse viewMap.Fields Is Nothing Then Continue For

        Dim aliasName = availSpec.AliasMap(viewName)

        For Each fld In viewMap.Fields
          If String.Equals(fld.Role, "Key", StringComparison.OrdinalIgnoreCase) Then
            Dim internalAlias As String =
                        "__Key_" & viewName & "_" & fld.Name

            availSpec.SelectParts.Add(
                        aliasName & ".[" & fld.Name & "] AS [" & internalAlias & "]"
                    )
          End If
        Next
      Next

      ' ----------------------------------------------------------------------
      ' 4. Add pattern JOIN
      ' ----------------------------------------------------------------------
      Dim patternAlias As String = "V" & availSpec.AliasCounter
      availSpec.AliasCounter += 1

      Dim anchorAlias As String = availSpec.AliasMap(availSpec.AnchorView)

      availSpec.JoinGraph.AppendLine(
            "LEFT JOIN vwFact_ResourceAvailabilityPattern AS " & patternAlias &
            " ON " & patternAlias & ".ResourceID = " & anchorAlias & ".ResourceID"
        )

      ' ----------------------------------------------------------------------
      ' 5. SELECT pattern fields
      ' ----------------------------------------------------------------------
      availSpec.SelectParts.Add(patternAlias & ".PatternType")
      availSpec.SelectParts.Add(patternAlias & ".PatternEndType")
      availSpec.SelectParts.Add(patternAlias & ".PatternEndAfterOccurrences")
      availSpec.SelectParts.Add(patternAlias & ".Mode")
      availSpec.SelectParts.Add(patternAlias & ".AllDay")
      availSpec.SelectParts.Add(patternAlias & ".StartTime")
      availSpec.SelectParts.Add(patternAlias & ".EndTime")
      availSpec.SelectParts.Add(patternAlias & ".RangeStartDate")
      availSpec.SelectParts.Add(patternAlias & ".RangeEndDate")
      availSpec.SelectParts.Add(patternAlias & ".PatternStartDate")
      availSpec.SelectParts.Add(patternAlias & ".PatternEndDate")
      availSpec.SelectParts.Add(patternAlias & ".RecurWeeks")
      availSpec.SelectParts.Add(patternAlias & ".WeeklyDayOfWeek")
      availSpec.SelectParts.Add(patternAlias & ".MonthlyType")
      availSpec.SelectParts.Add(patternAlias & ".MonthlyDayOfMonth")
      availSpec.SelectParts.Add(patternAlias & ".MonthlyOrdinal")
      availSpec.SelectParts.Add(patternAlias & ".MonthlyDayOfWeek")
      availSpec.SelectParts.Add(patternAlias & ".RecurMonths")

      ' ----------------------------------------------------------------------
      ' 6. WHERE: coarse overlap
      ' ----------------------------------------------------------------------
      Dim fromDate As Object = Nothing
      Dim toDate As Object = Nothing

      For Each ctx In contexts
        Select Case ctx.FieldMap.Name
          Case "AvailabilityFromDate" : fromDate = ctx.Value
          Case "AvailabilityToDate" : toDate = ctx.Value
        End Select
      Next

      If fromDate IsNot Nothing AndAlso toDate IsNot Nothing Then

        Dim fromLit = FormatSqlLiteral(fromDate)
        Dim toLit = FormatSqlLiteral(toDate)

        Dim dateRangeExpr As String =
                "(" &
                patternAlias & ".PatternType = 'DateRange' AND " &
                patternAlias & ".RangeEndDate >= " & fromLit & " AND " &
                patternAlias & ".RangeStartDate <= " & toLit &
                ")"

        Dim recurringExpr As String =
                "(" &
                patternAlias & ".PatternType IN ('Weekly','Monthly') AND " &
                patternAlias & ".PatternStartDate <= " & toLit & " AND " &
                "(" &
                    patternAlias & ".PatternEndType = 'None' OR " &
                    "(" & patternAlias & ".PatternEndType = 'By' AND " &
                        patternAlias & ".PatternEndDate >= " & fromLit &
                    ")" &
                    " OR " &
                    patternAlias & ".PatternEndType = 'After'" &
                ")" &
                ")"

        Dim nullExpr As String = "(" & patternAlias & ".PatternType IS NULL)"

        Dim coreExpr As String =
                "(" & nullExpr & " OR " & dateRangeExpr & " OR " & recurringExpr & ")"

        Dim finalExpr As String =
                If(availSpec.WhereParts.Count = 0, coreExpr, "AND " & coreExpr)

        availSpec.WhereParts.Add(finalExpr)
      End If

      ' ----------------------------------------------------------------------
      ' 7. Build final SQL
      ' ----------------------------------------------------------------------
      Dim sb As New System.Text.StringBuilder()

      sb.AppendLine("SELECT DISTINCT " & String.Join(", ", availSpec.SelectParts))
      sb.Append(availSpec.JoinGraph.ToString())

      If availSpec.WhereParts.Count > 0 Then
        sb.AppendLine("WHERE " & String.Join(" ", availSpec.WhereParts))
      End If

      ' ----------------------------------------------------------------------
      ' 8. Store results
      ' ----------------------------------------------------------------------
      sqlQueries("Availability") = sb.ToString()
      querySpecs("Availability") = availSpec

    Catch ex As Exception
      ErrorHandler.UnHandleError(ex)
      sqlQueries("Availability") = String.Empty

    Finally
      ' No cleanup required
    End Try

  End Sub

  ' ==========================================================================================
  ' Routine: AvailabilityTimeBoundHandler
  ' Purpose:
  '   Mutate the SQL builders (JOINs, SELECT columns, WHERE expressions) to support
  '   time-bounded Availability filtering inside BuildRuleQuery.
  '
  ' Parameters:
  '   contexts      - List of TimeBoundFilterContext objects representing the conceptual
  '                   AvailabilityFromDate / AvailabilityToDate filters and their values.
  '
  '   aliasMap      - Dictionary mapping view names to their assigned SQL aliases (e.g. "V0").
  '
  '   anchorView    - The name of the anchor view used by BuildRuleQuery (e.g. "vwDim_Resource").
  '
  '   aliasCounter  - BYREF integer used to allocate new SQL aliases (V1, V2, V3...).
  '                   This routine increments it for each JOIN it adds.
  '
  '   sqlFrom       - BYREF StringBuilder containing the FROM/JOIN clause being built.
  '
  '   selectParts   - BYREF List(Of String) containing SELECT column expressions.
  '
  '   whereParts    - BYREF List(Of String) containing WHERE clause fragments, including
  '                   any leading boolean operators (AND/OR) as required.
  '
  ' Notes:
  '   - This routine injects ONLY SQL fragments. It does NOT perform availability evaluation,
  '     slicing, expansion, or mode precedence. All semantic logic occurs in post-processing.
  '
  '   - This routine is fully responsible for deciding whether to prepend AND/OR or nothing
  '     when adding its own WHERE expression, based on existing whereParts.
  ' ==========================================================================================
  Friend Sub AvailabilityTimeBoundHandler(
    contexts As List(Of TimeBoundFilterContext),
    aliasMap As Dictionary(Of String, String),
    anchorView As String,
    ByRef aliasCounter As Integer,
    ByRef sqlFrom As Text.StringBuilder,
    ByRef selectParts As List(Of String),
    ByRef whereParts As List(Of String)
)

    Try
      ' ------------------------------------------------------------
      ' 1. Determine anchor alias (e.g. V0)
      ' ------------------------------------------------------------
      Dim anchorAlias As String = aliasMap(anchorView)

      ' ------------------------------------------------------------
      ' 2. Allocate alias for the availability pattern table
      ' ------------------------------------------------------------
      Dim patternAlias As String = "V" & aliasCounter
      aliasCounter += 1

      ' ------------------------------------------------------------
      ' 3. JOIN: bring in availability pattern rows
      ' ------------------------------------------------------------
      sqlFrom.AppendLine(
            "LEFT JOIN vwFact_ResourceAvailabilityPattern AS " & patternAlias &
            " ON " & patternAlias & ".ResourceID = " & anchorAlias & ".ResourceID"
        )

      ' ------------------------------------------------------------
      ' 4. SELECT: raw pattern fields needed for post-processing
      ' ------------------------------------------------------------
      selectParts.Add(patternAlias & ".PatternType")
      selectParts.Add(patternAlias & ".PatternEndType")
      selectParts.Add(patternAlias & ".PatternEndAfterOccurrences")
      selectParts.Add(patternAlias & ".Mode")
      selectParts.Add(patternAlias & ".AllDay")
      selectParts.Add(patternAlias & ".StartTime")
      selectParts.Add(patternAlias & ".EndTime")
      selectParts.Add(patternAlias & ".RangeStartDate")
      selectParts.Add(patternAlias & ".RangeEndDate")
      selectParts.Add(patternAlias & ".PatternStartDate")
      selectParts.Add(patternAlias & ".PatternEndDate")
      selectParts.Add(patternAlias & ".RecurWeeks")
      selectParts.Add(patternAlias & ".WeeklyDayOfWeek")
      selectParts.Add(patternAlias & ".MonthlyType")
      selectParts.Add(patternAlias & ".MonthlyDayOfMonth")
      selectParts.Add(patternAlias & ".MonthlyOrdinal")
      selectParts.Add(patternAlias & ".MonthlyDayOfWeek")
      selectParts.Add(patternAlias & ".RecurMonths")

      ' ------------------------------------------------------------
      ' 5. WHERE: coarse date overlap
      ' ------------------------------------------------------------
      Dim fromDate As Object = Nothing
      Dim toDate As Object = Nothing

      For Each ctx In contexts
        Select Case ctx.FieldMap.Name
          Case "AvailabilityFromDate"
            fromDate = ctx.Value
          Case "AvailabilityToDate"
            toDate = ctx.Value
        End Select
      Next

      If fromDate IsNot Nothing AndAlso toDate IsNot Nothing Then

        Dim fromLit = FormatSqlLiteral(fromDate)
        Dim toLit = FormatSqlLiteral(toDate)

        ' DateRange patterns: use RangeStartDate / RangeEndDate
        Dim dateRangeExpr As String =
                "(" &
                patternAlias & ".PatternType = 'DateRange' AND " &
                patternAlias & ".RangeEndDate >= " & fromLit & " AND " &
                patternAlias & ".RangeStartDate <= " & toLit &
                ")"

        ' Weekly / Monthly patterns: use PatternStartDate / PatternEndDate / PatternEndType
        Dim recurringExpr As String =
                "(" &
                patternAlias & ".PatternType IN ('Weekly','Monthly') AND " &
                patternAlias & ".PatternStartDate <= " & toLit & " AND " &
                "(" &
                    patternAlias & ".PatternEndType = 'None' OR " &
                    "(" & patternAlias & ".PatternEndType = 'By' AND " &
                        patternAlias & ".PatternEndDate >= " & fromLit &
                    ")" &
                    " OR " &
                    patternAlias & ".PatternEndType = 'After'" &
                ")" &
                ")"

        ' Preserve resources with no patterns (DefaultMode applies later)
        Dim nullExpr As String = "(" & patternAlias & ".PatternType IS NULL)"

        Dim coreExpr As String =
                "(" & nullExpr & " OR " & dateRangeExpr & " OR " & recurringExpr & ")"

        ' Decide AND/OR/none based on existing whereParts – handler owns this logic
        Dim finalExpr As String
        If whereParts.Count = 0 Then
          finalExpr = coreExpr
        Else
          ' Default to AND; if you ever need OR, change it here consciously
          finalExpr = "AND " & coreExpr
        End If

        whereParts.Add(finalExpr)
      End If

    Catch ex As Exception
      ErrorHandler.UnHandleError(ex)

    Finally
      ' No cleanup required here
    End Try

  End Sub

  ' ==========================================================================================
  ' Routine: ApplyAvailabilitySlicing
  ' Purpose:
  '   Slice Availability patterns and project back to the original user-selected fields,
  '   using STRICT key-based stitching between Base and Availability result sets.
  '
  '   - WholeRange:
  '       * Row count is driven by Base rows
  '       * Availability is resolved per Base row over the full window
  '   - Other slice modes:
  '       * Row count is driven by slice windows
  '       * Base data is cloned into each slice row via key-matched segments
  '
  ' Parameters:
  '   detail          - Rule detail (SelectedValues, Filters)
  '   resultSets      - Dictionary of query results ("Base", "Availability", etc.)
  '                     Each entry is a tuple: (Rows As List(Of Object()), ColNames As List(Of String))
  '   vm              - View map helper
  '   header          - Header list; initially contains only user-selected field names.
  '                     This routine appends headers for any added slice columns.
  '   normalizedArgs  - Runtime parameter values (for parameter-bound filters)
  '
  ' Returns:
  '   List(Of Object()) - Rows shaped as [user fields] + [slice columns]
  '
  ' Notes:
  '   - If no Availability filters, or keys cannot be matched, falls back to
  '     projecting Base rows only (no slicing).
  ' ==========================================================================================
  Friend Function ApplyAvailabilitySlicing(
    detail As UIExcelRuleDesignerRuleRowDetail,
    resultSets As Dictionary(Of String, (Rows As List(Of Object()), ColNames As List(Of String))),
    vm As ExcelRuleViewMapHelper,
    header As List(Of String),
    normalizedArgs As Object()
) As List(Of Object())

    Dim result As New List(Of Object())()

    Try
      ' ----------------------------------------------------------------------
      ' 0. Guard: require Base result set
      ' ----------------------------------------------------------------------
      If Not resultSets.ContainsKey("Base") Then
        Return New List(Of Object())()
      End If

      Dim baseRows = resultSets("Base").Rows
      Dim baseCols = resultSets("Base").ColNames

      If baseRows Is Nothing OrElse baseRows.Count = 0 Then
        Return New List(Of Object())()
      End If

      Dim hasAvailability As Boolean = resultSets.ContainsKey("Availability")
      Dim availRows As List(Of Object()) = Nothing
      Dim availCols As List(Of String) = Nothing

      If hasAvailability Then
        availRows = resultSets("Availability").Rows
        availCols = resultSets("Availability").ColNames
      End If

      ' ----------------------------------------------------------------------
      ' 1. Identify Availability filters
      ' ----------------------------------------------------------------------
      Dim availabilityFilters As New List(Of UIExcelRuleDesignerRuleFilter)()

      For Each f In detail.Filters
        Dim fm = vm.GetField(f.SourceView, f.SourceField)
        If fm IsNot Nothing AndAlso fm.IsTimeBounded AndAlso
               String.Equals(fm.TimeBoundHandler, "Availability", StringComparison.OrdinalIgnoreCase) Then
          availabilityFilters.Add(f)
        End If
      Next

      Dim selectedColNamesForUser As New List(Of String)
      For Each sv In detail.SelectedValues
        selectedColNamesForUser.Add(sv.SourceField)
      Next

      ' No Availability filters or no Availability result ? just project Base rows
      If availabilityFilters.Count = 0 OrElse Not hasAvailability OrElse
           availRows Is Nothing OrElse availRows.Count = 0 Then

        Return ProjectRowsToUserSelection(detail, baseRows, baseCols, selectedColNamesForUser)
      End If

      ' ----------------------------------------------------------------------
      ' 2. Determine slice mode
      ' ----------------------------------------------------------------------
      Dim sliceMode As String =
            availabilityFilters.Select(Function(f) f.SlicingMode).
                                FirstOrDefault(Function(m) Not String.IsNullOrWhiteSpace(m))

      If String.IsNullOrWhiteSpace(sliceMode) Then
        Return ProjectRowsToUserSelection(detail, baseRows, baseCols, selectedColNamesForUser)
      End If

      Dim isWholeRange As Boolean = (sliceMode = "WholeRange")
      Dim isHourly As Boolean = (sliceMode = "Hourly")

      ' ----------------------------------------------------------------------
      ' 3. Extract window (using SafeParseDate)
      ' ----------------------------------------------------------------------
      Dim windowStart As Date? = Nothing
      Dim windowEnd As Date? = Nothing

      ' Track which parameter index we’re consuming for parameter-bound filters
      Dim argIndex As Integer = 0

      For Each f In detail.Filters

        Dim rawValue As String = Nothing

        ' Consume value according to binding
        If String.Equals(f.ValueBinding, "Rule", StringComparison.OrdinalIgnoreCase) Then
          ' Use the literal defined in the rule
          rawValue = f.LiteralValue

        ElseIf String.Equals(f.ValueBinding, "Parameter", StringComparison.OrdinalIgnoreCase) Then
          ' Use the runtime parameter, do NOT morph it into LiteralValue
          If normalizedArgs IsNot Nothing AndAlso argIndex < normalizedArgs.Length AndAlso normalizedArgs(argIndex) IsNot Nothing Then
            rawValue = CStr(normalizedArgs(argIndex))
          End If
          argIndex += 1 ' IMPORTANT: always advance for every parameter-bound filter
        End If

        ' Now decide if this filter is an Availability time-bounded filter
        Dim fm = vm.GetField(f.SourceView, f.SourceField)
        If fm IsNot Nothing AndAlso fm.IsTimeBounded AndAlso
           String.Equals(fm.TimeBoundHandler, "Availability", StringComparison.OrdinalIgnoreCase) Then

          If f.SourceField = "AvailabilityFromDate" AndAlso Not String.IsNullOrWhiteSpace(rawValue) Then
            windowStart = SafeParseDate(rawValue)
          ElseIf f.SourceField = "AvailabilityToDate" AndAlso Not String.IsNullOrWhiteSpace(rawValue) Then
            windowEnd = SafeParseDate(rawValue)
          End If

        End If
      Next


      If Not windowStart.HasValue OrElse Not windowEnd.HasValue Then
        Return ProjectRowsToUserSelection(detail, baseRows, baseCols, selectedColNamesForUser)
      End If

      ' ----------------------------------------------------------------------
      ' 4. DefaultMode
      ' ----------------------------------------------------------------------
      Dim defaultModeRaw As String = LoadMetadataValue("DefaultMode")
      Dim defaultMode As String =
            If(defaultModeRaw.Equals("Unavailable", StringComparison.OrdinalIgnoreCase),
               "Unavailable",
               "Available")

      ' ----------------------------------------------------------------------
      ' 5. Availability selected field?
      ' ----------------------------------------------------------------------
      Dim availabilitySelectedIndex As Integer = -1
      For i = 0 To selectedColNamesForUser.Count - 1
        If selectedColNamesForUser(i).Equals("Availability", StringComparison.OrdinalIgnoreCase) Then
          availabilitySelectedIndex = i
          Exit For
        End If
      Next

      ' Availability filter?
      Dim availabilityFilterValue As String = Nothing
      For Each f In detail.Filters
        If f.SourceField = "Availability" Then
          availabilityFilterValue = f.LiteralValue
        End If
      Next

      ' ----------------------------------------------------------------------
      ' 6. STRICT key-based stitching between Base and Availability
      ' ----------------------------------------------------------------------
      Dim keyColsBase = baseCols.Where(Function(c) c.StartsWith("__Key_", StringComparison.OrdinalIgnoreCase)).ToList()
      Dim keyColsAvail = availCols.Where(Function(c) c.StartsWith("__Key_", StringComparison.OrdinalIgnoreCase)).ToList()

      ' Strict: keys must match exactly
      If keyColsBase.Count = 0 OrElse keyColsAvail.Count = 0 OrElse
           keyColsBase.Count <> keyColsAvail.Count OrElse
           Not keyColsBase.SequenceEqual(keyColsAvail, StringComparer.OrdinalIgnoreCase) Then

        Return ProjectRowsToUserSelection(detail, baseRows, baseCols, selectedColNamesForUser)
      End If

      ' Build lookup: key → Base row
      Dim baseLookup As New Dictionary(Of String, Object())(StringComparer.Ordinal)

      For Each row In baseRows
        Dim k = BuildKeyString(row, baseCols, keyColsBase)
        If Not baseLookup.ContainsKey(k) Then
          baseLookup(k) = row
        End If
      Next

      ' Stitch Availability rows to Base rows via keys
      Dim patterns As New List(Of AvailabilityPatternRow)()

      For Each row In availRows
        Dim k = BuildKeyString(row, availCols, keyColsAvail)
        If Not baseLookup.ContainsKey(k) Then Continue For

        Dim baseRow = baseLookup(k)
        Dim p = ConvertAvailabilityRowToPattern(row, availCols)
        p.OriginalRow = baseRow
        patterns.Add(p)
      Next

      ' No patterns ? just project Base
      If patterns.Count = 0 Then
        Return ProjectRowsToUserSelection(detail, baseRows, baseCols, selectedColNamesForUser)
      End If

      ' ----------------------------------------------------------------------
      ' 7. WholeRange vs sliced modes
      ' ----------------------------------------------------------------------
      If isWholeRange Then
        ' WholeRange: one row per Base row
        Dim segments = ExpandPatternsIntoSegments(patterns, windowStart.Value, windowEnd.Value)

        ' Group segments by OriginalRow reference
        Dim segsByRow As New Dictionary(Of Object(), List(Of AvailabilitySegment))()

        For Each s In segments
          If s.OriginalRow Is Nothing Then Continue For
          If Not segsByRow.ContainsKey(s.OriginalRow) Then
            segsByRow(s.OriginalRow) = New List(Of AvailabilitySegment)()
          End If
          segsByRow(s.OriginalRow).Add(s)
        Next

        For Each baseRow In baseRows
          Dim segsForRow As List(Of AvailabilitySegment) = Nothing
          If Not segsByRow.TryGetValue(baseRow, segsForRow) Then
            segsForRow = New List(Of AvailabilitySegment)()
          End If

          Dim sliceAvailability As String = ResolveAvailability(segsForRow, defaultMode)

          ' Apply Availability filter
          If availabilityFilterValue IsNot Nothing AndAlso
                   Not sliceAvailability.Equals(availabilityFilterValue, StringComparison.OrdinalIgnoreCase) Then
            Continue For
          End If

          Dim baseCount As Integer = selectedColNamesForUser.Count
          Dim outRow(baseCount - 1) As Object

          ' User-selected fields from Base
          For i = 0 To baseCount - 1
            Dim fieldName = selectedColNamesForUser(i)
            ' Special case: Availability is conceptual → not in Base SQL
            If fieldName.Equals("Availability", StringComparison.OrdinalIgnoreCase) Then
              ' Leave placeholder; will be overridden later
              outRow(i) = Nothing
            Else
              Dim idx = baseCols.IndexOf(fieldName)
              If idx >= 0 AndAlso idx < baseRow.Length Then
                outRow(i) = baseRow(idx)
              Else
                outRow(i) = Nothing
              End If
            End If

          Next

          ' Override Availability if selected
          If availabilitySelectedIndex >= 0 Then
            outRow(availabilitySelectedIndex) = sliceAvailability
          End If

          result.Add(outRow)
        Next

        Return result
      End If

      ' ----------------------------------------------------------------------
      ' 8. Non-WholeRange slice modes
      ' ----------------------------------------------------------------------

      ' 8a. Append slice headers
      header.Add("Segment From Date")
      header.Add("Segment To Date")
      If isHourly Then
        header.Add("Segment From Time")
        header.Add("Segment To Time")
      End If

      ' 8b. Expand patterns into segments and build windows
      Dim allSegments = ExpandPatternsIntoSegments(patterns, windowStart.Value, windowEnd.Value)
      Dim windows = GenerateSliceWindows(windowStart.Value, windowEnd.Value, sliceMode)

      ' Group segments by base row
      Dim segsByBase As New Dictionary(Of Object(), List(Of AvailabilitySegment))()
      For Each seg In allSegments
        If seg.OriginalRow Is Nothing Then Continue For
        If Not segsByBase.ContainsKey(seg.OriginalRow) Then
          segsByBase(seg.OriginalRow) = New List(Of AvailabilitySegment)()
        End If
        segsByBase(seg.OriginalRow).Add(seg)
      Next

      ' ----------------------------------------------------------------------
      ' 8c. Build output rows: **BaseRows × SliceWindows**
      ' ----------------------------------------------------------------------
      For Each baseRow In baseRows

        Dim segsForBase As List(Of AvailabilitySegment) = Nothing
        If Not segsByBase.TryGetValue(baseRow, segsForBase) Then
          segsForBase = New List(Of AvailabilitySegment)()
        End If

        For Each w In windows

          ' segments for this base row that intersect this slice window
          Dim sliceSegs = segsForBase.
              Where(Function(s) s.EndDateTime >= w.StartDT AndAlso s.StartDateTime <= w.EndDT).
              ToList()

          Dim sliceAvailability As String = ResolveAvailability(sliceSegs, defaultMode)

          ' Apply Availability filter
          If availabilityFilterValue IsNot Nothing AndAlso
             Not sliceAvailability.Equals(availabilityFilterValue, StringComparison.OrdinalIgnoreCase) Then
            Continue For
          End If

          ' Build output row
          Dim baseCount As Integer = selectedColNamesForUser.Count
          Dim outRow(header.Count - 1) As Object

          ' Copy user-selected fields from baseRow
          For i = 0 To baseCount - 1
            Dim fieldName = selectedColNamesForUser(i)
            If fieldName.Equals("Availability", StringComparison.OrdinalIgnoreCase) Then
              outRow(i) = Nothing
            Else
              Dim idx = baseCols.IndexOf(fieldName)
              outRow(i) = If(idx >= 0 AndAlso idx < baseRow.Length, baseRow(idx), Nothing)
            End If
          Next

          ' Override Availability if selected
          If availabilitySelectedIndex >= 0 Then
            outRow(availabilitySelectedIndex) = sliceAvailability
          End If

          ' Slice columns (real Date/Time, not OADate)
          Dim colIdx As Integer = baseCount
          outRow(colIdx) = w.StartDT.Date : colIdx += 1
          outRow(colIdx) = w.EndDT.Date : colIdx += 1

          If isHourly Then
            outRow(colIdx) = w.StartDT.TimeOfDay.TotalDays : colIdx += 1 ' Convert the time-of-day into an Excel time serial
            outRow(colIdx) = w.EndDT.TimeOfDay.TotalDays : colIdx += 1
          End If

          result.Add(outRow)
        Next
      Next

      Return result
    Catch ex As Exception
      ErrorHandler.UnHandleError(ex)

      Dim selectedColNames As New List(Of String)
      For Each sv In detail.SelectedValues
        selectedColNames.Add(sv.SourceField)
      Next

      Return ProjectRowsToUserSelection(detail, resultSets("Base").Rows, resultSets("Base").ColNames, selectedColNames)

    End Try

  End Function

  ' ==========================================================================================
  ' Routine: BuildKeyString
  ' Purpose:
  '   Build a deterministic composite key string from a row using the given key columns.
  ' Parameters:
  '   row      - Data row
  '   colNames - Column names
  '   keyCols  - Names of key columns (must exist in colNames)
  ' Returns:
  '   String - Composite key (pipe-delimited)
  ' ==========================================================================================
  Private Function BuildKeyString(
    row As Object(),
    colNames As List(Of String),
    keyCols As List(Of String)
) As String

    Dim parts As New List(Of String)

    For Each k In keyCols
      Dim idx = colNames.IndexOf(k)
      If idx >= 0 AndAlso idx < row.Length AndAlso row(idx) IsNot Nothing AndAlso row(idx) IsNot DBNull.Value Then
        parts.Add(row(idx).ToString())
      Else
        parts.Add(String.Empty)
      End If
    Next

    Return String.Join("|", parts)
  End Function

  ' ==========================================================================================
  ' Routine: ConvertAvailabilityRowToPattern
  ' Purpose:
  '   Convert an Availability query row (keys + pattern fields) into an AvailabilityPatternRow.
  '   The OriginalRow is assigned separately by the caller (stitched Base row).
  ' Parameters:
  '   row      - Availability row
  '   colNames - Column names for the Availability query
  ' Returns:
  '   AvailabilityPatternRow - Pattern object (OriginalRow left Nothing)
  ' ==========================================================================================
  Private Function ConvertAvailabilityRowToPattern(
    row As Object(),
    colNames As List(Of String)
) As AvailabilityPatternRow

    Dim p As New AvailabilityPatternRow()

    p.PatternType = GetValue(Of String)(row, colNames, "PatternType")
    p.PatternEndType = GetValue(Of String)(row, colNames, "PatternEndType")
    p.PatternEndAfterOccurrences = GetValue(Of Long?)(row, colNames, "PatternEndAfterOccurrences")
    p.Mode = GetValue(Of String)(row, colNames, "Mode")
    p.AllDay = (GetValue(Of String)(row, colNames, "AllDay") = "1")
    p.StartTime = ParseTime(GetValue(Of String)(row, colNames, "StartTime"))
    p.EndTime = ParseTime(GetValue(Of String)(row, colNames, "EndTime"))
    p.RangeStartDate = ParseDate(GetValue(Of String)(row, colNames, "RangeStartDate"))
    p.RangeEndDate = ParseDate(GetValue(Of String)(row, colNames, "RangeEndDate"))
    p.PatternStartDate = ParseDate(GetValue(Of String)(row, colNames, "PatternStartDate"))
    p.PatternEndDate = ParseDate(GetValue(Of String)(row, colNames, "PatternEndDate"))
    p.RecurWeeks = GetValue(Of Long?)(row, colNames, "RecurWeeks")
    p.WeeklyDayOfWeek = GetValue(Of String)(row, colNames, "WeeklyDayOfWeek")
    p.MonthlyType = GetValue(Of String)(row, colNames, "MonthlyType")
    p.MonthlyDayOfMonth = GetValue(Of Long?)(row, colNames, "MonthlyDayOfMonth")
    p.MonthlyOrdinal = GetValue(Of String)(row, colNames, "MonthlyOrdinal")
    p.MonthlyDayOfWeek = GetValue(Of String)(row, colNames, "MonthlyDayOfWeek")
    p.RecurMonths = GetValue(Of Long?)(row, colNames, "RecurMonths")

    Return p
  End Function

  Private Function ParseTime(ByVal s As String) As TimeSpan
    If String.IsNullOrEmpty(s) Then Return TimeSpan.Zero
    Return TimeSpan.Parse(s)
  End Function
  Private Function ParseDate(ByVal s As String) As DateTime
    If String.IsNullOrEmpty(s) Then Return Date.MinValue
    Return Date.ParseExact(s, "yyyy-MM-dd", Nothing)
  End Function

  ' ==========================================================================================
  ' Routine: ProjectRowsToUserSelection
  ' Purpose:
  '   Project internal SQL rows back to the original user-selected fields only.
  '   Removes all internal/morphed SQL columns and returns rows shaped exactly
  '   as the user requested in SelectedValues.
  '
  ' Parameters:
  '   detail            - Rule detail (SelectedValues)
  '   rawRows           - Raw SQL rows (internal projection)
  '   sqlColNames       - Column names from the SQL query (internal projection)
  '   selectedColNames  - Column names the user actually selected (PreferredName, etc.)
  '
  ' Returns:
  '   List(Of Object()) - Rows containing ONLY the user-selected fields, in order.
  '
  ' Notes:
  '   - This routine is used by ApplyAvailabilitySlicing for all early exits
  '     and for WholeRange mode.
  '   - It NEVER returns internal SQL columns.
  '   - It NEVER adds slice columns.
  ' ==========================================================================================
  Private Function ProjectRowsToUserSelection(
    detail As UIExcelRuleDesignerRuleRowDetail,
    rawRows As List(Of Object()),
    sqlColNames As List(Of String),
    selectedColNames As List(Of String)
) As List(Of Object())

    Dim projected As New List(Of Object())()

    Try
      ' --- Normal execution ---
      Dim baseCount As Integer = selectedColNames.Count

      For Each row In rawRows
        Dim outRow(baseCount - 1) As Object

        For i = 0 To baseCount - 1
          Dim fieldName As String = selectedColNames(i)
          Dim idx As Integer = sqlColNames.IndexOf(fieldName)

          If idx >= 0 AndAlso idx < row.Length Then
            outRow(i) = row(idx)
          Else
            outRow(i) = Nothing
          End If
        Next

        projected.Add(outRow)
      Next

      Return projected

    Catch ex As Exception
      ErrorHandler.UnHandleError(ex)
      Return New List(Of Object())()

    Finally
      ' --- Cleanup ---
      ' Nothing to dispose
    End Try

  End Function

  ' ==========================================================================================
  ' Routine: EnsureColumn
  ' Purpose:
  '   Ensure a column name exists in colNames; if missing, append it and return its index.
  ' Parameters:
  '   colNames - Column names list (modified in place)
  '   name     - Column name to ensure
  ' Returns:
  '   Integer - Index of the column in colNames
  ' Notes:
  '   - Used to append SegmentFromDate/SegmentToDate/SegmentFromTime/SegmentToTime
  ' ==========================================================================================
  Private Function EnsureColumn(colNames As List(Of String), name As String) As Integer

    Try
      Dim idx = colNames.IndexOf(name)
      If idx >= 0 Then Return idx
      colNames.Add(name)
      Return colNames.Count - 1

    Catch ex As Exception
      ErrorHandler.UnHandleError(ex)
      Return Math.Max(0, colNames.Count - 1)

    Finally
      ' --- Cleanup ---
    End Try

  End Function

  ' ==========================================================================================
  ' Routine: ExpandPatternsIntoSegments
  ' Purpose:
  '   Expand AvailabilityPatternRow definitions into concrete AvailabilitySegment intervals
  '   within the requested window.
  ' Parameters:
  '   patterns    - Pattern rows
  '   windowStart - Start of requested window
  '   windowEnd   - End of requested window
  ' Returns:
  '   List(Of AvailabilitySegment) - Concrete availability intervals
  ' Notes:
  '   - Supports DateRange, Weekly, Monthly pattern types
  ' ==========================================================================================
  Private Function ExpandPatternsIntoSegments(
    patterns As List(Of AvailabilityPatternRow),
    windowStart As Date,
    windowEnd As Date
) As List(Of AvailabilitySegment)

    Dim segments As New List(Of AvailabilitySegment)()

    Try
      For Each p In patterns
        Select Case p.PatternType
          Case "DateRange"
            ExpandDateRangePattern(p, windowStart, windowEnd, segments)

          Case "Weekly"
            ExpandWeeklyPattern(p, windowStart, windowEnd, segments)

          Case "Monthly"
            ExpandMonthlyPattern(p, windowStart, windowEnd, segments)

          Case Else
            ' Unknown pattern type → ignore
        End Select
      Next

      Return segments

    Catch ex As Exception
      ErrorHandler.UnHandleError(ex)
      Return New List(Of AvailabilitySegment)()

    Finally
      ' --- Cleanup ---
    End Try

  End Function

  ' ==========================================================================================
  ' Routine: ExpandDateRangePattern
  ' Purpose:
  '   Expand a DateRange pattern into a single AvailabilitySegment clipped to the window.
  ' Parameters:
  '   p           - Pattern row
  '   windowStart - Start of requested window
  '   windowEnd   - End of requested window
  '   segments    - Target list to append segments to
  ' Returns:
  '   None
  ' Notes:
  '   - Respects AllDay vs timed StartTime/EndTime
  ' ==========================================================================================
  Private Sub ExpandDateRangePattern(
    p As AvailabilityPatternRow,
    windowStart As Date,
    windowEnd As Date,
    segments As List(Of AvailabilitySegment)
)

    Try
      If Not p.RangeStartDate.HasValue OrElse Not p.RangeEndDate.HasValue Then Exit Sub

      Dim startDate = p.RangeStartDate.Value
      Dim endDate = p.RangeEndDate.Value

      If endDate < windowStart OrElse startDate > windowEnd Then Exit Sub

      Dim segStart = If(startDate < windowStart, windowStart, startDate)
      Dim segEnd = If(endDate > windowEnd, windowEnd, endDate)

      Dim startDT As DateTime
      Dim endDT As DateTime

      If p.AllDay.GetValueOrDefault(True) Then
        startDT = segStart.Date
        endDT = segEnd.Date.AddDays(1).AddTicks(-1)
      Else
        startDT = segStart.Date + p.StartTime.GetValueOrDefault(TimeSpan.Zero)
        endDT = segEnd.Date + p.EndTime.GetValueOrDefault(TimeSpan.Zero)
      End If

      segments.Add(New AvailabilitySegment With {
            .StartDateTime = startDT,
            .EndDateTime = endDT,
            .Mode = p.Mode,
            .OriginalRow = p.OriginalRow
        })

    Catch ex As Exception
      ErrorHandler.UnHandleError(ex)

    Finally
      ' --- Cleanup ---
    End Try

  End Sub

  ' ==========================================================================================
  ' Routine: ExpandWeeklyPattern
  ' Purpose:
  '   Expand a Weekly pattern into AvailabilitySegments within the requested window.
  ' Parameters:
  '   p           - Pattern row
  '   windowStart - Start of requested window
  '   windowEnd   - End of requested window
  '   segments    - Target list to append segments to
  ' Returns:
  '   None
  ' Notes:
  '   - Uses RecurWeeks and WeeklyDayOfWeek
  ' ==========================================================================================
  Private Sub ExpandWeeklyPattern(
    p As AvailabilityPatternRow,
    windowStart As Date,
    windowEnd As Date,
    segments As List(Of AvailabilitySegment)
)

    Try
      If Not p.PatternStartDate.HasValue Then Exit Sub

      Dim startDate = p.PatternStartDate.Value
      Dim endDate As Date = If(p.PatternEndType = "By" AndAlso p.PatternEndDate.HasValue,
                                 p.PatternEndDate.Value,
                                 windowEnd)

      If endDate < windowStart Then Exit Sub

      Dim recur = p.RecurWeeks.GetValueOrDefault(1)
      Dim targetDay = ParseDayOfWeek(p.WeeklyDayOfWeek)

      Dim d = startDate
      While d.DayOfWeek <> targetDay
        d = d.AddDays(1)
      End While

      While d <= endDate AndAlso d <= windowEnd
        If d >= windowStart Then
          Dim startDT As DateTime
          Dim endDT As DateTime

          If p.AllDay.GetValueOrDefault(True) Then
            startDT = d.Date
            endDT = d.Date.AddDays(1).AddTicks(-1)
          Else
            startDT = d.Date + p.StartTime.GetValueOrDefault(TimeSpan.Zero)
            endDT = d.Date + p.EndTime.GetValueOrDefault(TimeSpan.Zero)
          End If

          segments.Add(New AvailabilitySegment With {
                    .StartDateTime = startDT,
                    .EndDateTime = endDT,
                    .Mode = p.Mode,
                    .OriginalRow = p.OriginalRow
                })
        End If

        d = d.AddDays(7 * recur)
      End While

    Catch ex As Exception
      ErrorHandler.UnHandleError(ex)

    Finally
      ' --- Cleanup ---
    End Try

  End Sub

  ' ==========================================================================================
  ' Routine: ExpandMonthlyPattern
  ' Purpose:
  '   Expand a Monthly pattern into AvailabilitySegments within the requested window.
  ' Parameters:
  '   p           - Pattern row
  '   windowStart - Start of requested window
  '   windowEnd   - End of requested window
  '   segments    - Target list to append segments to
  ' Returns:
  '   None
  ' Notes:
  '   - Supports DayOfMonth and OrdinalDay monthly types
  ' ==========================================================================================
  Private Sub ExpandMonthlyPattern(
    p As AvailabilityPatternRow,
    windowStart As Date,
    windowEnd As Date,
    segments As List(Of AvailabilitySegment)
)

    Try
      If Not p.PatternStartDate.HasValue Then Exit Sub

      Dim d = p.PatternStartDate.Value
      Dim endDate As Date = If(p.PatternEndType = "By" AndAlso p.PatternEndDate.HasValue,
                                 p.PatternEndDate.Value,
                                 windowEnd)

      Dim recur = p.RecurMonths.GetValueOrDefault(1)

      While d <= endDate AndAlso d <= windowEnd

        Dim occurrence As Date? = Nothing

        Select Case p.MonthlyType
          Case "DayOfMonth"
            occurrence = SafeDayOfMonth(d.Year, d.Month, p.MonthlyDayOfMonth)

          Case "OrdinalDay"
            occurrence = FindOrdinalDay(d.Year, d.Month, p.MonthlyOrdinal, p.MonthlyDayOfWeek)
        End Select

        If occurrence.HasValue AndAlso occurrence.Value >= windowStart AndAlso occurrence.Value <= windowEnd Then

          Dim startDT As DateTime
          Dim endDT As DateTime

          If p.AllDay.GetValueOrDefault(True) Then
            startDT = occurrence.Value.Date
            endDT = occurrence.Value.Date.AddDays(1).AddTicks(-1)
          Else
            startDT = occurrence.Value.Date + p.StartTime.GetValueOrDefault(TimeSpan.Zero)
            endDT = occurrence.Value.Date + p.EndTime.GetValueOrDefault(TimeSpan.Zero)
          End If

          segments.Add(New AvailabilitySegment With {
                    .StartDateTime = startDT,
                    .EndDateTime = endDT,
                    .Mode = p.Mode,
                    .OriginalRow = p.OriginalRow
                })
        End If

        d = d.AddMonths(recur)
      End While

    Catch ex As Exception
      ErrorHandler.UnHandleError(ex)

    Finally
      ' --- Cleanup ---
    End Try

  End Sub

  ' ==========================================================================================
  ' Routine: GenerateSliceWindows
  ' Purpose:
  '   Generate slice windows for the requested slice mode over the given window.
  ' Parameters:
  '   windowStart - Start of requested window
  '   windowEnd   - End of requested window
  '   sliceMode   - "WholeRange", "Hourly", "Daily", "Weekly", "Monthly"
  ' Returns:
  '   List(Of SliceWindow) - Time windows for slicing
  ' Notes:
  '   - WholeRange → single window [windowStart, windowEnd]
  ' ==========================================================================================
  Private Function GenerateSliceWindows(
    windowStart As Date,
    windowEnd As Date,
    sliceMode As String
) As List(Of SliceWindow)

    Dim list As New List(Of SliceWindow)()

    Try
      Select Case sliceMode

        Case "WholeRange"
          list.Add(New SliceWindow With {
                    .StartDT = windowStart,
                    .EndDT = windowEnd
                })

        Case "Daily"
          Dim d = windowStart
          While d <= windowEnd
            list.Add(New SliceWindow With {
                        .StartDT = d,
                        .EndDT = d.AddDays(1).AddTicks(-1)
                    })
            d = d.AddDays(1)
          End While

        Case "Weekly"
          Dim d = StartOfWeek(windowStart, DayOfWeek.Monday)
          While d <= windowEnd
            list.Add(New SliceWindow With {
                        .StartDT = d,
                        .EndDT = d.AddDays(7).AddTicks(-1)
                    })
            d = d.AddDays(7)
          End While

        Case "Monthly"
          Dim d = New Date(windowStart.Year, windowStart.Month, 1)
          While d <= windowEnd
            Dim monthEnd = d.AddMonths(1).AddTicks(-1)
            list.Add(New SliceWindow With {
                        .StartDT = d,
                        .EndDT = monthEnd
                    })
            d = d.AddMonths(1)
          End While

        Case "Hourly"
          Dim dt = windowStart
          While dt <= windowEnd
            list.Add(New SliceWindow With {
                        .StartDT = dt,
                        .EndDT = dt.AddHours(1).AddTicks(-1)
                    })
            dt = dt.AddHours(1)
          End While

        Case Else
          Throw New InvalidOperationException("Unknown slice mode: " & sliceMode)

      End Select

      Return list

    Catch ex As Exception
      ErrorHandler.UnHandleError(ex)
      Return New List(Of SliceWindow)()

    Finally
      ' --- Cleanup ---
    End Try

  End Function

  ' ==========================================================================================
  ' Routine: ResolveAvailability
  ' Purpose:
  '   Resolve the Availability value for a slice based on intersecting segments and
  '   DefaultMode, applying the conflict rule (opposite of DefaultMode wins).
  ' Parameters:
  '   sliceSegs   - Segments intersecting the slice window
  '   defaultMode - "Available" or "Unavailable"
  ' Returns:
  '   String - "Available", "Unavailable", or "PartiallyAvailable"
  ' Notes:
  '   - No segments → defaultMode
  '   - All Available → "Available"
  '   - All Unavailable → "Unavailable"
  '   - Mixed → opposite of defaultMode
  ' ==========================================================================================
  Private Function ResolveAvailability(
    sliceSegs As List(Of AvailabilitySegment),
    defaultMode As String
) As String

    Try
      If sliceSegs Is Nothing OrElse sliceSegs.Count = 0 Then
        Return defaultMode
      End If

      Dim allAvailable = sliceSegs.All(Function(s) String.Equals(s.Mode, "Available", StringComparison.OrdinalIgnoreCase))
      Dim allUnavailable = sliceSegs.All(Function(s) String.Equals(s.Mode, "Unavailable", StringComparison.OrdinalIgnoreCase))

      If allAvailable Then Return "Available"
      If allUnavailable Then Return "Unavailable"

      ' Conflict → opposite of defaultMode wins
      If String.Equals(defaultMode, "Available", StringComparison.OrdinalIgnoreCase) Then
        Return "Unavailable"
      Else
        Return "Available"
      End If

    Catch ex As Exception
      ErrorHandler.UnHandleError(ex)
      Return defaultMode

    Finally
      ' --- Cleanup ---
    End Try

  End Function

  ' ==========================================================================================
  ' Routine: GetValue
  ' Purpose:
  '   Safely extract a typed value from a row by column name, handling missing columns and
  '   DBNull.
  ' Parameters:
  '   row      - Data row as Object()
  '   colNames - Column names
  '   name     - Target column name
  ' Returns:
  '   T - Typed value or Nothing if missing/DBNull
  ' Notes:
  '   - Caller must handle Nothing as appropriate
  ' ==========================================================================================
  Private Function GetValue(Of T)(row As Object(), colNames As List(Of String), name As String) As T

    Try
      Dim idx = colNames.IndexOf(name)
      If idx < 0 Then Return Nothing
      Dim v = row(idx)
      If v Is Nothing OrElse v Is DBNull.Value Then Return Nothing
      Return CType(v, T)

    Catch ex As Exception
      ErrorHandler.UnHandleError(ex)
      Return Nothing

    Finally
      ' --- Cleanup ---
    End Try

  End Function

  ' ==========================================================================================
  ' Routine: ParseDayOfWeek
  ' Purpose:
  '   Parse a day-of-week name into a DayOfWeek enum, defaulting to Monday on error.
  ' Parameters:
  '   name - Day name (e.g. "Monday")
  ' Returns:
  '   DayOfWeek - Parsed value or Monday if invalid
  ' Notes:
  '   - Case-insensitive
  ' ==========================================================================================
  Private Function ParseDayOfWeek(name As String) As DayOfWeek

    Try
      If String.IsNullOrWhiteSpace(name) Then Return DayOfWeek.Monday
      Return CType([Enum].Parse(GetType(DayOfWeek), name, True), DayOfWeek)

    Catch ex As Exception
      ErrorHandler.UnHandleError(ex)
      Return DayOfWeek.Monday

    Finally
      ' --- Cleanup ---
    End Try

  End Function

  ' ==========================================================================================
  ' Routine: SafeDayOfMonth
  ' Purpose:
  '   Safely construct a date for a given year, month, and day, clamping the day to the
  '   number of days in the month.
  ' Parameters:
  '   year  - Year
  '   month - Month
  '   day   - Desired day (nullable)
  ' Returns:
  '   Date? - Valid date or Nothing if day is null/invalid
  ' Notes:
  '   - Day <= 0 → Nothing
  '   - Day > days in month → clamped to last day
  ' ==========================================================================================
  Private Function SafeDayOfMonth(year As Integer, month As Integer, day As Integer?) As Date?

    Try
      If Not day.HasValue OrElse day.Value <= 0 Then Return Nothing
      Dim daysInMonth = Date.DaysInMonth(year, month)
      Dim d = Math.Min(day.Value, daysInMonth)
      Return New Date(year, month, d)

    Catch ex As Exception
      ErrorHandler.UnHandleError(ex)
      Return Nothing

    Finally
      ' --- Cleanup ---
    End Try

  End Function

  ' ==========================================================================================
  ' Routine: FindOrdinalDay
  ' Purpose:
  '   Find the date of an ordinal weekday in a given month (e.g. "second Tuesday").
  ' Parameters:
  '   year    - Year
  '   month   - Month
  '   ordinal - "first", "second", "third", "fourth", "last"
  '   dayName - Day name (e.g. "Tuesday")
  ' Returns:
  '   Date? - Matching date or Nothing if invalid
  ' Notes:
  '   - "last" walks backwards from end of month
  ' ==========================================================================================
  Private Function FindOrdinalDay(year As Integer, month As Integer, ordinal As String, dayName As String) As Date?

    Try
      If String.IsNullOrWhiteSpace(ordinal) OrElse String.IsNullOrWhiteSpace(dayName) Then Return Nothing

      Dim targetDay = ParseDayOfWeek(dayName)
      Dim firstOfMonth As New Date(year, month, 1)

      Dim d = firstOfMonth
      While d.DayOfWeek <> targetDay
        d = d.AddDays(1)
      End While

      Dim offset As Integer
      Select Case ordinal.ToLowerInvariant()
        Case "first" : offset = 0
        Case "second" : offset = 7
        Case "third" : offset = 14
        Case "fourth" : offset = 21
        Case "last"
          Dim nextMonth = firstOfMonth.AddMonths(1)
          Dim last = nextMonth.AddDays(-1)
          While last.DayOfWeek <> targetDay
            last = last.AddDays(-1)
          End While
          Return last
        Case Else
          Return Nothing
      End Select

      Dim candidate = d.AddDays(offset)
      If candidate.Month <> month Then Return Nothing
      Return candidate

    Catch ex As Exception
      ErrorHandler.UnHandleError(ex)
      Return Nothing

    Finally
      ' --- Cleanup ---
    End Try

  End Function

  ' ==========================================================================================
  ' Routine: StartOfWeek
  ' Purpose:
  '   Compute the start of the week for a given date and first-day-of-week.
  ' Parameters:
  '   d        - Input date
  '   firstDay - First day of week (e.g. Monday)
  ' Returns:
  '   Date - Start-of-week date
  ' Notes:
  '   - Returns date with time truncated
  ' ==========================================================================================
  Private Function StartOfWeek(d As Date, firstDay As DayOfWeek) As Date

    Try
      Dim diff = (7 + (d.DayOfWeek - firstDay)) Mod 7
      Return d.AddDays(-diff).Date

    Catch ex As Exception
      ErrorHandler.UnHandleError(ex)
      Return d.Date

    Finally
      ' --- Cleanup ---
    End Try

  End Function

  ' ==========================================================================================
  ' Types: AvailabilityPatternRow, AvailabilitySegment, AvailabilitySlice, SliceWindow
  ' Purpose:
  '   Internal models for pattern expansion, segment representation, and slice shaping.
  ' ==========================================================================================
  Friend Class AvailabilityPatternRow
    Public Property OriginalRow As Object()

    Public Property PatternType As String
    Public Property PatternEndType As String
    Public Property PatternEndAfterOccurrences As Integer?
    Public Property Mode As String
    Public Property AllDay As Boolean?
    Public Property StartTime As TimeSpan?
    Public Property EndTime As TimeSpan?

    Public Property RangeStartDate As Date?
    Public Property RangeEndDate As Date?

    Public Property PatternStartDate As Date?
    Public Property PatternEndDate As Date?

    Public Property RecurWeeks As Integer?
    Public Property WeeklyDayOfWeek As String

    Public Property MonthlyType As String
    Public Property MonthlyDayOfMonth As Integer?
    Public Property MonthlyOrdinal As String
    Public Property MonthlyDayOfWeek As String
    Public Property RecurMonths As Integer?
  End Class

  Friend Class AvailabilitySegment
    Public Property StartDateTime As DateTime
    Public Property EndDateTime As DateTime
    Public Property Mode As String
    Public Property OriginalRow As Object()
  End Class

  Friend Class AvailabilitySlice
    Public Property SliceStart As DateTime
    Public Property SliceEnd As DateTime
    Public Property Availability As String
    Public Property OriginalRow As Object()
  End Class

  Friend Class SliceWindow
    Public Property StartDT As DateTime
    Public Property EndDT As DateTime
  End Class
End Module
