Imports System.Drawing
Imports System.Text.Json
Imports System.Windows.Forms
Imports ExcelDna.Integration
Imports Microsoft.Data.Sqlite
Imports Microsoft.Office.Interop
Imports Microsoft.Office.Interop.Excel
Imports ResourceManagement.WindowsAPIs
Imports Excel = Microsoft.Office.Interop.Excel

Friend Class QuerySpec
  Public Property AliasMap As Dictionary(Of String, String)
  Public Property JoinGraph As System.Text.StringBuilder
  Public Property WhereParts As List(Of String)
  Public Property SelectParts As List(Of String)
  Public Property AnchorView As String
  Public Property AliasCounter As Integer
End Class

Public Module ExcelRuleEngine

  ' ==========================================================================================
  ' Routine:    GetRuleValues
  ' Purpose:
  '   Evaluate a rule using the new multi-query architecture. This routine:
  '       - Resolves the rule by ID or name
  '       - Loads and validates the rule definition
  '       - Builds ALL required SQL queries (Base + handler queries) via BuildRuleQuery
  '       - Executes each query and collects their result sets
  '       - Passes all result sets and query specifications to post-processing
  '         (e.g., ApplyAvailabilitySlicing) for handler-driven transformations
  '       - Removes internal surrogate key columns (__Key_*) before shaping the final output
  '       - Returns a spill-compatible result:
  '             * Scalar (Object)
  '             * Vertical list (Object())
  '             * Table (Object(,))
  ' Parameters:
  '   ruleId        - String (ByRef)
  '                   Canonical RuleID (GUID). If supplied, takes precedence over ruleName.
  '   ruleName      - String (ByRef)
  '                   User-facing rule name. Used only when ruleId is blank. Updated to the
  '                   resolved canonical name on success.
  '   args          - Object()
  '                   Runtime parameter values for the rule. Passed to BuildRuleQuery for
  '                   parameter substitution.
  '   includeHeaders - Boolean (Optional)
  '                   When True, the returned table includes a header row of user-visible
  '                   column names. Internal key columns are always removed.
  ' Returns:
  '   Object
  '       One of:
  '           - Scalar:     Object
  '           - List:       Object()
  '           - Table:      Object(,)
  '       Error codes:
  '           "#RuleNotFound"   → RuleID invalid or ruleName not found
  '           "#BadRule"        → Rule definition missing or malformed
  '           "#BadParameters"  → Runtime parameters invalid
  '           "#NoQuery"        → SQL generation failed
  '           "#NoData"         → Base query returned zero rows
  '           "#Error"          → Unexpected exception (logged)
  ' Notes:
  '   - Internal surrogate keys (__Key_*) are ALWAYS included in SQL queries but are ALWAYS
  '     removed before returning results to Excel.
  '   - Multiple handler queries may be generated (Availability, Skills, etc.). All are executed
  '     and passed to post-processing.
  '   - Post-processing routines must use querySpecs to understand structure, keys, and handler
  '     metadata.
  '   - This routine replaces all previous single-query rule evaluation paths.
  ' ==========================================================================================
  Public Function GetRuleValues(ByRef ruleId As String, ByRef ruleName As String,
                              args() As Object, Optional includeHeaders As Boolean = False) As Object

    Dim conn As SQLiteConnectionWrapper = Nothing
    Dim cmd As SQLiteCommandWrapper = Nothing
    Dim ds As SQLiteDataRowReader = Nothing

    Try
      ' ----------------------------------------------------------------------
      ' Resolve RuleID
      ' ----------------------------------------------------------------------
      Dim resolvedRuleId As String = Nothing

      If Not String.IsNullOrWhiteSpace(ruleId) Then
        resolvedRuleId = ruleId
      ElseIf Not String.IsNullOrWhiteSpace(ruleName) Then
        resolvedRuleId = LookupRuleIdByName(ruleName)
      End If

      If String.IsNullOrWhiteSpace(resolvedRuleId) Then
        Return "#RuleNotFound"
      End If

      ' ----------------------------------------------------------------------
      ' Load rule record
      ' ----------------------------------------------------------------------
      conn = OpenDatabase(AddInContext.Current.Config.ActiveDbPath)

      Dim rec As RecordExcelRule =
            RecordLoader.LoadRecord(Of RecordExcelRule)(conn, {resolvedRuleId})

      If rec Is Nothing OrElse rec.IsDeleted Then
        Return "#RuleNotFound"
      End If

      ruleName = rec.RuleName
      ruleId = resolvedRuleId

      ' ----------------------------------------------------------------------
      ' Deserialize JSON
      ' ----------------------------------------------------------------------
      Dim detail As UIExcelRuleDesignerRuleRowDetail =
            JsonSerializer.Deserialize(Of UIExcelRuleDesignerRuleRowDetail)(
                rec.DefinitionJson,
                New JsonSerializerOptions With {.PropertyNameCaseInsensitive = True})

      If detail Is Nothing OrElse detail.SelectedValues Is Nothing OrElse detail.SelectedValues.Count = 0 Then
        Return "#BadRule"
      End If

      ' ----------------------------------------------------------------------
      ' Validate runtime parameters
      ' ----------------------------------------------------------------------
      Dim normalizedArgs() As Object = Nothing
      If Not ValidateRuntimeParameters(detail, args, normalizedArgs) Then
        Return "#BadParameters"
      End If

      ' ----------------------------------------------------------------------
      ' Build SelectFields list
      ' ----------------------------------------------------------------------
      Dim sfs As New List(Of SelectFields)
      For Each sv In detail.SelectedValues
        sfs.Add(New SelectFields With {
                .sourceView = sv.SourceView,
                .sourceField = sv.SourceField,
                .fieldID = sv.FieldID,
                .listTypeID = sv.ListTypeID
            })
      Next

      Dim anchorView As String = detail.SelectedValues(0).View

      ' ----------------------------------------------------------------------
      ' Build ALL SQL queries (Base + handlers)
      ' ----------------------------------------------------------------------
      Dim vm As New ExcelRuleViewMapHelper(ExcelRuleViewMapLoader.LoadExcelRuleViewMap())

      Dim sqlQueries As Dictionary(Of String, String) = Nothing
      Dim querySpecs As Dictionary(Of String, QuerySpec) = Nothing

      vm.BuildRuleQuery(anchorView, sfs, detail.Filters, normalizedArgs,
                          sqlQueries, querySpecs)

      If sqlQueries Is Nothing OrElse sqlQueries.Count = 0 Then
        Return "#NoQuery"
      End If

      ' ----------------------------------------------------------------------
      ' Execute ALL queries
      ' ----------------------------------------------------------------------
      Dim resultSets As New Dictionary(Of String, (Rows As List(Of Object()), ColNames As List(Of String)))

      For Each kvp In sqlQueries
        Dim queryName = kvp.Key
        Dim sql = kvp.Value

        If String.IsNullOrWhiteSpace(sql) Then Continue For

        cmd = conn.CreateCommand(sql)
        ds = cmd.OpenDataSet()

        Dim rows As New List(Of Object())
        Dim colNames As List(Of String) = Nothing

        While ds.Read()
          Dim row = ds.Row

          If colNames Is Nothing Then
            colNames = row.FieldNames.ToList()
          End If

          Dim values(colNames.Count - 1) As Object
          For i = 0 To colNames.Count - 1
            values(i) = row.GetValue(colNames(i))
          Next

          rows.Add(values)
        End While

        resultSets(queryName) = (rows, colNames)
      Next

      ' ----------------------------------------------------------------------
      ' Validate Base result
      ' ----------------------------------------------------------------------
      If Not resultSets.ContainsKey("Base") OrElse resultSets("Base").Rows.Count = 0 Then
        Return "#NoData"
      End If

      Dim baseRows = resultSets("Base").Rows
      Dim baseCols = resultSets("Base").ColNames

      Dim rowCount = baseRows.Count
      Dim colCount = baseCols.Count

      ' ----------------------------------------------------------------------
      ' Build user-visible headers
      ' ----------------------------------------------------------------------
      Dim header As New List(Of String)
      Dim selectedColNames As New List(Of String)

      For Each sv In detail.SelectedValues
        header.Add(vm.GetFieldDisplayName(sv.SourceView, sv.SourceField))
        selectedColNames.Add(sv.SourceField)
      Next

      ' ----------------------------------------------------------------------
      ' Post-processing (Availability, etc.)
      ' Pass ALL result sets + ALL querySpecs
      ' ----------------------------------------------------------------------
      Dim finalRows As List(Of Object()) =
          ApplyAvailabilitySlicing(
              detail,
              resultSets,
              vm,
              header,
              normalizedArgs
          )

      ' ----------------------------------------------------------------------
      ' Strip internal key columns (__Key_*)
      ' ----------------------------------------------------------------------
      Dim keepIndexes As New List(Of Integer)

      For i = 0 To header.Count - 1
        If Not header(i).StartsWith("__Key_", StringComparison.OrdinalIgnoreCase) Then
          keepIndexes.Add(i)
        End If
      Next

      Dim cleanedHeader As New List(Of String)
      For Each idx In keepIndexes
        cleanedHeader.Add(header(idx))
      Next
      header = cleanedHeader

      Dim cleanedRows As New List(Of Object())
      For Each r In finalRows
        Dim newRow(keepIndexes.Count - 1) As Object
        For i = 0 To keepIndexes.Count - 1
          newRow(i) = r(keepIndexes(i))
        Next
        cleanedRows.Add(newRow)
      Next

      finalRows = cleanedRows

      ' ----------------------------------------------------------------------
      ' Shape result for Excel
      ' ----------------------------------------------------------------------
      rowCount = finalRows.Count
      colCount = header.Count

      ' Scalar
      If rowCount = 1 AndAlso colCount = 1 Then
        Return finalRows(0)(0)
      End If

      ' Vertical list
      If colCount = 1 Then
        Dim arr(rowCount - 1) As Object
        For i = 0 To rowCount - 1
          arr(i) = finalRows(i)(0)
        Next
        Return arr
      End If

      ' Table
      Dim table(rowCount - 1, colCount - 1) As Object
      For r = 0 To rowCount - 1
        For c = 0 To colCount - 1
          table(r, c) = finalRows(r)(c)
        Next
      Next

      If includeHeaders Then
        Dim withHeaders(rowCount, colCount - 1) As Object

        For c = 0 To colCount - 1
          withHeaders(0, c) = header(c)
        Next

        For r = 0 To rowCount - 1
          For c = 0 To colCount - 1
            withHeaders(r + 1, c) = finalRows(r)(c)
          Next
        Next

        Return withHeaders
      End If

      Return table

    Catch ex As Exception
      ErrorHandler.UnHandleError(ex)
      Return "#Error"

    Finally
      If ds IsNot Nothing Then ds.Dispose()
      If conn IsNot Nothing Then conn.Close()
    End Try

  End Function

  ' ==========================================================================================
  ' Routine: ValidateRuleParameterType
  ' Purpose:
  '   Validates that a parameter value supplied for a field is compatible with the field's
  '   declared DataType and the selected ExcelRefType. Ensures literals parse correctly and
  '   referenced cells contain values of the correct type.
  ' Parameters:
  '   field -         The ExcelRuleViewMapField describing the field being filtered (contains DataType).
  '   refType -       The ExcelRefType selected by the user (Literal, Address, Name, Offset).
  '   literalValue -  The literal value entered by the user when refType = Literal. Ignored otherwise.
  '   refValue -      The reference string (address or name) entered by the user when refType is
  '                   Address, Name, or Offset. Ignored when refType = Literal.
  ' Returns:
  '   Boolean -       True  if the parameter is valid for the field's DataType and refType.
  '                   False if the parameter is invalid or cannot be resolved.
  ' Notes:
  '   - Literal values must parse according to the field's DataType:
  '         * Date    → Date.TryParse
  '         * Number  → Double.TryParse
  '         * Boolean → Boolean.TryParse
  '         * Text    → Always valid
  '   - Address/Name/Offset references must resolve to a valid Excel.Range using
  '     TryResolveRange, and the resolved cell value must match the field's DataType.
  '   - This routine performs no UI messaging; callers must display user-facing errors.
  ' ==========================================================================================
  Public Function ValidateRuleParameterType(dataType As String,
                                       refType As String,
                                       literalValue As String,
                                       refValue As String) As Boolean

    Try

      ' -------------------------------
      ' Literal validation
      ' -------------------------------
      If refType = ExcelRefType.Literal.ToString() Then
        If String.IsNullOrWhiteSpace(literalValue) Then Return False
        Select Case dataType
          Case "date"
            Dim tmpDate As Date
            If Not Date.TryParse(literalValue, tmpDate) Then Return False
          Case "number"
            Dim tmpNum As Double
            If Not Double.TryParse(literalValue, tmpNum) Then Return False
          Case "boolean"
            Dim tmpBool As Boolean
            If Not Boolean.TryParse(literalValue, tmpBool) Then Return False
          Case "text"
            ' Always valid
            Return True
          Case Else
            ' Unknown datatype → reject
            Return False
        End Select
        Return True
      End If

      ' -------------------------------
      ' Address / Name / Offset validation
      ' -------------------------------
      If refType = ExcelRefType.Address.ToString() OrElse
         refType = ExcelRefType.Name.ToString() OrElse
         refType = ExcelRefType.Offset.ToString() Then

        If String.IsNullOrWhiteSpace(refValue) Then Return False
        ' Try to resolve the reference
        Dim rng As Excel.Range = Nothing
        If Not TryResolveRange(refValue, rng) Then
          Return False
        End If
        If rng Is Nothing Then Return False
        ' Validate cell contents based on DataType
        Select Case dataType
          Case "date"
            Dim tmpDate As Date
            Dim v = rng.Value2
            If TypeOf v Is Double Then
              ' Excel serial date
              tmpDate = Date.FromOADate(CDbl(v))
              Return True
            ElseIf TypeOf v Is String Then
              ' User typed a literal string
              If Date.TryParse(CStr(v), tmpDate) Then
                Return True
              Else
                Return False
              End If
            ElseIf IsDate(v) Then
              ' Excel sometimes returns Date directly
              tmpDate = CDate(v)
              Return True
            Else
              Return False
            End If
          Case "number"
            Dim tmpNum As Double
            If Not Double.TryParse(Convert.ToString(rng.Value2), tmpNum) Then Return False
          Case "boolean"
            Dim tmpBool As Boolean
            Dim v = rng.Value2
            If TypeOf v Is Boolean Then
              tmpBool = CBool(v)
              Return True
            ElseIf TypeOf v Is Double Then
              ' Excel TRUE/FALSE stored as 1/0
              If CDbl(v) = 1 Then
                tmpBool = True
                Return True
              ElseIf CDbl(v) = 0 Then
                tmpBool = False
                Return True
              Else
                Return False
              End If
            ElseIf TypeOf v Is String Then
              ' Accept case-insensitive true/false
              If Boolean.TryParse(CStr(v), tmpBool) Then
                Return True
              Else
                Return False
              End If
            Else
              Return False
            End If
          Case "text"
            ' Always valid
            Return True
          Case Else
            Return False
        End Select
        Return True
      End If

      ' -------------------------------
      ' Unknown refType → reject
      ' -------------------------------
      Return False
    Catch ex As Exception
      ErrorHandler.UnHandleError(ex)
      Return False
    Finally
      ' --- Cleanup ---
      ' Dispose objects, close connections, release handles, etc.
    End Try
  End Function

  ' ==========================================================================================
  ' Routine: TryResolveRange
  ' Purpose:
  '   Safely attempts to resolve a user-supplied reference string into an Excel.Range.
  ' Parameters:
  '   refValue -  The user-supplied reference string (A1-style address or named range).
  '   rng -       [ByRef] Receives the resolved Excel.Range if successful; otherwise Nothing.
  ' Returns:
  '   Boolean - True  if the reference resolves to a valid Excel.Range.
  '             False if the reference cannot be resolved.
  ' Notes:
  '   - Supports:
  '       * A1-style addresses (e.g., "A1", "Sheet1!B3:C10").
  '       * Workbook-scoped named ranges.
  '       * Worksheet-scoped named ranges.
  '   - Never throws; all COM exceptions are caught and suppressed.
  '   - Does not validate the contents or data type of the range; callers must perform
  '     type validation separately (e.g., in ValidateParameterType).
  '   - Offset expressions are not interpreted here; callers must resolve offsets before
  '     calling this routine.
  ' ==========================================================================================
  Private Function TryResolveRange(refValue As String,
                                 ByRef rng As Excel.Range) As Boolean
    rng = Nothing
    Dim app As Excel.Application = Nothing
    Dim wb As Excel.Workbook = Nothing
    Try

      app = CType(ExcelDnaUtil.Application, Excel.Application)
      wb = app.ActiveWorkbook

      ' --- Try A1-style address ---
      Try
        rng = app.Range(refValue)
        If rng IsNot Nothing Then Return True
      Catch
        ' ignore and continue
      End Try

      ' --- Try workbook-scoped name ---
      Try
        Dim nm As Excel.Name = wb.Names(refValue)
        If nm IsNot Nothing Then
          rng = nm.RefersToRange
          If rng IsNot Nothing Then Return True
        End If
      Catch
        ' ignore and continue
      End Try

      ' --- Try worksheet-scoped names ---
      For Each ws As Excel.Worksheet In wb.Worksheets
        Try
          Dim nm As Excel.Name = ws.Names(refValue)
          If nm IsNot Nothing Then
            rng = nm.RefersToRange
            If rng IsNot Nothing Then Return True
          End If
        Catch
          ' ignore and continue
        End Try
      Next

      ' --- Offset expressions (optional) ---
      ' If you support expressions like "A1+1,0" or similar, parse here.

      Return False
    Catch ex As Exception
      ErrorHandler.UnHandleError(ex)
      Return False
    Finally
      ' --- Cleanup ---
      ' Dispose objects, close connections, release handles, etc.
      app = Nothing
      wb = Nothing
    End Try
  End Function

  ' ==========================================================================================
  ' Routine: ValidateRuntimeParameters
  ' Purpose:
  '   Validates the parameters supplied to GetRuleValues at runtime (formula call).
  '   Ensures:
  '     - Parameter count matches number of filters with a FieldOperator.
  '     - Each parameter is compatible with the field's DataType.
  '     - Range / Name references resolve correctly.
  ' Parameters:
  '   detail -
  '       The rule metadata loaded from JSON.
  '   args() -
  '       The runtime parameters supplied by the Excel formula.
  '   normalizedArgs() [ByRef] -
  '       An output array that receives normalized parameter values (e.g., dates converted to "yyyy-MM-dd",
  ' Returns:
  '   Boolean -
  '       True  if all parameters are valid.
  '       False if any parameter is invalid (and displays a message).
  ' Notes:
  '   - This is the ONLY place where formula parameters can be validated.
  '   - Rules do NOT store literal or reference values.
  ' ==========================================================================================
  Private Function ValidateRuntimeParameters(detail As UIExcelRuleDesignerRuleRowDetail,
                                           args() As Object, ByRef normalizedArgs() As Object) As Boolean

    Dim vm As New ExcelRuleViewMapHelper(ExcelRuleViewMapLoader.LoadExcelRuleViewMap())

    ' 1. Count how many filters actually consume a parameter
    Dim expected As Integer = 0
    For Each f In detail.Filters
      If String.IsNullOrWhiteSpace(f.FieldOperator) Then
        Continue For
      End If
      ' Skip rule-bound filters
      If String.Equals(f.ValueBinding, ValueBinding.Rule.ToString(),
                       StringComparison.OrdinalIgnoreCase) Then
        Continue For
      End If
      expected += 1
    Next

    If args Is Nothing Then args = Array.Empty(Of Object)()

    If args.Length <> expected Then
      MessageBox.Show($"Rule expects {expected} parameters but received {args.Length}.",
                    "Rule Values",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Warning)
      Return False
    End If

    ' Allocate normalized args
    ReDim normalizedArgs(expected - 1)

    ' 2. Validate each parameter as a literal against the field's DataType
    Dim argIndex As Integer = 0

    For Each f In detail.Filters

      If String.IsNullOrWhiteSpace(f.FieldOperator) Then
        Continue For
      End If

      ' Skip rule-bound filters
      If String.Equals(f.ValueBinding, ValueBinding.Rule.ToString(),
                       StringComparison.OrdinalIgnoreCase) Then
        Continue For
      End If

      ' Get field metadata
      Dim fieldMap As ExcelRuleViewMapField = vm.GetField(f.SourceView, f.SourceField)

      Dim dataType As String = fieldMap.DataType
      If dataType IsNot Nothing Then
        dataType = dataType.Trim().ToLowerInvariant()
      End If

      Dim rawArg As Object = args(argIndex)
      Dim literalValue As String = ""
      ' --- Normalization block ---
      If rawArg Is Nothing OrElse TypeOf rawArg Is ExcelEmpty Then
        literalValue = ""
      ElseIf TypeOf rawArg Is Double Then
        ' Could be a number OR an Excel serial date OR a boolean (1/0)
        Dim d As Double = CDbl(rawArg)
        Select Case dataType
          Case "date"
            literalValue = Date.FromOADate(d).ToString("yyyy-MM-dd")
          Case "boolean"
            If d = 1 Then
              literalValue = "true"
            ElseIf d = 0 Then
              literalValue = "false"
            Else
              literalValue = d.ToString()   ' invalid, will fail validation
            End If
          Case Else
            ' number or text-as-number
            literalValue = d.ToString()
        End Select
      ElseIf TypeOf rawArg Is Boolean Then
        literalValue = If(CBool(rawArg), "true", "false")
      Else
        ' String or anything else
        literalValue = rawArg.ToString().Trim()
      End If
      ' --- End normalization block ---

      ' Validate the normalized literal
      If Not ValidateRuleParameterType(dataType, ExcelRefType.Literal.ToString(), literalValue, "") Then
        'MessageBox.Show($"Parameter {argIndex + 1} is not valid for field '{f.Field}'.",
        '              "Rule Values",
        '              MessageBoxButtons.OK,
        '              MessageBoxIcon.Warning)
        Return False
      End If

      ' Store normalized literal for SQL builder
      normalizedArgs(argIndex) = literalValue

      argIndex += 1
    Next

    Return True

  End Function

  ' ==========================================================================================
  ' Routine: DeserializeRuntimeRule
  ' Purpose:
  '   Convert DefinitionJson into a runtime DTO (RecordExcelRuleRowDetail).
  ' ==========================================================================================
  Private Function DeserializeRuntimeRule(json As String) As RecordExcelRuleRowDetail
    If String.IsNullOrWhiteSpace(json) Then
      Return New RecordExcelRuleRowDetail()
    End If

    Dim options As New JsonSerializerOptions With {
      .PropertyNameCaseInsensitive = True
    }

    Return JsonSerializer.Deserialize(Of RecordExcelRuleRowDetail)(json, options)
  End Function

  ' ==========================================================================================
  ' Routine:    LookupRuleIdByName
  ' Purpose:
  '   Resolve a rule's primary key (RuleID) from its user-facing RuleName.
  '   This is used when callers supply a rule name instead of the canonical ID.
  '
  ' Parameters:
  '   ruleName - String
  '       The display name of the rule as entered by the user or stored in a cell.
  '
  ' Returns:
  '   String
  '       The RuleID (primary key) if found.
  '       Nothing if no matching rule exists or the rule is marked as deleted.
  '
  ' Notes:
  '   - Uses the standard SQLite wrapper pattern:
  '         CreateCommand → AddParameters → SetParameter → OpenDataSet
  '   - Only returns active (IsDeleted = 0) rules.
  '   - Caller is responsible for handling a Nothing return value.
  ' ==========================================================================================
  Private Function LookupRuleIdByName(ruleName As String) As String
    Dim conn As SQLiteConnectionWrapper = Nothing
    Dim cmd As SQLiteCommandWrapper = Nothing
    Dim ds As SQLiteDataRowReader = Nothing

    Try
      conn = OpenDatabase(AddInContext.Current.Config.ActiveDbPath)

      Dim sql As String =
      "SELECT RuleID " &
      "FROM tblExcelRule " &
      "WHERE RuleName = @RuleName COLLATE NOCASE " &
      "LIMIT 1"

      cmd = conn.CreateCommand(sql)

      ' Declare parameter(s) by name
      cmd.AddParameters({"RuleName"})

      ' Set value
      cmd.SetParameter("RuleName", ruleName)

      ds = cmd.OpenDataSet()

      If ds.Read() Then
        Dim row = ds.Row
        Return CStr(row.GetValue("RuleID"))
      End If

      Return Nothing

    Catch ex As Exception
      ErrorHandler.UnHandleError(ex)
      Return Nothing

    Finally
      If ds IsNot Nothing Then ds.Dispose()
      If conn IsNot Nothing Then conn.Close()
    End Try
  End Function

  ' ==========================================================================================
  ' Routine:    FormatSqlLiteral
  ' Purpose:
  '   Convert a .NET value into a safe SQL literal for inline SQL.
  '
  ' Parameters:
  '   value - Object
  '           The value to format.
  '
  ' Returns:
  '   String - A SQL literal (quoted string, number, or NULL).
  '
  ' Notes:
  '   - Strings are single-quoted and escaped.
  '   - Booleans become 1/0.
  '   - NULL becomes NULL.
  ' ==========================================================================================
  Friend Function FormatSqlLiteral(value As Object) As String

    Try
      ' --- Normal execution ---
      If value Is Nothing OrElse value Is DBNull.Value Then
        Return "NULL"
      End If

      If TypeOf value Is String Then
        Dim s As String = CStr(value).Replace("'", "''")
        Return "'" & s & "'"
      End If

      If TypeOf value Is Boolean Then
        Return If(CBool(value), "1", "0")
      End If

      Return Convert.ToString(value, Globalization.CultureInfo.InvariantCulture)

    Catch ex As Exception
      ErrorHandler.UnHandleError(ex)
      Return "NULL"

    Finally
      ' --- Cleanup ---
      ' No disposable resources
    End Try

  End Function

End Module
