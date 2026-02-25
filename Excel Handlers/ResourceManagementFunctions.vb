Imports ExcelDna.Integration

Namespace ResourceManagement.ExcelFunctions
  Public Module ResourceManagementFunctions
    ' Routines need to be Public

    <ExcelFunction(
      Name:="ResourceManagerRuleValues",
      Description:="Returns the specified rule values",
      Category:="Resource Management"
    )>
    Public Function ResourceManagerRuleValues(
      <ExcelArgument(Description:="The name of the rule to evaluate.")>
      ruleName As String,
      <ExcelArgument(Description:="TRUE to include column headers, FALSE to omit them.")>
       includeHeaders As Boolean,
      <ExcelArgument(Description:="Optional parameters required by the rule.")>
      ParamArray args() As Object
    ) As Object
      Dim result = ExcelRuleEngine.GetRuleValues("", ruleName, args, includeHeaders)
      ' Excel-DNA interprets a returned Object() from a function with a ParamArray as a horizontal vector.
      ' So we need to convert a 1D array to a 2D array with one column to get a vertical vector in Excel.
      If TypeOf result Is Object() Then
        Dim arr1d = DirectCast(result, Object())
        Dim arr2d(arr1d.Length - 1, 0) As Object
        For i = 0 To arr1d.Length - 1
          arr2d(i, 0) = arr1d(i)
        Next
        Return arr2d
      Else
        Return result
      End If
    End Function

    Private Function BuildViewMap() As Dictionary(Of String, String)
      Dim vm = ExcelRuleViewMapLoader.LoadExcelRuleViewMap()

      Dim map As New Dictionary(Of String, String)(StringComparer.OrdinalIgnoreCase)

      For Each v In vm.Views
        If v.IsConceptual Then Continue For

        ' Excel-friendly key: DisplayName with spaces removed
        Dim key As String = v.DisplayName.Replace(" ", "").Trim()

        ' Actual SQL view name
        Dim actual As String = v.Name

        If Not map.ContainsKey(key) Then
          map.Add(key, actual)
        End If
      Next

      Return map
    End Function

    Private ReadOnly ViewMap As Dictionary(Of String, String) = BuildViewMap()

    Private Function MakeUnknownViewError(input As String) As Object(,)
      Dim keys = ViewMap.Keys.ToList()
      Dim rowCount = keys.Count + 2 ' header + blank + names

      Dim result(rowCount - 1, 0) As Object

      result(0, 0) = "#UNKNOWN_VIEW: " & input
      result(1, 0) = "Valid view names:"

      For i = 0 To keys.Count - 1
        result(i + 2, 0) = keys(i)
      Next

      Return result
    End Function

    <ExcelFunction(
      Name:="ResourceManagerViewValues",
    Description:="Returns one of the following views as a table. " & GeneratedViewMetadata.Views.ValidViewNames,
    Category:="Resource Management"
    )>
    Public Function ResourceManagerViewValues(
      <ExcelArgument("The name of the view to return.")>
      viewName As String,
      <ExcelArgument("TRUE to include column headers, FALSE to omit them.")>
      includeHeaders As Boolean
    ) As Object(,)

      Try
        If String.IsNullOrWhiteSpace(viewName) Then
          Return MakeError("#INVALID_VIEW_NAME")
        End If

        ' Normalise: remove spaces
        Dim key = viewName.Replace(" ", "").Trim()

        If Not ViewMap.ContainsKey(key) Then
          Return MakeUnknownViewError(viewName)
        End If

        Dim actualView = ViewMap(key)
        Dim result = ResourceManagementViews.GetViewAsArray(actualView, includeHeaders)

        Return result

      Catch ex As Exception
        Return MakeError("#VIEW_ERROR: " & ex.Message)
      End Try

    End Function

    Private Function GetValidViewNamesDescription() As String
      Dim vm = ExcelRuleViewMapLoader.LoadExcelRuleViewMap()

      Dim names = vm.Views _
            .Where(Function(v) Not v.IsConceptual) _
            .Select(Function(v) v.DisplayName) _
            .ToList()

      Return "Valid names: " & String.Join(", ", names)
    End Function


    '    Private ReadOnly ViewMap As New Dictionary(Of String, String)(StringComparer.OrdinalIgnoreCase) From {
    '    {"resource", "vwDim_Resource"},
    '    {"customfield", "vwDim_CustomField"},
    '    {"customfieldvalue", "vwDim_CustomFieldValue"},
    '    {"resourcecustomfield", "vwFact_ResourceCustomField"},
    '    {"resourceavailabilitypattern", "vwFact_ResourceAvailabilityPattern"},
    '    {"closure", "vwDim_Closure"},
    '    {"list", "vwDim_List"}
    '}

    '<ExcelFunction(
    '  Name:="ResourceManagerViewValues",
    '  Description:="Returns the specified database view as a table. Valid names: Resource, CustomField, CustomFieldValue, ResourceCustomField, ResourceAvailabilityPattern, Closure, List.",
    '  Category:="Resource Management"
    ')>
    'Public Function ResourceManagerViewValues(
    '  <ExcelArgument("The name of the view to return.")>
    '  viewName As String,
    '  <ExcelArgument("TRUE to include column headers, FALSE to omit them.")>
    '  includeHeaders As Boolean
    ') As Object(,)

    '  Try
    '    If String.IsNullOrWhiteSpace(viewName) Then
    '      Return MakeError("#INVALID_VIEW_NAME")
    '    End If

    '    Dim key = viewName.Trim().Replace(" ", "")
    '    If Not ViewMap.ContainsKey(key) Then
    '      Return MakeError("#UNKNOWN_VIEW: " & viewName)
    '    End If

    '    Dim actualView = ViewMap(key)
    '    Dim result = ResourceManagementViews.GetViewAsArray(actualView, includeHeaders)

    '    Return result

    '  Catch ex As Exception
    '    Return MakeError("#VIEW_ERROR: " & ex.Message)
    '  End Try

    'End Function

    Private Function MakeError(msg As String) As Object(,)
      Dim arr(0, 0) As Object
      arr(0, 0) = msg
      Return arr
    End Function

    <ExcelFunction(
      Name:="ResourceManagerAvailability",
      Description:="Returns the resource availability between two dates.",
      Category:="Resource Management"
    )>
    Public Function ResourceManagerAvailability(
      <ExcelArgument("Start date (Excel date, text, or number).")>
      StartDate As Object,
      <ExcelArgument("End date (Excel date, text, or number).")>
      EndDate As Object,
      <ExcelArgument("TRUE to include column headers, FALSE to omit them.")>
      includeHeaders As Boolean
    ) As Object(,)
      Try
        Dim raw = ResourceAvailabilityFunction.GetResourceAvailability(StartDate, EndDate)
        If includeHeaders Then
          Return raw
        Else
          Return StripHeaders(raw)
        End If
      Catch ex As Exception
        Dim err(0, 0) As Object
        err(0, 0) = "#AVAILABILITY_ERROR: " & ex.Message
        Return err
      End Try

    End Function

    Private Function StripHeaders(raw As Object(,)) As Object(,)
      Dim rowCount = raw.GetLength(0)
      Dim colCount = raw.GetLength(1)

      If rowCount <= 1 Then
        Dim empty(0, 0) As Object
        empty(0, 0) = ""
        Return empty
      End If

      Dim result(rowCount - 2, colCount - 1) As Object
      For r = 1 To rowCount - 1
        For c = 0 To colCount - 1
          result(r - 1, c) = raw(r, c)
        Next
      Next

      Return result
    End Function
  End Module

End Namespace
