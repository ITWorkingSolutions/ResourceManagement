Option Strict On
Imports System.Collections.Generic

' ==========================================================================================
' Module: UILoaderSaverResourceManager
' Purpose:
'   Loads the UIModelResourceManager for the ResourceManager form.
'   - First call: initialise model and load all resource summaries.
'   - Subsequent calls: refresh availability summaries for the selected resource.
'   - Uses generic DAL loaders (LoadRecords, LoadRecordsByFields, LoadRecord).
'   - Does NOT load roles (not displayed in ResourceManager).
'   - Does NOT write availability.
' ==========================================================================================
Friend Module UILoaderSaverResourceManager

  ' ==========================================================================================
  ' Routine: LoadResourceManagerModel
  ' Purpose:
  '   Primary loader for the ResourceManager UI model.
  '   - Path 1: If model Is Nothing, initialise and load resource summaries.
  '   - Path 2: If model exists, refresh availability for SelectedResourceID.
  ' Parameters:
  '   model - UIModelResourceManager passed ByRef.
  ' Returns:
  '   (None)
  ' Notes:
  '   Uses internal loaders LoadResourceSummaries and LoadAvailabilitySummaries.
  ' ==========================================================================================
  Friend Sub LoadResourceManagerModel(ByRef model As UIModelResourceManager)

    '=== Path 1: First call – initialise model and load resource list ==
    If model Is Nothing Then
      model = New UIModelResourceManager()

      model.ResourceSummaries = LoadResourceSummaries()

      model.SelectedResourceID = ""
      model.AvailabilitySummaries.Clear()

      Exit Sub
    End If

    '=== Path 2: Subsequent calls – refresh availability for current selection ==
    If String.IsNullOrEmpty(model.SelectedResourceID) Then
      model.AvailabilitySummaries.Clear()
    Else
      model.AvailabilitySummaries = LoadAvailabilitySummaries(model.SelectedResourceID)
    End If

  End Sub

  ' ==========================================================================================
  ' Routine: LoadResourceSummaries
  ' Purpose:
  '   Loads all resources and converts them into UIResourceManagerResourceSummaryRow.
  ' Parameters:
  '   (None)
  ' Returns:
  '   List(Of UIResourceManagerResourceSummaryRow)
  ' Notes:
  '   Uses LoadRecords(Of RecordResource) to fetch all rows.
  ' ==========================================================================================
  Private Function LoadResourceSummaries() As List(Of UIResourceManagerResourceSummaryRow)
    Dim result As New List(Of UIResourceManagerResourceSummaryRow)()
    Dim conn As SQLiteConnectionWrapper = Nothing

    Try
      conn = OpenDatabase(AddInContext.Current.Config.ActiveDbPath)

      Dim records As List(Of RecordResource) = LoadRecords(Of RecordResource)(conn)
      Dim today As Date = Date.Today

      For Each rec As RecordResource In records
        Dim row As New UIResourceManagerResourceSummaryRow()

        row.ResourceID = rec.ResourceID
        row.PreferredName = rec.PreferredName
        row.FullName = rec.FirstName & " " & rec.LastName
        'row.EmployeeID = rec.EmployeeID

        If Not String.IsNullOrEmpty(rec.EndDate) Then
          row.IsInactive = (CDate(rec.EndDate) < today)
        Else
          row.IsInactive = False
        End If

        result.Add(row)
      Next rec

    Finally
      If conn IsNot Nothing Then conn.Close()
    End Try

    Return result
  End Function

  ' ==========================================================================================
  ' Routine: LoadAvailabilitySummaries
  ' Purpose:
  '   Loads availability summaries for a specific resource.
  ' Parameters:
  '   resourceID - String identifying the resource.
  ' Returns:
  '   List(Of UIResourceManagerAvailabilitySummaryRow)
  ' Notes:
  '   Loads RecordResourceAvailability rows using LoadRecordsByFields.
  '   Loads pattern tables using LoadRecord or LoadRecordsByFields.
  ' ==========================================================================================
  Private Function LoadAvailabilitySummaries(ByVal resourceID As String) As List(Of UIResourceManagerAvailabilitySummaryRow)
    Dim result As New List(Of UIResourceManagerAvailabilitySummaryRow)()
    Dim conn As SQLiteConnectionWrapper = Nothing

    Try
      conn = OpenDatabase(AddInContext.Current.Config.ActiveDbPath)

      Dim fieldNames() As String = {"ResourceID"}
      Dim fieldValues() As Object = {resourceID}

      Dim availList As List(Of RecordResourceAvailability) =
        LoadRecordsByFields(Of RecordResourceAvailability)(conn, fieldNames, fieldValues)

      For Each avail As RecordResourceAvailability In availList
        Dim row As New UIResourceManagerAvailabilitySummaryRow()

        row.AvailabilityID = avail.AvailabilityID
        row.Description = BuildAvailabilityDescription(conn, avail)

        result.Add(row)
      Next avail

    Finally
      If conn IsNot Nothing Then conn.Close()
    End Try

    Return result
  End Function

  ' ==========================================================================================
  ' Routine: BuildAvailabilityDescription
  ' Purpose:
  '   Builds the human-readable description for an availability row.
  ' Parameters:
  '   conn - SQLiteConnectionWrapper
  '   avail - RecordResourceAvailability
  ' Returns:
  '   String
  ' Notes:
  '   Combines: Mode, Time window, PatternType, Pattern details, Range wrapper.
  ' ==========================================================================================
  Private Function BuildAvailabilityDescription(
    ByVal conn As SQLiteConnectionWrapper,
    ByVal avail As RecordResourceAvailability) As String

    Dim modeText As String = If(avail.Mode = "Unavailable", "Unavailable", "Available")

    Dim timeText As String
    If (avail.AllDay = 1) Then
      timeText = "All day"
    Else
      timeText = avail.StartTime & "–" & avail.EndTime
    End If

    Dim patternText As String = ""

    Select Case avail.PatternType
      Case "DateRange"
        patternText = BuildDesc_DateRange(conn, avail.AvailabilityID)

      Case "Weekly"
        patternText = BuildDesc_Weekly(conn, avail.AvailabilityID)

      Case "Monthly"
        patternText = BuildDesc_Monthly(conn, avail.AvailabilityID)
    End Select

    Dim rangeText As String = BuildDesc_Range(conn, avail.AvailabilityID)

    Dim parts As New List(Of String)
    parts.Add(modeText)
    parts.Add(timeText)
    If patternText <> "" Then parts.Add(patternText)
    If rangeText <> "" Then parts.Add(rangeText)

    Return String.Join("; ", parts)
  End Function

  ' ==========================================================================================
  ' Routine: BuildDesc_DateRange
  ' Purpose:
  '   Builds description for DateRange pattern.
  ' Parameters:
  '   conn - SQLiteConnectionWrapper
  '   availabilityID - String
  ' Returns:
  '   String
  ' Notes:
  '   Loads RecordPatternDateRange using LoadRecord.
  ' ==========================================================================================
  Private Function BuildDesc_DateRange(
    ByVal conn As SQLiteConnectionWrapper,
    ByVal availabilityID As String) As String

    Dim pk() As Object = {availabilityID}
    Dim p As RecordPatternDateRange = LoadRecord(Of RecordPatternDateRange)(conn, pk)

    If p Is Nothing Then Return ""

    Return "Date range " & p.StartDate & " to " & p.EndDate
  End Function

  ' ==========================================================================================
  ' Routine: BuildDesc_Weekly
  ' Purpose:
  '   Builds description for Weekly pattern.
  ' Parameters:
  '   conn - SQLiteConnectionWrapper
  '   availabilityID - String
  ' Returns:
  '   String
  ' Notes:
  '   Loads RecordPatternWeekly and RecordPatternWeeklyDays.
  ' ==========================================================================================
  Private Function BuildDesc_Weekly(
    ByVal conn As SQLiteConnectionWrapper,
    ByVal availabilityID As String) As String

    Dim pk() As Object = {availabilityID}
    Dim p As RecordPatternWeekly = LoadRecord(Of RecordPatternWeekly)(conn, pk)

    If p Is Nothing Then Return ""

    Dim fieldNames() As String = {"AvailabilityID"}
    Dim fieldValues() As Object = {availabilityID}
    Dim days As List(Of RecordPatternWeeklyDays) =
      LoadRecordsByFields(Of RecordPatternWeeklyDays)(conn, fieldNames, fieldValues)

    Dim dayList As String = ""
    For Each d As RecordPatternWeeklyDays In days
      If dayList <> "" Then dayList &= ", "
      dayList &= d.DayOfWeek
    Next d

    If dayList = "" Then
      Return "Every " & p.RecurWeeks & " week(s)"
    Else
      Return "Every " & p.RecurWeeks & " week(s) on " & dayList
    End If

  End Function

  ' ==========================================================================================
  ' Routine: BuildDesc_Monthly
  ' Purpose:
  '   Builds description for Monthly pattern.
  ' Parameters:
  '   conn - SQLiteConnectionWrapper
  '   availabilityID - String
  ' Returns:
  '   String
  ' Notes:
  '   Loads RecordPatternMonthly using LoadRecord.
  ' ==========================================================================================
  Private Function BuildDesc_Monthly(
    ByVal conn As SQLiteConnectionWrapper,
    ByVal availabilityID As String) As String

    Dim pk() As Object = {availabilityID}
    Dim p As RecordPatternMonthly = LoadRecord(Of RecordPatternMonthly)(conn, pk)

    If p Is Nothing Then Return ""

    Select Case p.MonthlyType

      Case "DayOfMonth"
        Return "Day " & p.DayOfMonth & " of every " & p.RecurMonths & " month(s)"

      Case "OrdinalDay"
        Return p.Ordinal & " " & p.DayOfWeek & " of every " & p.RecurMonths & " month(s)"

      Case "OrdinalWeek"
        Return p.Ordinal & " week of every " & p.RecurMonths & " month(s)"

      Case Else
        Return ""
    End Select

  End Function

  ' ==========================================================================================
  ' Routine: BuildDesc_Range
  ' Purpose:
  '   Builds description for Range wrapper (applies to all patterns).
  ' Parameters:
  '   conn - SQLiteConnectionWrapper
  '   availabilityID - String
  ' Returns:
  '   String
  ' Notes:
  '   Loads RecordPatternRange using LoadRecord.
  ' ==========================================================================================
  Private Function BuildDesc_Range(
    ByVal conn As SQLiteConnectionWrapper,
    ByVal availabilityID As String) As String

    Dim pk() As Object = {availabilityID}
    Dim p As RecordPatternRange = LoadRecord(Of RecordPatternRange)(conn, pk)

    If p Is Nothing Then Return ""

    Select Case p.EndType

      Case "NoEnd"
        Return "From " & p.StartDate & ", no end"

      Case "EndDate"
        Return "From " & p.StartDate & " to " & p.EndDate

      Case "EndAfterOccurrences"
        Return "From " & p.StartDate & ", end after " & p.EndAfterOccurrences & " occurrence(s)"

      Case Else
        Return ""
    End Select

  End Function

End Module