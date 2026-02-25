Option Explicit On
Imports System.Globalization

' ==========================================================================================
' Routine: ResourceAvailabilityFunction
' Purpose: Computes resource availability over a date range using SQL views and pattern rules.
' Notes:
'   - Uses SQL views:
'       vwDim_Resource
'       vwFact_ResourceAvailabilityPattern
'       vwDim_Closure
'   - Returns one row per availability window per resource per date.
'   - AvailabilityID is nullable when derived from defaultMode or closures.
' ==========================================================================================
Friend Module ResourceAvailabilityFunction

  ' ==========================================================================================
  ' Routine: GetResourceAvailability
  ' Purpose: Entry point for Excel. Returns a 2D array of availability windows.
  ' Parameters:
  '   StartDate - Start of date range (string, any Excel-friendly format)
  '   EndDate   - End of date range (string, any Excel-friendly format)
  ' Returns:
  '   Object(,) - 2D array with header row and availability rows
  ' Notes:
  '   - Uses SafeParseDate (shared utility) to parse input dates.
  ' ==========================================================================================
  Friend Function GetResourceAvailability(StartDate As String, EndDate As String) As Object(,)

    Dim startDt As Date = SafeParseDate(StartDate)
    Dim endDt As Date = SafeParseDate(EndDate)

    If endDt < startDt Then
      Dim empty(0, 0) As Object
      empty(0, 0) = "No dates in range"
      Return empty
    End If

    Dim resources = ResourceManagementViews.GetViewAsArray("vwDim_Resource")
    Dim patterns = ResourceManagementViews.GetViewAsArray("vwFact_ResourceAvailabilityPattern")
    Dim closures = ResourceManagementViews.GetViewAsArray("vwDim_Closure")

    Dim defaultMode As String = LoadMetadataValue("DefaultMode")

    Dim dateList = GenerateDateList(startDt, endDt)

    Dim result As New List(Of AvailabilityWindow)

    Dim resRowCount As Integer = resources.GetLength(0)
    Dim resColCount As Integer = resources.GetLength(1)

    Dim resResourceIDCol As Integer = GetColumnIndex(resources, "ResourceID")
    Dim resPreferredNameCol As Integer = GetColumnIndex(resources, "PreferredName")
    Dim resStartDateCol As Integer = GetColumnIndex(resources, "StartDate")
    Dim resEndDateCol As Integer = GetColumnIndex(resources, "EndDate")

    For r As Integer = 1 To resRowCount - 1
      Dim resourceRow(resColCount - 1) As Object
      For c As Integer = 0 To resColCount - 1
        resourceRow(c) = resources(r, c)
      Next

      For Each d In dateList
        Dim windows = EvaluateAvailabilityForResourceOnDate(
                        resourceRow,
                        d,
                        patterns,
                        closures,
                        defaultMode,
                        resResourceIDCol,
                        resPreferredNameCol,
                        resStartDateCol,
                        resEndDateCol
                      )
        result.AddRange(windows)
      Next
    Next

    Return AvailabilityWindowsTo2DArray(result)

  End Function

  ' ==========================================================================================
  ' Routine: GenerateDateList
  ' Purpose: Generates a list of dates from startDt to endDt inclusive.
  ' Parameters:
  '   startDt - Start date
  '   endDt   - End date
  ' Returns:
  '   List(Of Date) - List of dates
  ' Notes:
  '   - Returns empty list if endDt < startDt.
  ' ==========================================================================================
  Private Function GenerateDateList(startDt As Date, endDt As Date) As List(Of Date)
    Dim list As New List(Of Date)
    If endDt < startDt Then Return list

    Dim d As Date = startDt
    While d <= endDt
      list.Add(d)
      d = d.AddDays(1)
    End While

    Return list
  End Function

  ' ==========================================================================================
  ' Routine: AvailabilityWindow (class)
  ' Purpose: Represents a single availability window for a resource on a given date.
  ' Parameters:
  '   d              - Date of the window
  '   resourceID     - Resource identifier (text)
  '   preferredName  - Preferred name for display
  '   availabilityID - Availability pattern ID (nullable)
  '   isAvailable    - True if available, False if unavailable
  '   allDay         - True if all-day window
  '   startTime      - Excel serial time (Double?) or Nothing for all-day
  '   endTime        - Excel serial time (Double?) or Nothing for all-day
  ' Returns:
  '   N/A (class)
  ' Notes:
  '   - StartTime/EndTime are in [0,1) as Excel serial fractions of a day.
  ' ==========================================================================================
  Friend Class AvailabilityWindow
    Public Property [Date] As Date
    Public Property ResourceID As String
    Public Property PreferredName As String
    Public Property AvailabilityID As String
    Public Property IsAvailable As Boolean
    Public Property AllDay As Boolean
    Public Property StartTime As Double?
    Public Property EndTime As Double?

    Public Sub New(d As Date,
                   resourceID As String,
                   preferredName As String,
                   availabilityID As String,
                   isAvailable As Boolean,
                   allDay As Boolean,
                   startTime As Double?,
                   endTime As Double?)

      Me.Date = d
      Me.ResourceID = resourceID
      Me.PreferredName = preferredName
      Me.AvailabilityID = availabilityID
      Me.IsAvailable = isAvailable
      Me.AllDay = allDay
      Me.StartTime = startTime
      Me.EndTime = endTime
    End Sub
  End Class

  ' ==========================================================================================
  ' Routine: EvaluateAvailabilityForResourceOnDate
  ' Purpose: Computes all availability windows for a single resource on a single date.
  ' Parameters:
  '   resource            - Resource row from vwDim_Resource
  '   d                   - Date to evaluate
  '   patterns            - 2D array from vwFact_ResourceAvailabilityPattern
  '   closures            - 2D array from vwDim_Closure
  '   defaultMode         - Default mode ("Available" or "Unavailable")
  '   resResourceIDCol    - Column index for ResourceID
  '   resPreferredNameCol - Column index for PreferredName
  '   resStartDateCol     - Column index for StartDate
  '   resEndDateCol       - Column index for EndDate
  ' Returns:
  '   List(Of AvailabilityWindow) - Windows for this resource/date
  ' Notes:
  '   - Applies resource active range, closures, patterns, and defaultMode.
  ' ==========================================================================================
  Private Function EvaluateAvailabilityForResourceOnDate(
          resource As Object(),
          d As Date,
          patterns As Object(,),
          closures As Object(,),
          defaultMode As String,
          resResourceIDCol As Integer,
          resPreferredNameCol As Integer,
          resStartDateCol As Integer,
          resEndDateCol As Integer
      ) As List(Of AvailabilityWindow)

    Dim windows As New List(Of AvailabilityWindow)

    Dim resourceID As String = CStr(resource(resResourceIDCol))
    Dim preferredName As String = CStr(resource(resPreferredNameCol))

    Dim resourceStartOpt = ParseIsoDate(resource(resStartDateCol))
    Dim resourceStart As Date = If(resourceStartOpt.HasValue, resourceStartOpt.Value, #1/1/1900#)

    Dim resourceEndOpt = ParseIsoDate(resource(resEndDateCol))

    If d < resourceStart Then Return windows
    If resourceEndOpt.HasValue AndAlso d > resourceEndOpt.Value Then Return windows

    If IsDateClosed(d, closures) Then
      windows.Add(New AvailabilityWindow(
                    d,
                    resourceID,
                    preferredName,
                    Nothing,
                    False,
                    True,
                    Nothing,
                    Nothing))
      Return windows
    End If

    Dim patternRows = GetPatternsForResourceOnDate(patterns, resourceID, d)

    If patternRows.Count = 0 Then
      Dim isAvail As Boolean = String.Equals(defaultMode, "Available", StringComparison.OrdinalIgnoreCase)
      windows.Add(New AvailabilityWindow(
                    d,
                    resourceID,
                    preferredName,
                    Nothing,
                    isAvail,
                    True,
                    Nothing,
                    Nothing))
      Return windows
    End If

    Dim rawWindows = BuildWindowsFromPatterns(d, resourceID, preferredName, patterns, patternRows)
    Dim merged = MergeWindows(rawWindows)
    Dim finalWindows = FillGapsWithDefaultMode(d, merged, resourceID, preferredName, defaultMode)

    Return finalWindows

  End Function

  ' ==========================================================================================
  ' Routine: IsDateClosed
  ' Purpose: Returns True if the given date falls within any closure range.
  ' Parameters:
  '   d        - Date to check
  '   closures - 2D array from vwDim_Closure
  ' Returns:
  '   Boolean - True if closed, False otherwise
  ' Notes:
  '   - StartDate/EndDate are stored as text "yyyy-MM-dd".
  ' ==========================================================================================
  Private Function IsDateClosed(d As Date, closures As Object(,)) As Boolean
    Dim rowCount As Integer = closures.GetLength(0)
    Dim startCol As Integer = GetColumnIndex(closures, "StartDate")
    Dim endCol As Integer = GetColumnIndex(closures, "EndDate")

    For r As Integer = 1 To rowCount - 1
      Dim startObj As Object = closures(r, startCol)
      Dim endObj As Object = closures(r, endCol)

      Dim sOpt = ParseIsoDate(startObj)
      If Not sOpt.HasValue Then Continue For

      Dim eOpt = ParseIsoDate(endObj)
      Dim s As Date = sOpt.Value
      Dim e As Date = If(eOpt.HasValue, eOpt.Value, s)

      If d >= s AndAlso d <= e Then
        Return True
      End If
    Next

    Return False
  End Function

  ' ==========================================================================================
  ' Routine: GetPatternsForResourceOnDate
  ' Purpose:
  '   Filters vwFact_ResourceAvailabilityPattern rows for a resource and date.
  '   Applies only the logic relevant to each row's PatternType (DateRange, Weekly, Monthly).
  ' Parameters:
  '   patterns   - 2D array from vwFact_ResourceAvailabilityPattern
  '   resourceID - Resource identifier (text)
  '   d          - Date to evaluate
  ' Returns:
  '   List(Of Object()) - Pattern rows that apply to this resource/date
  ' Notes:
  '   - Uses ISO date parsing for RangeStartDate and RangeEndDate
  '   - WeeklyDayOfWeek and MonthlyDayOfWeek are stored as text (e.g., "Monday")
  '   - PatternType is used to gate logic blocks
  ' ==========================================================================================
  Private Function GetPatternsForResourceOnDate(patterns As Object(,),
                                             resourceID As String,
                                             d As Date) As List(Of Object())

    Dim list As New List(Of Object())

    Dim rowCount As Integer = patterns.GetLength(0)
    Dim colCount As Integer = patterns.GetLength(1)

    ' Column indexes
    Dim resIDCol As Integer = GetColumnIndex(patterns, "ResourceID")
    Dim patternTypeCol As Integer = GetColumnIndex(patterns, "PatternType")
    Dim rangeStartCol As Integer = GetColumnIndex(patterns, "RangeStartDate")
    Dim rangeEndCol As Integer = GetColumnIndex(patterns, "RangeEndDate")
    Dim weeklyDayCol As Integer = GetColumnIndex(patterns, "WeeklyDayOfWeek")
    Dim recurWeeksCol As Integer = GetColumnIndex(patterns, "RecurWeeks")
    Dim monthlyTypeCol As Integer = GetColumnIndex(patterns, "MonthlyType")
    Dim monthlyDayOfMonthCol As Integer = GetColumnIndex(patterns, "MonthlyDayOfMonth")
    Dim monthlyOrdinalCol As Integer = GetColumnIndex(patterns, "MonthlyOrdinal")
    Dim monthlyDayOfWeekCol As Integer = GetColumnIndex(patterns, "MonthlyDayOfWeek")
    Dim recurMonthsCol As Integer = GetColumnIndex(patterns, "RecurMonths")

    For r As Integer = 1 To rowCount - 1
      Dim row(colCount - 1) As Object
      For c As Integer = 0 To colCount - 1
        row(c) = patterns(r, c)
      Next

      ' Resource match
      Dim patResID As String = CStr(row(resIDCol)).Trim().ToLower()
      Dim resID As String = resourceID.Trim().ToLower()
      If patResID <> resID Then Continue For

      ' Pattern type
      Dim patternType As String = CStr(row(patternTypeCol)).Trim().ToLower()

      ' ------------------------------------------------------------
      ' DateRange pattern logic
      ' ------------------------------------------------------------
      If patternType = "daterange" Then
        Dim rsOpt = ParseIsoDate(row(rangeStartCol))
        Dim reOpt = ParseIsoDate(row(rangeEndCol))

        If rsOpt.HasValue AndAlso d < rsOpt.Value Then Continue For
        If reOpt.HasValue AndAlso d > reOpt.Value Then Continue For

        list.Add(row)
        Continue For
      End If

      ' ------------------------------------------------------------
      ' Weekly pattern logic
      ' ------------------------------------------------------------
      If patternType = "weekly" Then
        Dim wdObj As Object = row(weeklyDayCol)
        If IsDBNull(wdObj) OrElse wdObj Is Nothing Then Continue For

        Dim patternDayName As String = CStr(wdObj).Trim().ToLower()
        Dim actualDayName As String = d.ToString("dddd", Globalization.CultureInfo.InvariantCulture).ToLower()
        If patternDayName <> actualDayName Then Continue For

        Dim rwObj As Object = row(recurWeeksCol)
        Dim rsOpt = ParseIsoDate(row(rangeStartCol))
        If Not IsDBNull(rwObj) AndAlso rwObj IsNot Nothing AndAlso rsOpt.HasValue Then
          Dim recurWeeks As Integer
          If Integer.TryParse(CStr(rwObj).Trim(), recurWeeks) AndAlso recurWeeks > 0 Then
            Dim weeksDiff As Integer = CInt((d - rsOpt.Value).TotalDays \ 7)
            If weeksDiff Mod recurWeeks <> 0 Then Continue For
          End If
        End If

        list.Add(row)
        Continue For
      End If

      ' ------------------------------------------------------------
      ' Monthly pattern logic
      ' ------------------------------------------------------------
      If patternType = "monthly" Then
        Dim mType As String = CStr(row(monthlyTypeCol)).Trim().ToLower()
        Dim rsOpt = ParseIsoDate(row(rangeStartCol))
        Dim rmObj As Object = row(recurMonthsCol)

        If rsOpt.HasValue AndAlso Not IsDBNull(rmObj) AndAlso rmObj IsNot Nothing Then
          Dim recurMonths As Integer
          If Integer.TryParse(CStr(rmObj).Trim(), recurMonths) AndAlso recurMonths > 0 Then
            Dim monthsDiff As Integer = (d.Year - rsOpt.Value.Year) * 12 + (d.Month - rsOpt.Value.Month)
            If monthsDiff Mod recurMonths <> 0 Then Continue For
          End If
        End If

        If mType = "dayofmonth" Then
          Dim domObj As Object = row(monthlyDayOfMonthCol)
          Dim dom As Integer
          If IsDBNull(domObj) OrElse domObj Is Nothing OrElse Not Integer.TryParse(CStr(domObj).Trim(), dom) Then Continue For
          If dom <> d.Day Then Continue For

        ElseIf mType = "ordinalday" OrElse mType = "ordinalweek" Then
          Dim ordObj As Object = row(monthlyOrdinalCol)
          Dim mdwObj As Object = row(monthlyDayOfWeekCol)
          If IsDBNull(ordObj) OrElse IsDBNull(mdwObj) Then Continue For

          Dim ordinal As Integer
          If Not Integer.TryParse(CStr(ordObj).Trim(), ordinal) Then Continue For

          Dim patternDayName As String = CStr(mdwObj).Trim().ToLower()
          If Not IsOrdinalMatchByName(d, ordinal, patternDayName) Then Continue For
        Else
          Continue For ' Unknown MonthlyType
        End If

        list.Add(row)
        Continue For
      End If

    Next

    Return list
  End Function

  ' ==========================================================================================
  ' Routine: IsOrdinalMatchByName
  ' Purpose: Determines if a date matches an ordinal weekday in its month by day name.
  ' Parameters:
  '   d              - Date to test
  '   ordinal        - 1..4 for nth, 5 for last
  '   patternDayName - Day name ("monday", "friday", etc.)
  ' Returns:
  '   Boolean - True if d matches the ordinal weekday, False otherwise
  ' Notes:
  '   - Uses invariant culture day names.
  ' ==========================================================================================
  Private Function IsOrdinalMatchByName(d As Date, ordinal As Integer, patternDayName As String) As Boolean
    Dim firstOfMonth As New Date(d.Year, d.Month, 1)
    Dim culture = CultureInfo.InvariantCulture

    Dim current As Date = firstOfMonth
    Dim matches As New List(Of Date)

    While current.Month = d.Month
      If current.ToString("dddd", culture).ToLower() = patternDayName Then
        matches.Add(current)
      End If
      current = current.AddDays(1)
    End While

    If ordinal > 0 AndAlso ordinal <= matches.Count Then
      Return d.Date = matches(ordinal - 1).Date
    End If

    If ordinal = 5 AndAlso matches.Count > 0 Then
      Return d.Date = matches(matches.Count - 1).Date
    End If

    Return False
  End Function

  ' ==========================================================================================
  ' Routine: BuildWindowsFromPatterns
  ' Purpose: Converts pattern rows into raw availability windows (per date).
  ' Parameters:
  '   d            - Date being evaluated
  '   resourceID   - Resource identifier
  '   preferredName- Preferred name
  '   patterns     - 2D array from vwFact_ResourceAvailabilityPattern
  '   patternRows  - Filtered pattern rows for this resource/date
  ' Returns:
  '   List(Of AvailabilityWindow) - Raw windows from patterns
  ' Notes:
  '   - Mode is 'Available' or 'Unavailable'.
  '   - AllDay is 0/1.
  '   - StartTime/EndTime are "hh\:mm" text or NULL.
  ' ==========================================================================================
  Private Function BuildWindowsFromPatterns(d As Date,
                                           resourceID As String,
                                           preferredName As String,
                                           patterns As Object(,),
                                           patternRows As List(Of Object())) As List(Of AvailabilityWindow)

    Dim windows As New List(Of AvailabilityWindow)
    If patternRows.Count = 0 Then Return windows

    Dim availabilityIDCol As Integer = GetColumnIndex(patterns, "AvailabilityID")
    Dim modeCol As Integer = GetColumnIndex(patterns, "Mode")
    Dim allDayCol As Integer = GetColumnIndex(patterns, "AllDay")
    Dim startTimeCol As Integer = GetColumnIndex(patterns, "StartTime")
    Dim endTimeCol As Integer = GetColumnIndex(patterns, "EndTime")

    For Each row In patternRows
      Dim availabilityID As String =
        If(IsDBNull(row(availabilityIDCol)) OrElse row(availabilityIDCol) Is Nothing,
           Nothing,
           CStr(row(availabilityIDCol)))

      Dim mode As String = If(IsDBNull(row(modeCol)) OrElse row(modeCol) Is Nothing,
                              "",
                              CStr(row(modeCol)).Trim())

      Dim allDay As Boolean = False
      If Not IsDBNull(row(allDayCol)) AndAlso row(allDayCol) IsNot Nothing Then
        Dim n As Integer
        If Integer.TryParse(CStr(row(allDayCol)).Trim(), n) Then
          allDay = (n <> 0)
        End If
      End If

      Dim startSerial As Double? = Nothing
      Dim endSerial As Double? = Nothing

      If Not allDay Then
        Dim st As Double = ParseTimeToSerial(row(startTimeCol))
        Dim et As Double = ParseTimeToSerial(row(endTimeCol))
        startSerial = st
        endSerial = et
      End If

      Dim isAvailable As Boolean = String.Equals(mode, "Available", StringComparison.OrdinalIgnoreCase)

      windows.Add(New AvailabilityWindow(
                    d,
                    resourceID,
                    preferredName,
                    availabilityID,
                    isAvailable,
                    allDay,
                    startSerial,
                    endSerial))
    Next

    Return windows
  End Function

  ' ==========================================================================================
  ' Routine: MergeWindows
  ' Purpose: Merges overlapping windows and resolves conflicts (Unavailable wins).
  ' Parameters:
  '   raw - Raw windows from patterns
  ' Returns:
  '   List(Of AvailabilityWindow) - Non-overlapping windows over [0,1)
  ' Notes:
  '   - AllDay windows are treated as [0,1).
  ' ==========================================================================================
  Private Function MergeWindows(raw As List(Of AvailabilityWindow)) As List(Of AvailabilityWindow)
    Dim result As New List(Of AvailabilityWindow)
    If raw.Count = 0 Then Return result

    Dim boundaries As New SortedSet(Of Double)
    boundaries.Add(0.0R)
    boundaries.Add(1.0R)

    For Each w In raw
      If w.AllDay Then
        boundaries.Add(0.0R)
        boundaries.Add(1.0R)
      Else
        If w.StartTime.HasValue Then boundaries.Add(w.StartTime.Value)
        If w.EndTime.HasValue Then boundaries.Add(w.EndTime.Value)
      End If
    Next

    Dim bList As List(Of Double) = boundaries.ToList()
    bList.Sort()

    Dim sample As AvailabilityWindow = raw(0)
    Dim d As Date = sample.Date
    Dim resourceID As String = sample.ResourceID
    Dim preferredName As String = sample.PreferredName

    For i As Integer = 0 To bList.Count - 2
      Dim segStartTime As Double = bList(i)
      Dim segEndTime As Double = bList(i + 1)
      If segEndTime <= segStartTime Then Continue For

      Dim hasAny As Boolean = False
      Dim anyUnavailable As Boolean = False
      Dim anyAvailable As Boolean = False
      Dim chosenAvailabilityID As String = Nothing

      For Each w In raw
        Dim wStartTime As Double = 0.0R
        Dim wEndTime As Double = 1.0R

        If Not w.AllDay Then
          wStartTime = If(w.StartTime.HasValue, w.StartTime.Value, 0.0R)
          wEndTime = If(w.EndTime.HasValue, w.EndTime.Value, 1.0R)
        End If

        If segStartTime >= wEndTime OrElse segEndTime <= wStartTime Then
          Continue For
        End If

        hasAny = True
        If Not w.IsAvailable Then
          anyUnavailable = True
          chosenAvailabilityID = w.AvailabilityID
        Else
          anyAvailable = True
          If chosenAvailabilityID Is Nothing Then chosenAvailabilityID = w.AvailabilityID
        End If
      Next

      If Not hasAny Then Continue For

      Dim segIsAvailable As Boolean = Not anyUnavailable
      Dim allDay As Boolean = (segStartTime <= 0.0R AndAlso segEndTime >= 1.0R)
      Dim startSerial As Double? = If(allDay, Nothing, segStartTime)
      Dim endSerial As Double? = If(allDay, Nothing, segEndTime)

      result.Add(New AvailabilityWindow(
                   d,
                   resourceID,
                   preferredName,
                   chosenAvailabilityID,
                   segIsAvailable,
                   allDay,
                   startSerial,
                   endSerial))
    Next

    Return result
  End Function

  ' ==========================================================================================
  ' Routine: FillGapsWithDefaultMode
  ' Purpose: Fills gaps in the day [0,1) with defaultMode windows.
  ' Parameters:
  '   d            - Date being evaluated
  '   merged       - Merged windows from patterns
  '   resourceID   - Resource identifier
  '   preferredName- Preferred name
  '   defaultMode  - Default mode ("Available" or "Unavailable")
  ' Returns:
  '   List(Of AvailabilityWindow) - Final windows including gaps
  ' Notes:
  '   - Gaps get AvailabilityID = Nothing.
  ' ==========================================================================================
  Private Function FillGapsWithDefaultMode(d As Date,
                                          merged As List(Of AvailabilityWindow),
                                          resourceID As String,
                                          preferredName As String,
                                          defaultMode As String) As List(Of AvailabilityWindow)

    Dim result As New List(Of AvailabilityWindow)
    If merged.Count = 0 Then Return result

    Dim boundaries As New SortedSet(Of Double)
    boundaries.Add(0.0R)
    boundaries.Add(1.0R)

    For Each w In merged
      If w.AllDay Then
        boundaries.Add(0.0R)
        boundaries.Add(1.0R)
      Else
        If w.StartTime.HasValue Then boundaries.Add(w.StartTime.Value)
        If w.EndTime.HasValue Then boundaries.Add(w.EndTime.Value)
      End If
    Next

    Dim bList As List(Of Double) = boundaries.ToList()
    bList.Sort()

    Dim isDefaultAvailable As Boolean = String.Equals(defaultMode, "Available", StringComparison.OrdinalIgnoreCase)

    For i As Integer = 0 To bList.Count - 2
      Dim segStartTime As Double = bList(i)
      Dim segEndTime As Double = bList(i + 1)
      If segEndTime <= segStartTime Then Continue For

      Dim covered As Boolean = False
      Dim segIsAvailable As Boolean = False
      Dim segAvailabilityID As String = Nothing
      Dim segAllDay As Boolean = False

      For Each w In merged
        Dim wStartTime As Double = 0.0R
        Dim wEndTime As Double = 1.0R
        If Not w.AllDay Then
          wStartTime = If(w.StartTime.HasValue, w.StartTime.Value, 0.0R)
          wEndTime = If(w.EndTime.HasValue, w.EndTime.Value, 1.0R)
        End If

        If segStartTime >= wEndTime OrElse segEndTime <= wStartTime Then
          Continue For
        End If

        covered = True
        segIsAvailable = w.IsAvailable
        segAvailabilityID = w.AvailabilityID
        segAllDay = (segStartTime <= 0.0R AndAlso segEndTime >= 1.0R)
        Exit For
      Next

      If covered Then
        Dim startSerial As Double? = If(segAllDay, Nothing, segStartTime)
        Dim endSerial As Double? = If(segAllDay, Nothing, segEndTime)
        result.Add(New AvailabilityWindow(
                     d,
                     resourceID,
                     preferredName,
                     segAvailabilityID,
                     segIsAvailable,
                     segAllDay,
                     startSerial,
                     endSerial))
      Else
        Dim gapAllDay As Boolean = (segStartTime <= 0.0R AndAlso segEndTime >= 1.0R)
        Dim startSerial As Double? = If(gapAllDay, Nothing, segStartTime)
        Dim endSerial As Double? = If(gapAllDay, Nothing, segEndTime)
        result.Add(New AvailabilityWindow(
                     d,
                     resourceID,
                     preferredName,
                     Nothing,
                     isDefaultAvailable,
                     gapAllDay,
                     startSerial,
                     endSerial))
      End If
    Next

    Return result
  End Function

  ' ==========================================================================================
  ' Routine: ParseIsoDate
  ' Purpose: Parses a value as ISO date "yyyy-MM-dd".
  ' Parameters:
  '   obj - Value from SQLite view (text or NULL)
  ' Returns:
  '   Nullable(Of Date) - Parsed date or Nothing
  ' Notes:
  '   - Uses invariant culture and exact format.
  ' ==========================================================================================
  Private Function ParseIsoDate(obj As Object) As Date?
    If IsDBNull(obj) OrElse obj Is Nothing Then Return Nothing
    Dim s As String = CStr(obj).Trim()
    If s = "" Then Return Nothing

    Dim d As Date
    If Date.TryParseExact(s, "yyyy-MM-dd",
                          CultureInfo.InvariantCulture,
                          DateTimeStyles.None,
                          d) Then
      Return d
    End If

    Return Nothing
  End Function

  ' ==========================================================================================
  ' Routine: ParseTimeToSerial
  ' Purpose: Converts a time value from SQLite to Excel serial time (Double).
  ' Parameters:
  '   value - Value from SQLite view (text "hh\:mm" or NULL)
  ' Returns:
  '   Double - Fraction of day in [0,1)
  ' Notes:
  '   - Uses exact "hh\:mm" format and invariant culture.
  ' ==========================================================================================
  Private Function ParseTimeToSerial(value As Object) As Double
    If IsDBNull(value) OrElse value Is Nothing Then Return 0.0R

    Dim s As String = CStr(value).Trim()
    If s = "" Then Return 0.0R

    Dim t As TimeSpan
    If TimeSpan.TryParseExact(s, "hh\:mm",
                              CultureInfo.InvariantCulture,
                              t) Then
      Return t.TotalDays
    End If

    Return 0.0R
  End Function

  ' ==========================================================================================
  ' Routine: AvailabilityWindowsTo2DArray
  ' Purpose: Converts a list of AvailabilityWindow to a 2D array for Excel.
  ' Parameters:
  '   windows - List of AvailabilityWindow
  ' Returns:
  '   Object(,) - 2D array with header row
  ' Notes:
  '   - Returns "No availability rows" if list is empty.
  ' ==========================================================================================
  Private Function AvailabilityWindowsTo2DArray(windows As List(Of AvailabilityWindow)) As Object(,)
    If windows Is Nothing OrElse windows.Count = 0 Then
      Dim empty(0, 0) As Object
      empty(0, 0) = "No availability rows"
      Return empty
    End If

    Dim rowCount As Integer = windows.Count
    Dim colCount As Integer = 8

    Dim result(rowCount, colCount - 1) As Object

    result(0, 0) = "Date"
    result(0, 1) = "ResourceID"
    result(0, 2) = "PreferredName"
    result(0, 3) = "AvailabilityID"
    result(0, 4) = "IsAvailable"
    result(0, 5) = "AllDay"
    result(0, 6) = "StartTime"
    result(0, 7) = "EndTime"

    Dim r As Integer = 1
    For Each w In windows
      result(r, 0) = w.Date
      result(r, 1) = w.ResourceID
      result(r, 2) = w.PreferredName
      result(r, 3) = w.AvailabilityID
      result(r, 4) = w.IsAvailable
      result(r, 5) = w.AllDay
      result(r, 6) = w.StartTime
      result(r, 7) = w.EndTime
      r += 1
    Next

    Return result
  End Function

  ' ==========================================================================================
  ' Routine: GetColumnIndex
  ' Purpose: Returns the column index for a given column name in a 2D array.
  ' Parameters:
  '   data      - 2D array with header row at index 0
  '   colName   - Column name to find
  ' Returns:
  '   Integer - Column index (0-based)
  ' Notes:
  '   - Throws if column not found.
  ' ==========================================================================================
  Private Function GetColumnIndex(data As Object(,), colName As String) As Integer
    Dim colCount As Integer = data.GetLength(1)
    For c As Integer = 0 To colCount - 1
      If String.Equals(CStr(data(0, c)), colName, StringComparison.OrdinalIgnoreCase) Then
        Return c
      End If
    Next
    Throw New ArgumentException("Column not found: " & colName)
  End Function

End Module