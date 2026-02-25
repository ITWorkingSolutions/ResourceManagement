Option Explicit On

' ==========================================================================================
' Module: UILoaderSaverAvailability
' Purpose:
'   Load and save a single Availability record (and all pattern tables) using the
'   UIModelAvailability contract.
'
'   - NO direct SQL. All database interaction is via:
'       * OpenDatabase
'       * RecordLoader
'       * RecordSaver
'
'   - Supports:
'       * Add (no AvailabilityID passed)
'       * Update (existing AvailabilityID)
'       * Delete (PendingAction = Delete)
'
'   - Pattern tables:
'       * tblResourceAvailability
'       * tblPatternRange
'       * tblPatternDateRange
'       * tblPatternWeekly
'       * tblPatternWeeklyDays
'       * tblPatternMonthly
'
'   - Loader/Saver does not infer Add/Update/Delete.
'     Caller sets PendingAction explicitly.
' ==========================================================================================
Friend Module UILoaderSaverAvailability

  ' ==========================================================================================
  ' Routine: LoadAvailabilityModel
  ' Purpose:
  '   Load an existing availability (and all pattern tables), OR return an initialised empty
  '   model when no AvailabilityID is supplied.
  '
  ' Parameters:
  '   availabilityID - "" means return an initialised empty model.
  '
  ' Returns:
  '   UIModelAvailability containing:
  '       - Availability (snapshot)
  '       - ActionAvailability (working copy)
  '       - PendingAction
  '
  ' Notes:
  '   - NO PendingAction set here.
  '   - Loader/Saver does not infer Add/Update/Delete.
  '   - Cloning is done here so UI has a safe working copy.
  ' ==========================================================================================
  Friend Function LoadAvailabilityModel(ByVal availabilityID As String) As UIModelAvailability

    Dim conn As SQLiteConnectionWrapper = Nothing
    Dim model As New UIModelAvailability()

    Try
      conn = OpenDatabase(AddInContext.Current.Config.ActiveDbPath)

      If String.IsNullOrEmpty(availabilityID) Then
        ' --------------------------------------------------------------
        '  No AvailabilityID: return an initialised empty model
        ' --------------------------------------------------------------
        model.Availability = New UIAvailabilityRow()
        model.ActionAvailability = CloneAvailabilityRow(model.Availability)
        model.PendingAction = AvailabilityAction.Add
        Return model
      End If

      ' --------------------------------------------------------------
      '  Load snapshot from DB
      ' --------------------------------------------------------------
      model.Availability = LoadAvailabilitySnapshot(conn, availabilityID)

      If model.Availability Is Nothing Then
        Throw New InvalidOperationException($"AvailabilityID '{availabilityID}' not found.")
      End If

      ' Working copy for UI edits
      model.ActionAvailability = CloneAvailabilityRow(model.Availability)
      model.PendingAction = AvailabilityAction.Update

      Return model

    Catch ex As Exception
      ErrorHandler.UnHandleError(ex)
      Throw
    Finally
      If conn IsNot Nothing Then conn.Close()
    End Try

  End Function


  ' ==========================================================================================
  ' Routine: SavePendingAvailabilityAction
  ' Purpose:
  '   Perform the pending availability-level action encoded in the UI model (Add / Update /
  '   Delete) and commit all pattern tables.
  '
  ' Parameters:
  '   model - UIModelAvailability containing PendingAction and ActionAvailability.
  '
  ' Returns:
  '   None
  '
  ' Notes:
  '   - Does NOT reload or replace the model.
  '   - Caller is responsible for any post-save refresh.
  ' ==========================================================================================
  Friend Sub SavePendingAvailabilityAction(ByRef model As UIModelAvailability)

    Dim conn As SQLiteConnectionWrapper = Nothing
    Dim a As UIAvailabilityRow = model.ActionAvailability
    Dim rec As RecordResourceAvailability
    Dim availabilityID As String

    Try
      conn = OpenDatabase(AddInContext.Current.Config.ActiveDbPath)

      Select Case model.PendingAction

        Case AvailabilityAction.Add
          ' === Insert new availability ===
          rec = MapNewValuesToRecord(a)
          RecordSaver.SaveRecord(conn, rec)
          availabilityID = rec.AvailabilityID

          ' Keep new ID on working copy
          model.ActionAvailability.AvailabilityID = availabilityID

          ' Save pattern tables
          SavePatterns(conn, model.ActionAvailability)

        Case AvailabilityAction.Update
          ' === Update existing availability ===
          If String.IsNullOrEmpty(a.AvailabilityID) Then
            Throw New InvalidOperationException("No availability selected for update.")
          End If

          rec = MapUpdateValuesToRecord(a)
          RecordSaver.SaveRecord(conn, rec)
          availabilityID = rec.AvailabilityID

          SavePatterns(conn, model.ActionAvailability)

        Case AvailabilityAction.Delete
          ' === Soft-delete availability and all pattern tables ===
          If String.IsNullOrEmpty(a.AvailabilityID) Then
            Throw New InvalidOperationException("No availability selected for delete.")
          End If

          availabilityID = a.AvailabilityID

          DeleteAllPatterns(conn, availabilityID)

          rec = New RecordResourceAvailability()
          rec.AvailabilityID = availabilityID
          rec.IsDeleted = True
          RecordSaver.SaveRecord(conn, rec)

        Case Else
          ' === No-op ===
          Return

      End Select

    Catch ex As Exception
      ErrorHandler.UnHandleError(ex)
      Throw
    Finally
      If conn IsNot Nothing Then conn.Close()
    End Try

  End Sub


  ' ==========================================================================================
  ' Routine: LoadAvailabilitySnapshot
  ' Purpose:
  '   Load the snapshot UIAvailabilityRow from all pattern tables.
  '
  ' Parameters:
  '   conn          - SQLite connection wrapper.
  '   availabilityID - AvailabilityID to load.
  '
  ' Returns:
  '   UIAvailabilityRow snapshot.
  '
  ' Notes:
  '   - Loads base availability + all pattern tables.
  '   - Does NOT create working copy.
  ' ==========================================================================================
  Private Function LoadAvailabilitySnapshot(ByVal conn As SQLiteConnectionWrapper,
                                            ByVal availabilityID As String) _
                                            As UIAvailabilityRow

    Dim baseRec As RecordResourceAvailability
    Dim pk() As Object = {availabilityID}

    baseRec = RecordLoader.LoadRecord(Of RecordResourceAvailability)(conn, pk)
    If baseRec Is Nothing OrElse baseRec.IsDeleted Then Return Nothing

    Dim row As New UIAvailabilityRow()

    ' Base availability
    row.AvailabilityID = baseRec.AvailabilityID
    row.ResourceID = baseRec.ResourceID
    row.Mode = baseRec.Mode
    row.PatternType = baseRec.PatternType
    row.AllDay = (baseRec.AllDay = 1)
    row.StartTime = ParseTime(baseRec.StartTime)
    row.EndTime = ParseTime(baseRec.EndTime)

    ' Pattern tables
    LoadPatternRange(conn, row)
    LoadPatternDateRange(conn, row)
    LoadPatternWeekly(conn, row)
    LoadPatternWeeklyDays(conn, row)
    LoadPatternMonthly(conn, row)

    Return row

  End Function

  ' ==========================================================================================
  ' Routine: SavePatterns
  ' Purpose:
  '   Save pattern tables for the given ActionAvailability based on PatternType.
  '
  ' Parameters:
  '   conn - SQLite connection wrapper.
  '   a    - UIAvailabilityRow (working copy).
  '
  ' Returns:
  '   None
  '
  ' Notes:
  '   - RecordPatternRange is ALWAYS saved.
  '   - Other tables are saved or deleted according to PatternType:
  '       * DateRange -> tblPatternDateRange only
  '       * Weekly    -> tblPatternWeekly + tblPatternWeeklyDays
  '       * Monthly   -> tblPatternMonthly
  ' ==========================================================================================
  Private Sub SavePatterns(ByVal conn As SQLiteConnectionWrapper,
                         ByVal a As UIAvailabilityRow)

    ' Range applies to all patterns
    SavePatternRange(conn, a)

    Select Case a.PatternType

      Case "DateRange"
        ' Keep DateRange; delete others
        SavePatternDateRange(conn, a)
        DeleteWeeklyPatterns(conn, a.AvailabilityID)
        DeleteMonthlyPattern(conn, a.AvailabilityID)

      Case "Weekly"
        ' Keep Weekly; delete others
        DeleteDateRangePattern(conn, a.AvailabilityID)
        SavePatternWeekly(conn, a)
        SavePatternWeeklyDays(conn, a)
        DeleteMonthlyPattern(conn, a.AvailabilityID)

      Case "Monthly"
        ' Keep Monthly; delete others
        DeleteDateRangePattern(conn, a.AvailabilityID)
        DeleteWeeklyPatterns(conn, a.AvailabilityID)
        SavePatternMonthly(conn, a)

      Case Else
        ' No recognised pattern: delete all pattern-specific tables, keep only Range
        DeleteDateRangePattern(conn, a.AvailabilityID)
        DeleteWeeklyPatterns(conn, a.AvailabilityID)
        DeleteMonthlyPattern(conn, a.AvailabilityID)

    End Select

  End Sub

  ' ==========================================================================================
  ' Routine: DeleteAllPatterns
  ' Purpose:
  '   Delete all pattern records for a given AvailabilityID.
  '
  ' Parameters:
  '   conn          - SQLite connection wrapper.
  '   availabilityID - AvailabilityID to delete patterns for.
  '
  ' Returns:
  '   None
  '
  ' Notes:
  '   - Uses RecordLoader + RecordSaver only.
  '   - RecordSaver.SaveRecord with IsDeleted = True performs a hard delete.
  ' ==========================================================================================
  Private Sub DeleteAllPatterns(ByVal conn As SQLiteConnectionWrapper,
                              ByVal availabilityID As String)

    Dim pk() As Object = {availabilityID}

    ' Range
    Dim r1 = RecordLoader.LoadRecord(Of RecordPatternRange)(conn, pk)
    If r1 IsNot Nothing Then r1.IsDeleted = True : RecordSaver.SaveRecord(conn, r1)

    ' DateRange
    Dim r2 = RecordLoader.LoadRecord(Of RecordPatternDateRange)(conn, pk)
    If r2 IsNot Nothing Then r2.IsDeleted = True : RecordSaver.SaveRecord(conn, r2)

    ' Weekly
    Dim r3 = RecordLoader.LoadRecord(Of RecordPatternWeekly)(conn, pk)
    If r3 IsNot Nothing Then r3.IsDeleted = True : RecordSaver.SaveRecord(conn, r3)

    ' WeeklyDays (multi-row)
    Dim weeklyDays =
    RecordLoader.LoadRecordsByFields(Of RecordPatternWeeklyDays)(
      conn, {"AvailabilityID"}, {availabilityID})

    For Each d In weeklyDays
      d.IsDeleted = True
      RecordSaver.SaveRecord(conn, d)
    Next

    ' Monthly
    Dim r4 = RecordLoader.LoadRecord(Of RecordPatternMonthly)(conn, pk)
    If r4 IsNot Nothing Then r4.IsDeleted = True : RecordSaver.SaveRecord(conn, r4)

  End Sub

  ' ==========================================================================================
  ' Routine: DeleteDateRangePattern
  ' Purpose:
  '   Delete tblPatternDateRange for a given AvailabilityID.
  ' ==========================================================================================
  Private Sub DeleteDateRangePattern(ByVal conn As SQLiteConnectionWrapper,
                                   ByVal availabilityID As String)

    Dim pk() As Object = {availabilityID}
    Dim rec As RecordPatternDateRange =
    RecordLoader.LoadRecord(Of RecordPatternDateRange)(conn, pk)

    If rec IsNot Nothing Then
      rec.IsDeleted = True
      RecordSaver.SaveRecord(conn, rec)
    End If

  End Sub


  ' ==========================================================================================
  ' Routine: DeleteWeeklyPatterns
  ' Purpose:
  '   Delete tblPatternWeekly and tblPatternWeeklyDays for a given AvailabilityID.
  ' ==========================================================================================
  Private Sub DeleteWeeklyPatterns(ByVal conn As SQLiteConnectionWrapper,
                                 ByVal availabilityID As String)

    Dim pk() As Object = {availabilityID}
    Dim weekly As RecordPatternWeekly =
    RecordLoader.LoadRecord(Of RecordPatternWeekly)(conn, pk)

    If weekly IsNot Nothing Then
      weekly.IsDeleted = True
      RecordSaver.SaveRecord(conn, weekly)
    End If

    Dim weeklyDays =
    RecordLoader.LoadRecordsByFields(Of RecordPatternWeeklyDays)(
      conn,
      {"AvailabilityID"},
      {availabilityID})

    For Each rec In weeklyDays
      rec.IsDeleted = True
      RecordSaver.SaveRecord(conn, rec)
    Next

  End Sub


  ' ==========================================================================================
  ' Routine: DeleteMonthlyPattern
  ' Purpose:
  '   Delete tblPatternMonthly for a given AvailabilityID.
  ' ==========================================================================================
  Private Sub DeleteMonthlyPattern(ByVal conn As SQLiteConnectionWrapper,
                                 ByVal availabilityID As String)

    Dim pk() As Object = {availabilityID}
    Dim rec As RecordPatternMonthly =
    RecordLoader.LoadRecord(Of RecordPatternMonthly)(conn, pk)

    If rec IsNot Nothing Then
      rec.IsDeleted = True
      RecordSaver.SaveRecord(conn, rec)
    End If

  End Sub

  ' ==========================================================================================
  ' Routine: MapNewValuesToRecord
  ' Purpose:
  '   Build a RecordResourceAvailability for insertion.
  '
  ' Parameters:
  '   a - UIAvailabilityRow (working copy).
  '
  ' Returns:
  '   RecordResourceAvailability ready for SaveRecord.
  '
  ' Notes:
  '   - Generates new GUID for AvailabilityID.
  ' ==========================================================================================
  Private Function MapNewValuesToRecord(ByVal a As UIAvailabilityRow) _
                                        As RecordResourceAvailability

    Dim r As New RecordResourceAvailability()

    r.AvailabilityID = Guid.NewGuid().ToString()
    r.IsNew = True

    r.ResourceID = a.ResourceID
    r.Mode = a.Mode
    r.PatternType = a.PatternType
    r.AllDay = If(a.AllDay, 1, 0)
    r.StartTime = FormatTime(a.StartTime)
    r.EndTime = FormatTime(a.EndTime)

    Return r

  End Function


  ' ==========================================================================================
  ' Routine: MapUpdateValuesToRecord
  ' Purpose:
  '   Build a RecordResourceAvailability for update.
  '
  ' Parameters:
  '   a - UIAvailabilityRow (working copy).
  '
  ' Returns:
  '   RecordResourceAvailability ready for SaveRecord.
  '
  ' Notes:
  '   - Preserves AvailabilityID.
  ' ==========================================================================================
  Private Function MapUpdateValuesToRecord(ByVal a As UIAvailabilityRow) _
                                           As RecordResourceAvailability

    Dim r As New RecordResourceAvailability()

    r.AvailabilityID = a.AvailabilityID
    r.IsDirty = True

    r.ResourceID = a.ResourceID
    r.Mode = a.Mode
    r.PatternType = a.PatternType
    r.AllDay = If(a.AllDay, 1, 0)
    r.StartTime = FormatTime(a.StartTime)
    r.EndTime = FormatTime(a.EndTime)

    Return r

  End Function


  ' ==========================================================================================
  ' Routine: CloneAvailabilityRow
  ' Purpose:
  '   Deep-copy a UIAvailabilityRow so ActionAvailability can diverge from snapshot.
  '
  ' Parameters:
  '   src - source UIAvailabilityRow.
  '
  ' Returns:
  '   UIAvailabilityRow clone.
  '
  ' Notes:
  '   - Copies all fields, including pattern fields.
  ' ==========================================================================================
  Private Function CloneAvailabilityRow(ByVal src As UIAvailabilityRow) _
                                        As UIAvailabilityRow

    If src Is Nothing Then Return New UIAvailabilityRow()

    Dim dst As New UIAvailabilityRow()

    dst.AvailabilityID = src.AvailabilityID
    dst.ResourceID = src.ResourceID
    dst.Mode = src.Mode
    dst.PatternType = src.PatternType
    dst.AllDay = src.AllDay
    dst.StartTime = src.StartTime
    dst.EndTime = src.EndTime

    dst.RangeStart = src.RangeStart
    dst.RangeEndType = src.RangeEndType
    dst.RangeEndDate = src.RangeEndDate
    dst.RangeEndAfterOccurrences = src.RangeEndAfterOccurrences

    dst.DateRangeStart = src.DateRangeStart
    dst.DateRangeEnd = src.DateRangeEnd

    dst.RecurWeeks = src.RecurWeeks
    dst.DaysOfWeek = New List(Of String)(src.DaysOfWeek)

    dst.MonthlyType = src.MonthlyType
    dst.DayOfMonth = src.DayOfMonth
    dst.Ordinal = src.Ordinal
    dst.MonthlyDayOfWeek = src.MonthlyDayOfWeek
    dst.RecurMonths = src.RecurMonths

    Return dst

  End Function

  ' ==========================================================================================
  ' Routine: LoadPatternRange
  ' Purpose:
  '   Load tblPatternRange into the UIAvailabilityRow snapshot.
  ' Parameters:
  '   conn - SQLite connection wrapper.
  '   row  - UIAvailabilityRow to populate.
  ' Returns:
  '   None
  ' Notes:
  '   - Loads only if a record exists and is not deleted.
  ' ==========================================================================================
  Private Sub LoadPatternRange(ByVal conn As SQLiteConnectionWrapper,
                             ByVal row As UIAvailabilityRow)

    Dim pk() As Object = {row.AvailabilityID}
    Dim rec As RecordPatternRange =
        RecordLoader.LoadRecord(Of RecordPatternRange)(conn, pk)

    If rec Is Nothing OrElse rec.IsDeleted Then Exit Sub

    row.RangeStart = ParseDate(rec.StartDate)
    row.RangeEndType = rec.EndType
    row.RangeEndDate = ParseDate(rec.EndDate)
    row.RangeEndAfterOccurrences = rec.EndAfterOccurrences

  End Sub

  ' ==========================================================================================
  ' Routine: LoadPatternDateRange
  ' Purpose:
  '   Load tblPatternDateRange into the UIAvailabilityRow snapshot.
  ' ==========================================================================================
  Private Sub LoadPatternDateRange(ByVal conn As SQLiteConnectionWrapper,
                                 ByVal row As UIAvailabilityRow)

    Dim pk() As Object = {row.AvailabilityID}
    Dim rec As RecordPatternDateRange =
        RecordLoader.LoadRecord(Of RecordPatternDateRange)(conn, pk)

    If rec Is Nothing OrElse rec.IsDeleted Then Exit Sub

    row.DateRangeStart = ParseDate(rec.StartDate)
    row.DateRangeEnd = ParseDate(rec.EndDate)

  End Sub

  ' ==========================================================================================
  ' Routine: LoadPatternWeekly
  ' Purpose:
  '   Load tblPatternWeekly into the UIAvailabilityRow snapshot.
  ' ==========================================================================================
  Private Sub LoadPatternWeekly(ByVal conn As SQLiteConnectionWrapper,
                              ByVal row As UIAvailabilityRow)

    Dim pk() As Object = {row.AvailabilityID}
    Dim rec As RecordPatternWeekly =
        RecordLoader.LoadRecord(Of RecordPatternWeekly)(conn, pk)

    If rec Is Nothing OrElse rec.IsDeleted Then Exit Sub

    row.RecurWeeks = rec.RecurWeeks

  End Sub

  ' ==========================================================================================
  ' Routine: LoadPatternWeeklyDays
  ' Purpose:
  '   Load tblPatternWeeklyDays (multi-row) into UIAvailabilityRow.DaysOfWeek.
  ' ==========================================================================================
  Private Sub LoadPatternWeeklyDays(ByVal conn As SQLiteConnectionWrapper,
                                  ByVal row As UIAvailabilityRow)

    row.DaysOfWeek.Clear()

    Dim list =
        RecordLoader.LoadRecordsByFields(Of RecordPatternWeeklyDays)(
            conn,
            {"AvailabilityID"},
            {row.AvailabilityID})

    For Each rec In list
      If rec IsNot Nothing AndAlso Not rec.IsDeleted Then
        row.DaysOfWeek.Add(rec.DayOfWeek)
      End If
    Next

  End Sub

  ' ==========================================================================================
  ' Routine: LoadPatternMonthly
  ' Purpose:
  '   Load tblPatternMonthly into the UIAvailabilityRow snapshot.
  ' ==========================================================================================
  Private Sub LoadPatternMonthly(ByVal conn As SQLiteConnectionWrapper,
                               ByVal row As UIAvailabilityRow)

    Dim pk() As Object = {row.AvailabilityID}
    Dim rec As RecordPatternMonthly =
        RecordLoader.LoadRecord(Of RecordPatternMonthly)(conn, pk)

    If rec Is Nothing OrElse rec.IsDeleted Then Exit Sub

    row.MonthlyType = rec.MonthlyType
    row.DayOfMonth = rec.DayOfMonth
    row.Ordinal = rec.Ordinal
    row.MonthlyDayOfWeek = rec.DayOfWeek
    row.RecurMonths = rec.RecurMonths

  End Sub

  ' ==========================================================================================
  ' Routine: SavePatternRange
  ' Purpose:
  '   Insert/update tblPatternRange for the working availability row.
  ' ==========================================================================================
  Private Sub SavePatternRange(ByVal conn As SQLiteConnectionWrapper,
                             ByVal a As UIAvailabilityRow)

    Dim pk() As Object = {a.AvailabilityID}
    Dim rec As RecordPatternRange =
        RecordLoader.LoadRecord(Of RecordPatternRange)(conn, pk)

    If rec Is Nothing Then
      rec = New RecordPatternRange()
      rec.AvailabilityID = a.AvailabilityID
      rec.IsNew = True
    Else
      rec.IsDirty = True
    End If

    rec.StartDate = FormatDate(a.RangeStart)
    rec.EndType = a.RangeEndType
    rec.EndDate = FormatDate(a.RangeEndDate)
    rec.EndAfterOccurrences = a.RangeEndAfterOccurrences

    RecordSaver.SaveRecord(conn, rec)

  End Sub

  ' ==========================================================================================
  ' Routine: SavePatternDateRange
  ' Purpose:
  '   Insert/update tblPatternDateRange for the working availability row.
  ' ==========================================================================================
  Private Sub SavePatternDateRange(ByVal conn As SQLiteConnectionWrapper,
                                 ByVal a As UIAvailabilityRow)

    Dim pk() As Object = {a.AvailabilityID}
    Dim rec As RecordPatternDateRange =
        RecordLoader.LoadRecord(Of RecordPatternDateRange)(conn, pk)

    If rec Is Nothing Then
      rec = New RecordPatternDateRange()
      rec.AvailabilityID = a.AvailabilityID
      rec.IsNew = True
    Else
      rec.IsDirty = True
    End If

    rec.StartDate = FormatDate(a.DateRangeStart)
    rec.EndDate = FormatDate(a.DateRangeEnd)

    RecordSaver.SaveRecord(conn, rec)

  End Sub

  ' ==========================================================================================
  ' Routine: SavePatternWeekly
  ' Purpose:
  '   Insert/update tblPatternWeekly for the working availability row.
  ' ==========================================================================================
  Private Sub SavePatternWeekly(ByVal conn As SQLiteConnectionWrapper,
                              ByVal a As UIAvailabilityRow)

    Dim pk() As Object = {a.AvailabilityID}
    Dim rec As RecordPatternWeekly =
        RecordLoader.LoadRecord(Of RecordPatternWeekly)(conn, pk)

    If rec Is Nothing Then
      rec = New RecordPatternWeekly()
      rec.AvailabilityID = a.AvailabilityID
      rec.IsNew = True
    Else
      rec.IsDirty = True
    End If

    rec.RecurWeeks = a.RecurWeeks

    RecordSaver.SaveRecord(conn, rec)

  End Sub

  ' ==========================================================================================
  ' Routine: SavePatternWeeklyDays
  ' Purpose:
  '   Replace all tblPatternWeeklyDays rows for this AvailabilityID.
  '   (Delete existing, insert new.)
  ' ==========================================================================================
  Private Sub SavePatternWeeklyDays(ByVal conn As SQLiteConnectionWrapper,
                                  ByVal a As UIAvailabilityRow)

    ' Delete existing rows
    Dim existing =
        RecordLoader.LoadRecordsByFields(Of RecordPatternWeeklyDays)(
            conn,
            {"AvailabilityID"},
            {a.AvailabilityID})

    For Each rec In existing
      rec.IsDeleted = True
      RecordSaver.SaveRecord(conn, rec)
    Next

    ' Insert new rows
    For Each dow In a.DaysOfWeek
      Dim rec As New RecordPatternWeeklyDays()
      rec.AvailabilityID = a.AvailabilityID
      rec.DayOfWeek = dow
      rec.IsNew = True
      RecordSaver.SaveRecord(conn, rec)
    Next

  End Sub

  ' ==========================================================================================
  ' Routine: SavePatternMonthly
  ' Purpose:
  '   Insert/update tblPatternMonthly for the working availability row.
  ' ==========================================================================================
  Private Sub SavePatternMonthly(ByVal conn As SQLiteConnectionWrapper,
                               ByVal a As UIAvailabilityRow)

    Dim pk() As Object = {a.AvailabilityID}
    Dim rec As RecordPatternMonthly =
        RecordLoader.LoadRecord(Of RecordPatternMonthly)(conn, pk)

    If rec Is Nothing Then
      rec = New RecordPatternMonthly()
      rec.AvailabilityID = a.AvailabilityID
      rec.IsNew = True
    Else
      rec.IsDirty = True
    End If

    rec.MonthlyType = a.MonthlyType
    rec.DayOfMonth = a.DayOfMonth
    rec.Ordinal = a.Ordinal
    rec.DayOfWeek = a.MonthlyDayOfWeek
    rec.RecurMonths = a.RecurMonths

    RecordSaver.SaveRecord(conn, rec)

  End Sub

  ' ==========================================================================================
  ' Helpers: Time parsing/formatting
  ' ==========================================================================================
  Private Function ParseTime(ByVal s As String) As TimeSpan
    If String.IsNullOrEmpty(s) Then Return TimeSpan.Zero
    Return TimeSpan.Parse(s)
  End Function

  Private Function FormatTime(ByVal t As TimeSpan) As String
    If t = TimeSpan.Zero Then Return ""
    Return t.ToString("hh\:mm")
  End Function

  ' ==========================================================================================
  ' Helpers: Date parsing/formatting
  ' ==========================================================================================
  Private Function ParseDate(ByVal s As String) As DateTime
    If String.IsNullOrEmpty(s) Then Return Date.MinValue
    Return Date.ParseExact(s, "yyyy-MM-dd", Nothing)
  End Function

  Private Function FormatDate(ByVal d As DateTime) As String
    If d = Date.MinValue Then Return ""
    Return d.ToString("yyyy-MM-dd")
  End Function

End Module
