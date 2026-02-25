Friend Enum AvailabilityAction
  None
  Add
  Update
  Delete
End Enum

' ==========================================================================================
' Class: UIModelAvailability
' Purpose:
'   Form-level model for the Availability form.
'   Contains snapshot row, working row, and pending action.
' ==========================================================================================
Friend Class UIModelAvailability

  ' Snapshot from DB
  Friend Property Availability As UIAvailabilityRow

  ' Working copy for UI edits
  Friend Property ActionAvailability As UIAvailabilityRow

  ' Pending action (Add / Update / Delete / None)
  Friend Property PendingAction As AvailabilityAction

  Friend Sub New()
    Availability = New UIAvailabilityRow()
    ActionAvailability = New UIAvailabilityRow()
    PendingAction = AvailabilityAction.None
  End Sub

End Class

' ==========================================================================================
' Class: UIAvailabilityRow
' Purpose:
'   Represents a single availability record (snapshot or working copy).
'   Fully de-normalised and UI-facing only.
' ==========================================================================================
Friend Class UIAvailabilityRow

  ' Core availability fields
  Friend Property AvailabilityID As String
  Friend Property ResourceID As String
  Friend Property Mode As String
  Friend Property PatternType As String
  Friend Property AllDay As Boolean
  Friend Property StartTime As TimeSpan
  Friend Property EndTime As TimeSpan

  ' Range logic
  Friend Property RangeStart As DateTime
  Friend Property RangeEndType As String
  Friend Property RangeEndDate As DateTime
  Friend Property RangeEndAfterOccurrences As Long

  ' DateRange pattern
  Friend Property DateRangeStart As DateTime
  Friend Property DateRangeEnd As DateTime

  ' Weekly pattern
  Friend Property RecurWeeks As Long
  Friend Property DaysOfWeek As List(Of String)

  ' Monthly pattern
  Friend Property MonthlyType As String
  Friend Property DayOfMonth As Long
  Friend Property Ordinal As String
  Friend Property MonthlyDayOfWeek As String
  Friend Property RecurMonths As Long

  Friend Sub New()
    DaysOfWeek = New List(Of String)
  End Sub

End Class