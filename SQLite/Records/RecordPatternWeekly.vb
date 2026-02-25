Option Explicit On

Friend Class RecordPatternWeekly
  Implements ISQLiteRecord

  ' ------------------------------------------------------------
  '  Lifecycle flags (runtime only)
  ' ------------------------------------------------------------
  Friend Property IsNew As Boolean Implements ISQLiteRecord.IsNew
  Friend Property IsDirty As Boolean Implements ISQLiteRecord.IsDirty
  Friend Property IsDeleted As Boolean Implements ISQLiteRecord.IsDeleted


  ' ------------------------------------------------------------
  '  Private backing fields
  ' ------------------------------------------------------------
  Private mAvailabilityID As String
  Private mRecurWeeks As Long

  ' ------------------------------------------------------------
  '  Properties with change tracking
  ' ------------------------------------------------------------
  Friend Property AvailabilityID As String
    Get
      Return mAvailabilityID
    End Get
    Set(value As String)
      If mAvailabilityID <> value Then
        mAvailabilityID = value
        IsDirty = True
      End If
    End Set
  End Property

  Friend Property RecurWeeks As Long
    Get
      Return mRecurWeeks
    End Get
    Set(value As Long)
      If mRecurWeeks <> value Then
        mRecurWeeks = value
        IsDirty = True
      End If
    End Set
  End Property

  ' ------------------------------------------------------------
  '  Logical delete
  ' ------------------------------------------------------------
  Friend Sub Delete()
    IsDeleted = True
  End Sub

  ' ------------------------------------------------------------
  '  ISQLiteRecord implementation
  ' ------------------------------------------------------------
  Friend Function FieldNames() As String() Implements ISQLiteRecord.FieldNames
    Return {
        "AvailabilityID",
        "RecurWeeks"
    }
  End Function
  Friend Function TableName() As String Implements ISQLiteRecord.TableName
    Return "tblPatternWeekly"
  End Function

  Friend Function PrimaryKeyNames() As String() Implements ISQLiteRecord.PrimaryKeyNames
    Return {"AvailabilityID"}
  End Function
End Class
