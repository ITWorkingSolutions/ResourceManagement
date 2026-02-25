Option Explicit On

Friend Class RecordResourceAvailability
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
  Private mResourceID As String
  Private mMode As String
  Private mPatternType As String
  Private mAllDay As Long        ' stored as 0 or 1
  Private mStartTime As String
  Private mEndTime As String

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

  Friend Property ResourceID As String
    Get
      Return mResourceID
    End Get
    Set(value As String)
      If mResourceID <> value Then
        mResourceID = value
        IsDirty = True
      End If
    End Set
  End Property

  Friend Property Mode As String
    Get
      Return mMode
    End Get
    Set(value As String)
      If mMode <> value Then
        mMode = value
        IsDirty = True
      End If
    End Set
  End Property

  Friend Property PatternType As String
    Get
      Return mPatternType
    End Get
    Set(value As String)
      If mPatternType <> value Then
        mPatternType = value
        IsDirty = True
      End If
    End Set
  End Property

  Friend Property AllDay As Long
    Get
      Return mAllDay
    End Get
    Set(value As Long)
      If mAllDay <> value Then
        mAllDay = value
        IsDirty = True
      End If
    End Set
  End Property

  Friend Property StartTime As String
    Get
      Return mStartTime
    End Get
    Set(value As String)
      If mStartTime <> value Then
        mStartTime = value
        IsDirty = True
      End If
    End Set
  End Property

  Friend Property EndTime As String
    Get
      Return mEndTime
    End Get
    Set(value As String)
      If mEndTime <> value Then
        mEndTime = value
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
        "ResourceID",
        "Mode",
        "PatternType",
        "AllDay",
        "StartTime",
        "EndTime"
    }
  End Function
  Friend Function TableName() As String Implements ISQLiteRecord.TableName
    Return "tblResourceAvailability"
  End Function

  Friend Function PrimaryKeyNames() As String() Implements ISQLiteRecord.PrimaryKeyNames
    Return {"AvailabilityID"}
  End Function

End Class