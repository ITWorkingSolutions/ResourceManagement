Option Explicit On

Friend Class RecordPatternMonthly
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
  Private mMonthlyType As String
  Private mDayOfMonth As Long
  Private mOrdinal As String
  Private mDayOfWeek As String
  Private mRecurMonths As Long

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

  Friend Property MonthlyType As String
    Get
      Return mMonthlyType
    End Get
    Set(value As String)
      If mMonthlyType <> value Then
        mMonthlyType = value
        IsDirty = True
      End If
    End Set
  End Property

  Friend Property DayOfMonth As Long
    Get
      Return mDayOfMonth
    End Get
    Set(value As Long)
      If mDayOfMonth <> value Then
        mDayOfMonth = value
        IsDirty = True
      End If
    End Set
  End Property

  Friend Property Ordinal As String
    Get
      Return mOrdinal
    End Get
    Set(value As String)
      If mOrdinal <> value Then
        mOrdinal = value
        IsDirty = True
      End If
    End Set
  End Property

  Friend Property DayOfWeek As String
    Get
      Return mDayOfWeek
    End Get
    Set(value As String)
      If mDayOfWeek <> value Then
        mDayOfWeek = value
        IsDirty = True
      End If
    End Set
  End Property

  Friend Property RecurMonths As Long
    Get
      Return mRecurMonths
    End Get
    Set(value As Long)
      If mRecurMonths <> value Then
        mRecurMonths = value
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
        "MonthlyType",
        "DayOfMonth",
        "Ordinal",
        "DayOfWeek",
        "RecurMonths"
    }
  End Function
  Friend Function TableName() As String Implements ISQLiteRecord.TableName
    Return "tblPatternMonthly"
  End Function

  Friend Function PrimaryKeyNames() As String() Implements ISQLiteRecord.PrimaryKeyNames
    Return {"AvailabilityID"}
  End Function
End Class
