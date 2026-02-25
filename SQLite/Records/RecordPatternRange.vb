Option Explicit On

Friend Class RecordPatternRange
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
  Private mStartDate As String
  Private mEndType As String
  Private mEndDate As String
  Private mEndAfterOccurrences As Long

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

  Friend Property StartDate As String
    Get
      Return mStartDate
    End Get
    Set(value As String)
      If mStartDate <> value Then
        mStartDate = value
        IsDirty = True
      End If
    End Set
  End Property

  Friend Property EndType As String
    Get
      Return mEndType
    End Get
    Set(value As String)
      If mEndType <> value Then
        mEndType = value
        IsDirty = True
      End If
    End Set
  End Property

  Friend Property EndDate As String
    Get
      Return mEndDate
    End Get
    Set(value As String)
      If mEndDate <> value Then
        mEndDate = value
        IsDirty = True
      End If
    End Set
  End Property

  Friend Property EndAfterOccurrences As Long
    Get
      Return mEndAfterOccurrences
    End Get
    Set(value As Long)
      If mEndAfterOccurrences <> value Then
        mEndAfterOccurrences = value
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
        "StartDate",
        "EndType",
        "EndDate",
        "EndAfterOccurrences"
    }
  End Function
  Friend Function TableName() As String Implements ISQLiteRecord.TableName
    Return "tblPatternRange"
  End Function

  Friend Function PrimaryKeyNames() As String() Implements ISQLiteRecord.PrimaryKeyNames
    Return {"AvailabilityID"}
  End Function
End Class
