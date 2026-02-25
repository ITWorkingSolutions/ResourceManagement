Option Explicit On
Friend Class RecordClosure
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
  Private mClosureID As String
  Private mClosureName As String
  Private mStartDate As String
  Private mEndDate As String

  ' ------------------------------------------------------------
  '  Properties with change tracking
  ' ------------------------------------------------------------
  Friend Property ClosureID As String
    Get
      Return mClosureID
    End Get
    Set(value As String)
      If mClosureID <> value Then
        mClosureID = value
        IsDirty = True
      End If
    End Set
  End Property

  Friend Property ClosureName As String
    Get
      Return mClosureName
    End Get
    Set(value As String)
      If mClosureName <> value Then
        mClosureName = value
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
          "ClosureID",
          "ClosureName",
          "StartDate",
          "EndDate"
      }
  End Function
  Friend Function TableName() As String Implements ISQLiteRecord.TableName
    Return "tblClosure"
  End Function

  Friend Function PrimaryKeyNames() As String() Implements ISQLiteRecord.PrimaryKeyNames
    Return {"ClosureID"}
  End Function
End Class
