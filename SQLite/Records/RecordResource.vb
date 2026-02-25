Option Explicit On

Friend Class RecordResource
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
  Private mResourceID As String
  Private mPreferredName As String
  Private mSalutationID As String
  Private mGenderID As String
  Private mFirstName As String
  Private mLastName As String
  'Private mEmployeeID As String
  Private mEmail As String
  Private mPhone As String
  Private mStartDate As String
  Private mEndDate As String
  Private mNotes As String


  ' ------------------------------------------------------------
  '  Properties with change tracking
  ' ------------------------------------------------------------
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

  Friend Property PreferredName As String
    Get
      Return mPreferredName
    End Get
    Set(value As String)
      If mPreferredName <> value Then
        mPreferredName = value
        IsDirty = True
      End If
    End Set
  End Property

  Friend Property SalutationID As String
    Get
      Return mSalutationID
    End Get
    Set(value As String)
      If mSalutationID <> value Then
        mSalutationID = value
        IsDirty = True
      End If
    End Set
  End Property

  Friend Property GenderID As String
    Get
      Return mGenderID
    End Get
    Set(value As String)
      If mGenderID <> value Then
        mGenderID = value
        IsDirty = True
      End If
    End Set
  End Property
  Friend Property FirstName As String
    Get
      Return mFirstName
    End Get
    Set(value As String)
      If mFirstName <> value Then
        mFirstName = value
        IsDirty = True
      End If
    End Set
  End Property

  Friend Property LastName As String
    Get
      Return mLastName
    End Get
    Set(value As String)
      If mLastName <> value Then
        mLastName = value
        IsDirty = True
      End If
    End Set
  End Property

  'Friend Property EmployeeID As String
  '  Get
  '    Return mEmployeeID
  '  End Get
  '  Set(value As String)
  '    If mEmployeeID <> value Then
  '      mEmployeeID = value
  '      IsDirty = True
  '    End If
  '  End Set
  'End Property

  Friend Property Email As String
    Get
      Return mEmail
    End Get
    Set(value As String)
      If mEmail <> value Then
        mEmail = value
        IsDirty = True
      End If
    End Set
  End Property

  Friend Property Phone As String
    Get
      Return mPhone
    End Get
    Set(value As String)
      If mPhone <> value Then
        mPhone = value
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

  Friend Property Notes As String
    Get
      Return mNotes
    End Get
    Set(value As String)
      If mNotes <> value Then
        mNotes = value
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
  Friend Function TableName() As String Implements ISQLiteRecord.TableName
    Return "tblResource"
  End Function

  Friend Function FieldNames() As String() Implements ISQLiteRecord.FieldNames
    Return {
      "ResourceID",
      "PreferredName",
      "SalutationID",
      "GenderID",
      "FirstName",
      "LastName",
      "Email",
      "Phone",
      "StartDate",
      "EndDate",
      "Notes"
    }
    '"EmployeeID",
  End Function

  Friend Function PrimaryKeyNames() As String() Implements ISQLiteRecord.PrimaryKeyNames
    Return {"ResourceID"}
  End Function

End Class