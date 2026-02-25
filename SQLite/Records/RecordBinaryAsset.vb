Option Explicit On

Friend Class RecordBinaryAsset
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
  Private mKey As String
  Private mMimeType As String
  Private mData As Byte()

  ' ------------------------------------------------------------
  '  Properties with change tracking
  ' ------------------------------------------------------------
  Friend Property Key As String
    Get
      Return mKey
    End Get
    Set(value As String)
      If mKey <> value Then
        mKey = value
        IsDirty = True
      End If
    End Set
  End Property

  Friend Property MimeType As String
    Get
      Return mMimeType
    End Get
    Set(value As String)
      If mMimeType <> value Then
        mMimeType = value
        IsDirty = True
      End If
    End Set
  End Property

  Friend Property Data As Byte()
    Get
      Return mData
    End Get
    Set(value As Byte())
      ' Byte() comparison must be reference‑based; deep compare is unnecessary
      If Not Object.ReferenceEquals(mData, value) Then
        mData = value
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
      "Key",
      "MimeType",
      "Data"
    }
  End Function

  Friend Function TableName() As String Implements ISQLiteRecord.TableName
    Return "tblBinaryAsset"
  End Function

  Friend Function PrimaryKeyNames() As String() Implements ISQLiteRecord.PrimaryKeyNames
    Return {"Key"}
  End Function

End Class
