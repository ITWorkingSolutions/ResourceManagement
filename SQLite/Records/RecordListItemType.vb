Friend Class RecordListItemType
  Implements ISQLiteRecord

  ' ------------------------------------------------------------
  '  Lifecycle flags
  ' ------------------------------------------------------------
  Friend Property IsNew As Boolean Implements ISQLiteRecord.IsNew
  Friend Property IsDirty As Boolean Implements ISQLiteRecord.IsDirty
  Friend Property IsDeleted As Boolean Implements ISQLiteRecord.IsDeleted

  ' ------------------------------------------------------------
  '  Private backing fields
  ' ------------------------------------------------------------
  Private mListItemTypeID As String
  Private mListItemTypeName As String
  Private mIsSystemType As Long

  ' ------------------------------------------------------------
  '  Properties
  ' ------------------------------------------------------------
  Friend Property ListItemTypeID As String
    Get
      Return mListItemTypeID
    End Get
    Set(value As String)
      If mListItemTypeID <> value Then
        mListItemTypeID = value
        IsDirty = True
      End If
    End Set
  End Property

  Friend Property ListItemTypeName As String
    Get
      Return mListItemTypeName
    End Get
    Set(value As String)
      If mListItemTypeName <> value Then
        mListItemTypeName = value
        IsDirty = True
      End If
    End Set
  End Property

  Friend Property IsSystemType As Long
    Get
      Return mIsSystemType
    End Get
    Set(value As Long)
      If mIsSystemType <> value Then
        mIsSystemType = value
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
      "ListItemTypeID",
      "ListItemTypeName",
      "IsSystemType"
    }
  End Function

  Friend Function TableName() As String Implements ISQLiteRecord.TableName
    Return "tblListItemType"
  End Function

  Friend Function PrimaryKeyNames() As String() Implements ISQLiteRecord.PrimaryKeyNames
    Return {"ListItemTypeID"}
  End Function

End Class