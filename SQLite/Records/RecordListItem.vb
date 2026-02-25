Friend Class RecordListItem
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
  Private mListItemID As String
  Private mListItemTypeID As String
  Private mListItemName As String

  ' ------------------------------------------------------------
  '  Properties with change tracking
  ' ------------------------------------------------------------
  Friend Property ListItemID As String
    Get
      Return mListItemID
    End Get
    Set(value As String)
      If mListItemID <> value Then
        mListItemID = value
        IsDirty = True
      End If
    End Set
  End Property

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

  Friend Property ListItemName As String
    Get
      Return mListItemName
    End Get
    Set(value As String)
      If mListItemName <> value Then
        mListItemName = value
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
      "ListItemID",
      "ListItemTypeID",
      "ListItemName"
    }
  End Function

  Friend Function TableName() As String Implements ISQLiteRecord.TableName
    Return "tblListItem"
  End Function

  Friend Function PrimaryKeyNames() As String() Implements ISQLiteRecord.PrimaryKeyNames
    Return {"ListItemID"}
  End Function

End Class
