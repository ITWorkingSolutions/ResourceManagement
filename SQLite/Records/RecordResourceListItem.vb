Friend Class RecordResourceListItem
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
  Private mResourceListItemID As String
  Private mResourceListItemName As String
  Private mListItemTypeID As String
  Private mValueType As String

  ' ------------------------------------------------------------
  '  Properties
  ' ------------------------------------------------------------
  Friend Property ResourceListItemID As String
    Get
      Return mResourceListItemID
    End Get
    Set(value As String)
      If mResourceListItemID <> value Then
        mResourceListItemID = value
        IsDirty = True
      End If
    End Set
  End Property

  Friend Property ResourceListItemName As String
    Get
      Return mResourceListItemName
    End Get
    Set(value As String)
      If mResourceListItemName <> value Then
        mResourceListItemName = value
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

  ' The value mode of the resource value 
  Friend Property ValueType As String
    Get
      Return mValueType
    End Get
    Set(value As String)
      If mValueType <> value Then
        mValueType = value
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
      "ResourceListItemID",
      "ResourceListItemName",
      "ListItemTypeID",
      "ValueType"
    }
  End Function

  Friend Function TableName() As String Implements ISQLiteRecord.TableName
    Return "tblResourceListItem"
  End Function

  Friend Function PrimaryKeyNames() As String() Implements ISQLiteRecord.PrimaryKeyNames
    Return {"ResourceListItemID"}
  End Function

End Class
