Friend Enum ListItemTypeAction
  None
  Add
  Update
  Delete
End Enum

Friend Enum ListItemAction
  None
  Add
  Update
  Delete
End Enum

Friend Class UIModelListItemTypeAndItem

  ' === ListItemTypes ===
  Friend Property ListItemTypes As SortableBindingList(Of UIListItemTypeRow)
  Friend Property SelectedListItemType As UIListItemTypeRow
  Friend Property ActionListItemType As UIListItemTypeRow
  Friend Property PendingListItemTypeAction As ListItemTypeAction = ListItemTypeAction.None

  ' === ListItems ===
  Friend Property ListItems As SortableBindingList(Of UIListItemRow)
  Friend Property SelectedListItem As UIListItemRow
  Friend Property ActionListItem As UIListItemRow
  Friend Property PendingListItemAction As ListItemAction = ListItemAction.None

End Class

Friend Class UIListItemTypeRow
  Public Property ListItemTypeID As String
  Public Property ListItemTypeName As String
  Public Property IsSystemType As Boolean
End Class

Friend Class UIListItemRow
  Public Property ListItemID As String ' GUID string
  Public Property ListItemTypeID As String
  Public Property ListItemName As String
End Class