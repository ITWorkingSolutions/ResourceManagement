Friend Enum ResourceListItemAction
  None
  Add
  Update
  Delete
End Enum

Friend Class UIModelResourceListItem

  ' All ResourceListItems
  Friend Property ResourceListItems As SortableBindingList(Of UIResourceListItemRow)

  ' Lookup list for the combo, including explicit "None"
  Friend Property ListItemTypes As SortableBindingList(Of UIListItemTypeRow)

  ' Current action and row being edited
  Friend Property PendingResourceListItemAction As ResourceListItemAction
  Friend Property ActionResourceListItem As UIResourceListItemRow

End Class

Friend Class UIResourceListItemRow

  Friend Property ResourceListItemID As String
  Friend Property ResourceListItemName As String
  Friend Property ValueType As ResourceListItemValueType
  ' May be "" when not tied to any ListItemType (i.e. "None")
  Friend Property ListItemTypeID As String

End Class