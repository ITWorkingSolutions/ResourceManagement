Friend Enum ResourceAction
  None
  Add
  Update
  Delete
End Enum

Friend Class UIModelResources

  ' --- Core resource record (fixed fields only) ---
  Friend Property Resource As UIResourceRow

  ' --- Lookups for fixed fields ---
  Friend Property Salutations As List(Of UIListItemRow)
  Friend Property Genders As List(Of UIListItemRow)

  ' --- Labels for fixed fields ---
  Friend Property SalutationListItemTypeName As String
  Friend Property GenderListItemTypeName As String

  ' --- Name/value pairs (e.g. Department, Location, and any other extensible attributes) ---
  Friend Property ResourceNameValues As SortableBindingList(Of UIResourceNameValueRow)

  ' --- Pending action (Add/Update/Delete) ---
  Friend Property PendingAction As ResourceAction = ResourceAction.None

  ' --- Action objects for commit ---
  Friend Property ActionResource As UIResourceRow
  Friend Property ActionResourceNameValues As SortableBindingList(Of UIResourceNameValueRow)

  Friend Sub New()
    Salutations = New List(Of UIListItemRow)()
    Genders = New List(Of UIListItemRow)()

    ResourceNameValues = New SortableBindingList(Of UIResourceNameValueRow)()
    ActionResourceNameValues = New SortableBindingList(Of UIResourceNameValueRow)()
  End Sub

End Class

Friend Class UIResourceRow
  ' --- Identity & naming ---
  Friend Property ResourceID As String
  Friend Property PreferredName As String
  Friend Property GenderID As String
  Friend Property SalutationID As String
  Friend Property FirstName As String
  Friend Property LastName As String

  ' --- Contact ---
  Friend Property Email As String
  Friend Property Phone As String

  ' --- Lifecycle ---
  Friend Property StartDate As Date
  Friend Property EndDate As Date

  ' --- Meta ---
  Friend Property Notes As String

End Class

Friend Enum ResourceNameValueAction
  None
  Add
  Update
  Delete
End Enum

Friend Class UIResourceNameValueRow
  ' The attribute being set (e.g., Department, Location)
  Public Property ResourceListItemID As String

  ' Display name for the left column (resolved from ListItem table)
  Public Property ResourceListItemName As String

  ' If this attribute is based on a ListItemType, this tells the UI which lookup to use
  Public Property ListItemTypeID As String

  ' --- VALUE STORAGE ---

  ' For SingleSelectList selected item
  Public Property SelectedListItemID As String

  ' For MultiSelectList selected items
  Public Property SelectedListItemIDs As List(Of String)

  ' The free-text value (if not lookup-based)
  Public Property ResourceListItemValue As String

  ' Lookup items for the lists (if ValueType is SingleSelectList or MultiSelectList)
  Public Property ListItems As List(Of UIListItemRow)

  ' The data type of the value
  Public Property ValueType As ResourceListItemValueType

  ' Per-row pending action in ActionResourceRoleFunctions
  Public Property PendingAction As ResourceNameValueAction = ResourceNameValueAction.None
End Class