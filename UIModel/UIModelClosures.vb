Imports System.ComponentModel

Friend Enum ClosureAction
  None
  Add
  Update
  Delete
End Enum

Friend Class UIModelClosures

  ' The list bound to the UI grid/list
  Friend Property Closures As New SortableBindingList(Of UIClosureRow)

  'Friend Property Closures As BindingList(Of UIClosureRow)

  ' The closure being added, updated, or deleted
  Friend Property ActionClosure As UIClosureRow

  ' The action the UI wants to perform
  Friend Property PendingAction As ClosureAction = ClosureAction.None

End Class

Friend Class UIClosureRow
  Public Property ClosureID As String ' GUID string
  Public Property ClosureName As String
  Public Property StartDate As Date
  Public Property EndDate As Date

End Class
