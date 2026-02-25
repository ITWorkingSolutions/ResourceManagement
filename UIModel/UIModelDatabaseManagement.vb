Imports System.ComponentModel
Friend Enum DatabaseManagementAction
  None
  AddNew
  AddExisting
  Update
  Delete
  Activate
End Enum

Friend Class UIModelDatabaseManagement
  ' ------------------------------------------------------------
  '  Existing UI-bound state
  ' ------------------------------------------------------------
  Friend Property Paths As BindingList(Of String)
  Friend Property ActivePath As String

  ' ------------------------------------------------------------
  '  UI > Loader/Saver communication
  '  These are set by the UI BEFORE calling saver/loader methods
  ' ------------------------------------------------------------

  ' The item selected in the list (existing path)
  Friend Property SelectedPath As String

  ' The new path entered/selected by the user
  ' Used for AddNew, AddExisting, Update operations.
  Friend Property NewPath As String

  ' The explicit action the user chose (AddNew, AddExisting, Update, Delete, Activate)
  Friend Property PendingAction As DatabaseManagementAction = DatabaseManagementAction.None


End Class