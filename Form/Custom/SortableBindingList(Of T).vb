Imports System.ComponentModel

' ===========================================================================
' SortableBindingList(Of T)
'
' WinForms DataGridView controls can only perform column sorting when the underlying data source
' implements sorting support. The standard BindingList(Of T) does NOT implement sorting, which means:
'
'   • DataGridView.Sort(...) throws a "data-bound grid cannot be sorted" exception
'   • BindingSource.Sort is ignored
'   • Columns with SortMode = Automatic still cannot sort
'
' This class inherits from BindingList(Of T) and adds the minimal sorting support required by both
' BindingSource and DataGridView. It preserves all normal BindingList behaviour (add/remove/change
' notifications, strong typing, stable identity) while enabling UI-level sorting without modifying
' the underlying data model.
'
' Key points:
'   • Sorting affects only the view order, not the underlying objects
'   • No changes are required in existing code that uses the list
'   • Safe for use as a DataSource for BindingSource and DataGridView
'   • Generic and reusable for any T with sortable properties
'
' This class is intentionally generic and domain‑agnostic so it can be reused anywhere a sortable
' BindingList is required in WinForms applications.
' ===========================================================================
Friend Class SortableBindingList(Of T)
  Inherits BindingList(Of T)

  Private _isSorted As Boolean
  Private _sortDirection As ListSortDirection
  Private _sortProperty As PropertyDescriptor

  Protected Overrides ReadOnly Property SupportsSortingCore As Boolean
    Get
      Return True
    End Get
  End Property

  Protected Overrides ReadOnly Property IsSortedCore As Boolean
    Get
      Return _isSorted
    End Get
  End Property

  Protected Overrides ReadOnly Property SortDirectionCore As ListSortDirection
    Get
      Return _sortDirection
    End Get
  End Property

  Protected Overrides ReadOnly Property SortPropertyCore As PropertyDescriptor
    Get
      Return _sortProperty
    End Get
  End Property

  Protected Overrides Sub ApplySortCore(prop As PropertyDescriptor, direction As ListSortDirection)
    Dim items = CType(Me.Items, List(Of T))

    items.Sort(Function(x, y)
                 Dim xValue = prop.GetValue(x)
                 Dim yValue = prop.GetValue(y)

                 Return Comparer.Default.Compare(xValue, yValue) *
                        If(direction = ListSortDirection.Ascending, 1, -1)
               End Function)

    _sortProperty = prop
    _sortDirection = direction
    _isSorted = True

    Me.OnListChanged(New ListChangedEventArgs(ListChangedType.Reset, -1))
  End Sub

  Protected Overrides Sub RemoveSortCore()
    _isSorted = False
    _sortProperty = Nothing
  End Sub
End Class
