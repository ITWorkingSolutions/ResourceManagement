Friend Class UserFriendlyException
  Inherits Exception

  Friend Sub New(message As String)
    MyBase.New(message)
  End Sub

  Friend Sub New(message As String, inner As Exception)
    MyBase.New(message, inner)
  End Sub

End Class
