Imports System.Windows.Forms
Friend Class SelectDefaultMode
  Friend Property SelectedMode As String = Nothing

  Friend Sub New()
    Try
      ' Disable WinForms autoscaling completely
      Me.AutoScaleMode = AutoScaleMode.None

      InitializeComponent()

      ' Apply manual DPI scaling AFTER controls exist
      ApplyDpiScaling(Me)

    Catch ex As Exception
      ErrorHandler.UnHandleError(ex)

    Finally
      ' Cleanup
    End Try
  End Sub
  Private Sub SelectDefaultMode_Load(sender As Object, e As EventArgs) Handles MyBase.Load
    ' Center relative to Excel if helper available
    Try
      FormHelpers.CenterFormOnExcel(Me)
    Catch
      ' ignore if helper not present
    End Try
  End Sub
  Private Sub btnAvailable_Click(sender As Object, e As EventArgs) _
      Handles btnAvailable.Click
    SelectedMode = "Available"
  End Sub

  Private Sub btnUnavailable_Click(sender As Object, e As EventArgs) _
      Handles btnUnavailable.Click

    SelectedMode = "Unavailable"
  End Sub

  Private Sub btnCancel_Click(sender As Object, e As EventArgs) _
      Handles btnCancel.Click

    SelectedMode = Nothing
  End Sub
End Class