Imports System.Windows.Forms
Imports System.Runtime.CompilerServices

Module ErrorHandler
  Friend Sub UnHandleError(ex As Exception,
                         <CallerMemberName> Optional routine As String = Nothing)
    MessageBoxHelper.Show(UIMessageOwner.ExcelOwner,
        "Error #" & ex.HResult & " in " & routine & vbCrLf & ex.Message,
        "Unhandled Error",
        MessageBoxButtons.OK,
        MessageBoxIcon.Error
    )
  End Sub
End Module
