Imports System.Reflection
Imports System.Windows.Forms
Imports ExcelDna.Integration.CustomUI

Module About
  ' ==========================================================================================
  ' Routine: AboutDisplay
  ' Purpose: Displays version and contact information for the Resource Management Add-in.
  '
  ' Parameters: None
  ' Returns:    None
  '
  ' Notes:
  '   - Uses public constant [version]
  '   - Message includes author contact and disclaimer
  ' ==========================================================================================
  Friend Sub AboutDisplay()

    MessageBoxHelper.Show(UIMessageOwner.ExcelOwner, "You are running Resource Management Add-in version " & Assembly.GetExecutingAssembly().GetName().Version.ToString() & "." & vbCrLf &
           "You can contact the author at ITWorkingSolutions@gmail.com" & vbCrLf &
           "As this is an open source solution all efforts are best endeavour.", "Resource Management Add-in", MessageBoxButtons.OK,
                      MessageBoxIcon.Information)
  End Sub
End Module
