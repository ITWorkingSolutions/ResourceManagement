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

    MessageBoxHelper.Show(UIMessageOwner.ExcelOwner,
            "You are running Resource Management Add-in version " & Assembly.GetExecutingAssembly().GetName().Version.ToString() & "." & vbCrLf & vbCrLf &
            "You can contact the author at ITWorkingSolutions@gmail.com" & vbCrLf &
            "All efforts are best endeavour." & vbCrLf & vbCrLf &
            "License" & vbCrLf &
            "This add-in is licensed under the MIT License." & vbCrLf &
            "A copy of the license is included in the file LICENSE in the project repository." & vbCrLf & vbCrLf &
            "Third-party components" & vbCrLf &
            "This add-in uses the following open-source libraries:" & vbCrLf &
            "- Excel-DNA (zlib license)" & vbCrLf &
            "- Excel-DNA IntelliSense (zlib license)" & vbCrLf &
            "- Microsoft.Data.Sqlite (MIT license)" & vbCrLf &
            "- SQLitePCLRaw (Apache 2.0 license)" & vbCrLf &
            "- Microsoft .NET BCL libraries (MIT license)" & vbCrLf & vbCrLf &
            "Full third-party license texts are included in LICENSES.txt distributed with the add-in.",
            "Resource Management Add-in", MessageBoxButtons.OK,
                      MessageBoxIcon.Question)
  End Sub
End Module
