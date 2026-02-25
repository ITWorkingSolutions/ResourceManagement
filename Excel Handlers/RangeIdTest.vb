'Imports ExcelDna.Integration
'Imports Microsoft.Office.Interop.Excel
'Imports System.Reflection

'Public Module RangeIdTest

'  <ExcelCommand(Description:="Set Range.ID on active cell")>
'  Public Sub SetRangeID()
'    Dim xlApp As Application = CType(ExcelDnaUtil.Application, Application)
'    Dim target As Range = CType(xlApp.ActiveCell, Range)

'    Try
'      ' Late binding to set the ID property
'      target.GetType().InvokeMember("ID",
'                                    BindingFlags.SetProperty,
'                                    Nothing,
'                                    target,
'                                    New Object() {"target-guid-123"})

'      ' Read it back
'      Dim idValue As String = target.GetType().InvokeMember("ID",
'                                                            BindingFlags.GetProperty,
'                                                            Nothing,
'                                                            target,
'                                                            Nothing).ToString()

'      xlApp.StatusBar = $"Address: {target.Address}, ID: {idValue}"
'    Catch ex As Exception
'      xlApp.StatusBar = $"Error: {ex.Message}"
'    End Try
'  End Sub

'  <ExcelCommand(Description:="Get Range.ID for selected range")>
'  Public Sub GetRangeID()

'    Dim xlApp As Application = CType(ExcelDnaUtil.Application, Application)
'    Dim sel As Range = CType(xlApp.Selection, Range)

'    Try
'      Dim sb As New System.Text.StringBuilder()

'      ' ---------------------------------------------------------
'      ' Walk every cell in the selection and read Range.ID
'      ' ---------------------------------------------------------
'      For Each cell As Range In sel.Cells

'        Dim idValue As String = "<null>"

'        Try
'          Dim idObj As Object = cell.GetType().InvokeMember("ID",
'                                        BindingFlags.GetProperty,
'                                        Nothing,
'                                        cell,
'                                        Nothing)
'          If idObj IsNot Nothing Then
'            idValue = idObj.ToString()
'          End If
'        Catch
'          idValue = "<error>"
'        End Try

'        sb.AppendLine($"{cell.Address(False, False)} : {idValue}")

'      Next cell

'      ' ---------------------------------------------------------
'      ' Display results
'      ' ---------------------------------------------------------
'      Dim output As String = sb.ToString()

'      ' If short enough, show in status bar
'      If output.Length < 200 Then
'        xlApp.StatusBar = output
'      Else
'        ' Long output → message box
'        MsgBox(output, vbInformation, "Range.ID Values")
'      End If

'    Catch ex As Exception
'      xlApp.StatusBar = $"Error: {ex.Message}"
'    End Try

'  End Sub

'End Module
