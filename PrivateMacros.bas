Attribute VB_Name = "PrivateMacros"
Option Explicit
Sub DisableForEfficiency()

' -----------
' Turns off functionality to speed up Excel
' -----------

    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.DisplayStatusBar = False

End Sub
Sub DisableForEfficiencyOff()

' -----------
' Turns functionality back on
' -----------

    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    Application.DisplayStatusBar = True

End Sub

Sub RemoveAllCustomStyles()
' Removes all custom styles from the active workbook
' User must confirm removal
' I recommend saving the file before running the macro
'
' WARNING: Macro may take many minutes to run if there are
' a large number of custom styles or the file is large
 
Dim tmpSt As Style
Dim Wkb As Workbook
 
On Error GoTo HandleExit
 
Set Wkb = ActiveWorkbook

  For Each tmpSt In Wkb.Styles
    With tmpSt
      If .BuiltIn = False Then
        .Locked = False
        .Delete
      End If
    End With
  Next tmpSt
 
HandleExit:
Set tmpSt = Nothing
Set Wkb = Nothing
 
End Sub
