Attribute VB_Name = "modBuildAll"
Option Explicit
'==============================================================================
' Entry points. The "Build Schedules" button on INPUT runs BuildAllSchedules.
'==============================================================================
Private Sub RunBuild(ByVal quiet As Boolean)
    On Error GoTo EH
    Application.ScreenUpdating = False
    BuildDetail
    BuildQtrlySum
    Application.ScreenUpdating = True
    If Not quiet Then MsgBox "DETAIL and QTRLY SUM have been rebuilt.", vbInformation, "Build Schedules"
    Exit Sub
EH:
    Application.ScreenUpdating = True
    If quiet Then Err.Raise Err.Number, , Err.Description
    MsgBox "Build failed: " & Err.Description, vbExclamation, "Build Schedules"
End Sub

Public Sub BuildAllSchedules()
    RunBuild False
End Sub

Public Sub BuildAllSchedulesQuiet()
    RunBuild True
End Sub
