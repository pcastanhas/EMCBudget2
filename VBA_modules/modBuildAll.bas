Attribute VB_Name = "modBuildAll"
Option Explicit
'==============================================================================
' Entry points. The "Build Schedules" button on INPUT runs BuildAllSchedules.
'==============================================================================
Private Sub RunBuild(ByVal quiet As Boolean)
    On Error GoTo EH
    Application.ScreenUpdating = False
    BuildRevenue
    BuildDetail
    BuildProforma
    BuildQtrlySum
    Application.ScreenUpdating = True
    If Not quiet Then MsgBox "REVENUE, DETAIL, PROFORMA and QTRLY SUM have been rebuilt.", vbInformation, "Build Schedules"
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

' Runs when the workbook is opened with macros enabled. Excel keeps the +/- outline buttons dead on a protected sheet unless EnableOutlining is set in code
' (the setting is not saved in the file), so re-apply it on the input sheets whose detail rows are grouped (ExpSheetName 1-11). No password is used.
Public Sub Auto_Open()
    Dim i As Long, ws As Worksheet
    On Error Resume Next
    For i = 1 To 11
        Set ws = ThisWorkbook.Worksheets(ExpSheetName(i))
        If ws.ProtectContents Then
            ws.Unprotect
            ws.EnableOutlining = True
            ws.Protect DrawingObjects:=True, Contents:=True, Scenarios:=True, UserInterfaceOnly:=True, AllowFormattingColumns:=True, AllowFormattingRows:=True
        End If
    Next i
    On Error GoTo 0
End Sub
