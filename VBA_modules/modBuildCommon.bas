Attribute VB_Name = "modBuildCommon"
Option Explicit
'==============================================================================
' Shared helpers for the budget-schedule builders.
' These macros only WRITE FORMULAS AND FORMATTING. All numbers stay in live
' worksheet formulas, so the workbook calculates correctly with macros disabled;
' macros are needed only to rebuild the schedule layout (rows / sections).
'==============================================================================
Public Const REV_SHEET As String = "LEASE_CALC"   ' tenant-revenue calculation sheet (in-place leases + lease deals); the purple REVENUE tab only presents it
Public Const REV_FIRST As Long = 8          ' first LEASE_CALC row scanned
Public Const REV_LAST As Long = 1500        ' last LEASE_CALC row scanned
Public Const OR_FIRST As Long = 9           ' first OTHER_REVENUE row scanned
Public Const OR_LAST As Long = 600          ' last OTHER_REVENUE row scanned
Public Const EXP_FIRST As Long = 9          ' first EXP_ sheet row scanned
Public Const EXP_LAST As Long = 1500        ' last EXP_ sheet row scanned
Public Const ACCT_MIN As Double = 1000000   ' GL numbers are 7 digits (>= 1,000,000)
Public Const RC_SHEET As String = "REPAIRS_CAPITAL"   ' major repairs + capital improvements input sheet
Public Const RC_MCOL As Long = 28           ' REPAIRS_CAPITAL: column AB = first budget-by-month column (AB:AM), total in AN, check text in AO
Public Const NUMFMT As String = "#,##0;(#,##0);""-"""
Public Const NUMFMT2 As String = "#,##0.00;(#,##0.00);""-"""

' Module-level copies of the chart of accounts used to classify accounts found on LEASE_CALC / OTHER_REVENUE
' (module-level declarations must stay above the first procedure)
Private mvNo As Variant, mvTy As Variant, mvFG As Variant
Private mwsI As Worksheet

' Column number -> column letter(s)
Public Function ColL(ByVal c As Long) As String
    Dim s As String, n As Long
    n = c
    Do While n > 0
        s = Chr(65 + ((n - 1) Mod 26)) & s
        n = (n - 1) \ 26
    Loop
    ColL = s
End Function

' Formula text: amount for one account in one month column, from LEASE_CALC plus the
' OTHER_REVENUE input lines. revCol = month column letter (G..R on both sheets);
' acctRef = cell holding the GL number (e.g. $A9).
Public Function RevSumTerm(ByVal revCol As String, ByVal acctRef As String) As String
    RevSumTerm = "SUMIFS(" & REV_SHEET & "!" & revCol & "$" & REV_FIRST & ":" & revCol & "$" & REV_LAST & _
        "," & REV_SHEET & "!$D$" & REV_FIRST & ":$D$" & REV_LAST & "," & acctRef & ")" & _
        "+SUMIFS(OTHER_REVENUE!" & revCol & "$" & OR_FIRST & ":" & revCol & "$" & OR_LAST & _
        ",OTHER_REVENUE!$D$" & OR_FIRST & ":$D$" & OR_LAST & "," & acctRef & _
        ",OTHER_REVENUE!$E$" & OR_FIRST & ":$E$" & OR_LAST & ",""INPUT"")"
End Function

' Formula text: total DETAIL should show = signed LEASE_CALC amounts (LEASE_CALC col W), the OTHER_REVENUE input lines and the section input lines
Public Function AcctTotalTerm() As String
    Dim t As String, i As Long, sh As String
    t = "SUM(" & REV_SHEET & "!$W$" & REV_FIRST & ":$W$" & REV_LAST & ")" & _
        "+SUMIFS(OTHER_REVENUE!$S$" & OR_FIRST & ":$S$" & OR_LAST & ",OTHER_REVENUE!$D$" & OR_FIRST & ":$D$" & OR_LAST & ","">=1000000"",OTHER_REVENUE!$E$" & OR_FIRST & ":$E$" & OR_LAST & ",""INPUT"")"
    For i = 1 To 12
        sh = ExpSheetName(i)
        t = t & "+SUMIFS(" & sh & "!$" & SheetTotCol(sh) & "$" & EXP_FIRST & ":$" & SheetTotCol(sh) & "$" & EXP_LAST & "," & sh & "!$D$" & EXP_FIRST & ":$D$" & EXP_LAST & ","">=1000000""," & sh & "!$E$" & EXP_FIRST & ":$E$" & EXP_LAST & ",""" & ExpTag(i + 3) & """)"
    Next i
    t = t & "+SUMIFS(" & RC_SHEET & "!$" & SheetTotCol(RC_SHEET) & "$" & EXP_FIRST & ":$" & SheetTotCol(RC_SHEET) & "$" & EXP_LAST & "," & RC_SHEET & "!$D$" & EXP_FIRST & ":$D$" & EXP_LAST & ","">=1000000""," & RC_SHEET & "!$E$" & EXP_FIRST & ":$E$" & EXP_LAST & ",""REPAIR"")"
    AcctTotalTerm = t
End Function

' The six expense input sheets plus TAXES, INSURANCE, PROF_FEES, DEBT, OTHER_PROF_FEES and REPAIRS_CAPITAL (section number = index + 3, except REPAIRS_CAPITAL: see SheetSec)
Public Function ExpSheetName(ByVal i As Long) As String
    Select Case i
        Case 1: ExpSheetName = "EXP_UTILITIES"
        Case 2: ExpSheetName = "EXP_JANITORIAL"
        Case 3: ExpSheetName = "EXP_ELEVATOR"
        Case 4: ExpSheetName = "EXP_HVAC"
        Case 5: ExpSheetName = "EXP_LANDSCAPING"
        Case 6: ExpSheetName = "EXP_GENERAL_RM"
        Case 7: ExpSheetName = "TAXES"
        Case 8: ExpSheetName = "INSURANCE"
        Case 9: ExpSheetName = "PROF_FEES"
        Case 10: ExpSheetName = "DEBT"
        Case 11: ExpSheetName = "OTHER_PROF_FEES"
        Case 12: ExpSheetName = RC_SHEET
    End Select
End Function

' DETAIL section that a section sheet feeds: index + 3, except REPAIRS_CAPITAL (index 12): its capital improvement lines list under section 3 (IMPRVMTS & COMMISSIONS)
Public Function SheetSec(ByVal i As Long) As Long
    If i = 12 Then
        SheetSec = 3
    Else
        SheetSec = i + 3
    End If
End Function

' Month column letter on a section sheet. Every section sheet keeps its months in G:R except REPAIRS_CAPITAL (AB:AM).
Public Function SheetMonthCol(ByVal sh As String, ByVal revCol As String) As String
    If sh = RC_SHEET Then
        SheetMonthCol = ColL(RC_MCOL + Asc(revCol) - Asc("G"))
    Else
        SheetMonthCol = revCol
    End If
End Function

' Total column letter on a section sheet (S everywhere except REPAIRS_CAPITAL: AN)
Public Function SheetTotCol(ByVal sh As String) As String
    If sh = RC_SHEET Then
        SheetTotCol = ColL(RC_MCOL + 12)
    Else
        SheetTotCol = "S"
    End If
End Function

' Formula text: major repair amount for one account in one month column (REPAIRS_CAPITAL lines tagged REPAIR)
Public Function RepSumTerm(ByVal revCol As String, ByVal acctRef As String) As String
    Dim mc As String
    mc = SheetMonthCol(RC_SHEET, revCol)
    RepSumTerm = "SUMIFS(" & RC_SHEET & "!" & mc & "$" & EXP_FIRST & ":" & mc & "$" & EXP_LAST & "," & RC_SHEET & "!$D$" & EXP_FIRST & ":$D$" & EXP_LAST & "," & acctRef & _
        "," & RC_SHEET & "!$E$" & EXP_FIRST & ":$E$" & EXP_LAST & ",""REPAIR"")"
End Function

' Row tag that marks the budget-year amount lines on a section sheet (EXP_ sheets: INPUT lines; TAXES, INSURANCE, PROF_FEES and DEBT: FINAL total rows)
Public Function ExpTag(ByVal sec As Long) As String
    If sec >= 10 Then
        ExpTag = "FINAL"
    Else
        ExpTag = "INPUT"
    End If
End Function

' Formula text: budget-year amount for one account in one month column, from the INPUT lines of its section sheet
Public Function ExpSumTerm(ByVal sec As Long, ByVal revCol As String, ByVal acctRef As String) As String
    Dim sh As String, mc As String
    sh = ExpSheetName(sec - 3)
    mc = SheetMonthCol(sh, revCol)
    ExpSumTerm = "SUMIFS(" & sh & "!" & mc & "$" & EXP_FIRST & ":" & mc & "$" & EXP_LAST & "," & sh & "!$D$" & EXP_FIRST & ":$D$" & EXP_LAST & "," & acctRef & _
        "," & sh & "!$E$" & EXP_FIRST & ":$E$" & EXP_LAST & ",""" & ExpTag(sec) & """)"
End Function

' Budget-year amount for one account in one month, as DETAIL / QTRLY SUM show it:
'   section 1 (income)           = amounts on LEASE_CALC + OTHER_REVENUE input lines
'   section 3 (leasing costs)    = cost lines on LEASE_CALC less any tenant-billed amounts on the same account, plus REPAIRS_CAPITAL capital improvement lines and major repairs
'   sections 4-14 (expenses etc) = amounts on the section sheet plus REPAIRS_CAPITAL major repairs, less tenant-billed amounts on LEASE_CALC for the same account
Public Function AcctMonthTerm(ByVal sec As Long, ByVal revCol As String, ByVal acctRef As String) As String
    Dim rv As String
    rv = "" & REV_SHEET & "!" & revCol & "$" & REV_FIRST & ":" & revCol & "$" & REV_LAST & "," & REV_SHEET & "!$D$" & REV_FIRST & ":$D$" & REV_LAST & "," & acctRef & _
         "," & REV_SHEET & "!$V$" & REV_FIRST & ":$V$" & REV_LAST & ",""<>IGNORED*""," & REV_SHEET & "!$E$" & REV_FIRST & ":$E$" & REV_LAST & ","
    Select Case sec
        Case 1
            AcctMonthTerm = RevSumTerm(revCol, acctRef)
        Case 3
            AcctMonthTerm = "SUMIFS(" & rv & """COST"")-SUMIFS(" & rv & """<>COST"")+" & ExpSumTerm(15, revCol, acctRef) & "+" & RepSumTerm(revCol, acctRef)      ' + capital improvements (REPAIRS_CAPITAL FINAL rows) + major repairs
        Case Else
            AcctMonthTerm = ExpSumTerm(sec, revCol, acctRef) & "+" & RepSumTerm(revCol, acctRef) & "-SUMIFS(" & rv & """<>COST"")"
    End Select
End Function

' Section display order on DETAIL / QTRLY SUM: income, tenant charges, six expense sections, taxes, insurance, professional fees, other professional fees, debt service, leasing costs
Public Function SecOrder(ByVal i As Long) As Long
    Select Case i
        Case 1: SecOrder = 1
        Case 2: SecOrder = 2
        Case 3 To 8: SecOrder = i + 1
        Case 9: SecOrder = 10
        Case 10: SecOrder = 11
        Case 11: SecOrder = 12
        Case 12: SecOrder = 14
        Case 13: SecOrder = 13
        Case Else: SecOrder = 3
    End Select
End Function

' DETAIL section for an MRI Financial Format group (0 = no DETAIL section for that group)
Private Function GroupSec(ByVal grp As String) As Long
    Select Case grp
        Case "INCOME": GroupSec = 1
        Case "TENANT IMP. & COMM": GroupSec = 3
        Case "UTILITIES": GroupSec = 4
        Case "JANITORIAL & CLEANING": GroupSec = 5
        Case "ELEVATOR": GroupSec = 6
        Case "HVAC & PAYROLL": GroupSec = 7
        Case "LANDSCAPING & GROUNDS": GroupSec = 8
        Case "GENERAL R&M": GroupSec = 9
        Case "TAXES": GroupSec = 10
        Case "INSURANCE": GroupSec = 11
        Case "PROFESSIONAL FEES/G&A": GroupSec = 12
        Case "DEBT SERVICE": GroupSec = 13
        Case "OTHER PROFESSIONAL FEES": GroupSec = 14
        Case Else: GroupSec = 0
    End Select
End Function

' DETAIL section for an account found on LEASE_CALC / OTHER_REVENUE. 0 = ignored: the account is not on the
' MRI Financial Format, or its format group has no DETAIL section yet (flagged on LEASE_CALC column V).
' Public so PROFORMA can classify prior-year-only accounts; CollectAccounts must run first (it loads the chart of accounts).
Public Function RevSec(ByVal a As Double) As Long
    Dim k As Long, j As Long, grp As String
    RevSec = 0
    For k = 24 To 28                                   ' leasing-cost accounts listed on INPUT
        If mwsI.Cells(k, 3).Value = a Then
            RevSec = 3
            Exit Function
        End If
    Next k
    For j = 1 To UBound(mvNo, 1)
        If VarType(mvNo(j, 1)) = vbDouble Then
            If mvNo(j, 1) = a Then
                grp = UCase(Trim(CStr(mvFG(j, 1))))
                If grp = "" Then Exit Function         ' not on the Financial Format
                If CStr(mvTy(j, 1)) = "INCOME" Then
                    RevSec = 1
                Else
                    RevSec = GroupSec(grp)
                End If
                Exit Function
            End If
        End If
    Next j
End Function

' Add every GL account with a non-zero annual amount on one source sheet to tmp()/tsec()
' (needTag <> "" = only lines with that tag in column E count). secFixed = DETAIL section for the sheet
' (0 = classified per account by RevSec; accounts classified 0 are ignored).
Private Sub AddFromSheet(ByVal wsName As String, ByVal firstRow As Long, ByVal lastRow As Long, _
                         ByVal needTag As String, ByVal secFixed As Long, ByVal totCol As String, _
                         ByRef tmp() As Double, ByRef tsec() As Long, ByRef n As Long)
    Dim ws As Worksheet, vD As Variant, vS As Variant, vE As Variant
    Dim i As Long, j As Long, a As Double, found As Boolean, ok As Boolean, sc As Long
    Set ws = ThisWorkbook.Worksheets(wsName)
    vD = ws.Range("D" & firstRow & ":D" & lastRow).Value
    vS = ws.Range(totCol & firstRow & ":" & totCol & lastRow).Value
    vE = ws.Range("E" & firstRow & ":E" & lastRow).Value
    For i = 1 To UBound(vD, 1)
        If VarType(vD(i, 1)) = vbDouble And VarType(vS(i, 1)) = vbDouble Then
            If vD(i, 1) >= ACCT_MIN And Abs(vS(i, 1)) > 0.004 Then
                ok = True
                If needTag <> "" Then ok = (CStr(vE(i, 1)) = needTag)
                If ok Then
                    a = vD(i, 1)
                    sc = secFixed
                    If secFixed = 0 Then sc = RevSec(a)
                    If sc > 0 Then
                        found = False
                        For j = 1 To n
                            If tmp(j) = a And tsec(j) = sc Then
                                found = True
                                Exit For
                            End If
                        Next j
                        If Not found And n < 800 Then
                            n = n + 1
                            tmp(n) = a
                            tsec(n) = sc
                        End If
                    End If
                End If
            End If
        End If
    Next i
End Sub

' Collect every GL account that has a non-zero annual amount on LEASE_CALC, on an OTHER_REVENUE
' input line, an EXP_ input line, TAXES, INSURANCE, PROF_FEES, DEBT, OTHER_PROF_FEES or REPAIRS_CAPITAL, sorted by account then section.
' sect(i): 1 = income, 3 = leasing costs (Imprvmts & commissions), 4-9 = the six expense sections,
' 10 = taxes, 11 = insurance, 12 = professional fees/G&A, 13 = debt service, 14 = other professional fees. An account is classified by its MRI Financial
' Format group; tenant-billed amounts on LEASE_CALC for an expense account land in that account's section
' as an offset. Accounts not on the Financial Format (or in a group with no section yet) are ignored.
Public Function CollectAccounts(ByRef accts() As Double, ByRef sect() As Long) As Long
    Dim wsC As Worksheet
    Dim i As Long, j As Long, n As Long, ts As Long
    Dim t As Double
    Dim tmp() As Double, tsec() As Long
    ReDim tmp(1 To 800)
    ReDim tsec(1 To 800)
    n = 0
    Set wsC = ThisWorkbook.Worksheets("DATA_COA")
    mvNo = wsC.Range("C2:C2000").Value
    mvTy = wsC.Range("E2:E2000").Value
    mvFG = wsC.Range("G2:G2000").Value
    Set mwsI = ThisWorkbook.Worksheets("INPUT")
    AddFromSheet REV_SHEET, REV_FIRST, REV_LAST, "", 0, "S", tmp, tsec, n
    AddFromSheet "OTHER_REVENUE", OR_FIRST, OR_LAST, "INPUT", 0, "S", tmp, tsec, n
    For i = 1 To 12
        AddFromSheet ExpSheetName(i), EXP_FIRST, EXP_LAST, ExpTag(i + 3), SheetSec(i), SheetTotCol(ExpSheetName(i)), tmp, tsec, n
    Next i
    AddFromSheet RC_SHEET, EXP_FIRST, EXP_LAST, "REPAIR", 0, SheetTotCol(RC_SHEET), tmp, tsec, n      ' major repairs: section by the account's Financial Format group

    ' insertion sort ascending by account, then section
    For i = 2 To n
        t = tmp(i)
        ts = tsec(i)
        j = i - 1
        Do While j >= 1
            If tmp(j) < t Then Exit Do
            If tmp(j) = t And tsec(j) <= ts Then Exit Do
            tmp(j + 1) = tmp(j)
            tsec(j + 1) = tsec(j)
            j = j - 1
        Loop
        tmp(j + 1) = t
        tsec(j + 1) = ts
    Next i

    If n = 0 Then
        ReDim accts(1 To 1)
        ReDim sect(1 To 1)
        CollectAccounts = 0
        Exit Function
    End If
    ReDim accts(1 To n)
    ReDim sect(1 To n)
    For i = 1 To n
        accts(i) = tmp(i)
        sect(i) = tsec(i)
    Next i
    CollectAccounts = n
End Function

' Reset a generated sheet: unprotect, clear, default font
Public Sub ResetSheet(ByVal ws As Worksheet, Optional ByVal fontName As String = "Arial")
    ws.Unprotect
    ws.Cells.Clear
    ws.Cells.Font.Name = fontName
    ws.Cells.Font.Size = 10
End Sub

' Protect a generated sheet (no password) - users can still resize columns/rows
Public Sub LockSheet(ByVal ws As Worksheet)
    On Error Resume Next
    ws.Protect DrawingObjects:=True, Contents:=True, Scenarios:=True, AllowFormattingColumns:=True, AllowFormattingRows:=True
    If Err.Number <> 0 Then
        Err.Clear
        ws.Protect
    End If
    On Error GoTo 0
End Sub

Public Sub BorderBottom(ByVal ws As Worksheet, ByVal r As Long, ByVal c1 As Long, ByVal c2 As Long)
    ws.Range(ws.Cells(r, c1), ws.Cells(r, c2)).Borders(9).LineStyle = 1
End Sub

Public Sub BorderTop(ByVal ws As Worksheet, ByVal r As Long, ByVal c1 As Long, ByVal c2 As Long)
    ws.Range(ws.Cells(r, c1), ws.Cells(r, c2)).Borders(8).LineStyle = 1
End Sub

' Section titles shared by DETAIL and QTRLY SUM
Public Function SectionTitle(ByVal sec As Long) As String
    Select Case sec
        Case 1: SectionTitle = "INCOME"
        Case 3: SectionTitle = "IMPRVMTS & COMMISSIONS"
        Case 4: SectionTitle = "UTILITIES"
        Case 5: SectionTitle = "JANITORIAL & CLEANING"
        Case 6: SectionTitle = "ELEVATOR"
        Case 7: SectionTitle = "HVAC & PAYROLL"
        Case 8: SectionTitle = "LANDSCAPING & GROUNDS"
        Case 9: SectionTitle = "GENERAL R&M"
        Case 10: SectionTitle = "TAXES"
        Case 11: SectionTitle = "INSURANCE"
        Case 12: SectionTitle = "PROFESSIONAL FEES/G&A"
        Case 13: SectionTitle = "DEBT SERVICE"
        Case 14: SectionTitle = "OTHER PROFESSIONAL FEES"
    End Select
End Function

Public Function SectionTotalTitle(ByVal sec As Long) As String
    Select Case sec
        Case 1: SectionTotalTitle = "TOTAL INCOME"
        Case 3: SectionTotalTitle = "TOTAL IMPRVMTS & COMMISSIONS"
        Case 4: SectionTotalTitle = "TOTAL UTILITIES"
        Case 5: SectionTotalTitle = "TOTAL JANITORIAL & CLEANING"
        Case 6: SectionTotalTitle = "TOTAL ELEVATOR"
        Case 7: SectionTotalTitle = "TOTAL HVAC & PAYROLL"
        Case 8: SectionTotalTitle = "TOTAL LANDSCAPING & GROUNDS"
        Case 9: SectionTotalTitle = "TOTAL GENERAL R&M"
        Case 10: SectionTotalTitle = "TOTAL TAXES"
        Case 11: SectionTotalTitle = "TOTAL INSURANCE"
        Case 12: SectionTotalTitle = "TOTAL PROFESSIONAL FEES/G&A"
        Case 13: SectionTotalTitle = "TOTAL DEBT SERVICE"
        Case 14: SectionTotalTitle = "TOTAL OTHER PROFESSIONAL FEES"
    End Select
End Function
