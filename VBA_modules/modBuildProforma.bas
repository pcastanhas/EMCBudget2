Attribute VB_Name = "modBuildProforma"
Option Explicit
'==============================================================================
' BuildProforma - creates the PROFORMA schedule: the budget year compared with the prior-year forecast
' (MRI actuals through the actuals-through month + prior-year budget for the remaining months) and the
' prior-year full-year budget, by GL account. Same sections, order and summary lines as DETAIL down to
' NET CASH FLOW (no distribution, reserve or cash-balance lines).
' Built from the source sheets and the MRI data tabs, never from DETAIL rows (DETAIL!B3 is the only DETAIL cell read).
' Accounts: every DETAIL account plus accounts with prior-year activity but no budget-year amount
' (placed in the section of their MRI Financial Format group), so the prior-year columns are complete.
' Signs follow DETAIL: income, expenses and costs positive; concessions negative.
' Column F variance: + = favourable (income: forecast - budget; expenses, debt and costs: budget - forecast).
' Columns G-J: plain change of the budget year against the prior year (+ = increase).
'==============================================================================
Private Const PF_SHEET As String = "PROFORMA"
Private Const PF_LASTCOL As Long = 10
Private Const PCTFMT As String = "0.0%;(0.0%);""-"""

Public Sub BuildProforma()
    Dim ws As Worksheet
    Dim accts() As Double, sect() As Long, n As Long
    Dim r As Long, sec As Long, oi As Long, i As Long, c As Long, cnt As Long, k As Long
    Dim secTot(1 To 14) As Long, rTE As Long, rTOE As Long, rNOI As Long, rOI As Long, rNBC As Long, pl As String, ml As String
    Dim firstR As Long, lastR As Long, endSched As Long
    Dim dec As Boolean, pal As String, wM As Double

    Set ws = ThisWorkbook.Worksheets(PF_SHEET)
    ResetSheet ws, FMT_FONT
    dec = ShowDecimals()
    n = CollectAccounts(accts, sect)
    n = AddPriorYearAccounts(accts, sect, n)

    ' ---- column widths
    If dec Then
        wM = 15
    Else
        wM = 13
    End If
    ws.Columns("A").ColumnWidth = 10
    ws.Columns("B").ColumnWidth = 44
    For c = 3 To 6
        ws.Columns(c).ColumnWidth = wM
    Next c
    For c = 7 To 10
        ws.Columns(c).ColumnWidth = 11
    Next c

    ' ---- title block (rows 1-4) and column headings (row 5)
    ws.Range("A1").Formula = "=BuildingName&"" - PROFORMA"""
    ws.Range("A2").Formula = "=""BUDGET ""&BudgetYear&"" COMPARED WITH ""&(BudgetYear-1)"
    ws.Range("A1:A2").Font.Bold = True
    ws.Range("A1:A2").Font.Color = PalColor("TITLE")
    ws.Range("A3").Value = "Status"
    ws.Range("A4").Value = "Last built"
    ws.Range("B4").Value = Now
    ws.Range("B4").NumberFormat = "yyyy-mm-dd hh:mm"
    ws.Range("B4").HorizontalAlignment = -4131
    ws.Range("A5").Value = "ACCT"
    ws.Range("B5").Value = "ACCOUNT NAME"
    ws.Range("C5").Formula = "=BudgetYear&"" BUDGET"""
    ws.Range("D5").Formula = "=(BudgetYear-1)&"" FORECAST (ACTUAL JAN-""&UPPER(TEXT(DATE(BudgetYear-1,ActualsThru,1),""mmm""))&"" + BUDGET)"""
    ws.Range("E5").Formula = "=(BudgetYear-1)&"" FULL-YEAR BUDGET"""
    ws.Range("F5").Formula = "=(BudgetYear-1)&"" FORECAST VS BUDGET VARIANCE (+ = FAVORABLE)"""
    ws.Range("G5").Formula = "=""% CHANGE ""&BudgetYear&"" VS ""&(BudgetYear-1)&"" FORECAST"""
    ws.Range("H5").Formula = "=""$/SF CHANGE ""&BudgetYear&"" VS ""&(BudgetYear-1)&"" FORECAST"""
    ws.Range("I5").Formula = "=""% CHANGE ""&BudgetYear&"" VS ""&(BudgetYear-1)&"" BUDGET"""
    ws.Range("J5").Formula = "=""$/SF CHANGE ""&BudgetYear&"" VS ""&(BudgetYear-1)&"" BUDGET"""
    PaintRow ws, 5, 1, PF_LASTCOL, "COLHDR", True, True
    ws.Range("C5:J5").HorizontalAlignment = -4108
    ws.Range("A5:J5").WrapText = True
    ws.Range("A5:J5").VerticalAlignment = -4107
    ws.Rows(5).RowHeight = 54

    ' ---- sections (same order and summary lines as DETAIL)
    r = 7
    For oi = 1 To 14
        sec = SecOrder(oi)
        cnt = 0
        For i = 1 To n
            If sect(i) = sec Then cnt = cnt + 1
        Next i
        If cnt > 0 Then
            pal = SecPalette(sec)
            ws.Cells(r, 1).Value = SectionTitle(sec)
            PaintRow ws, r, 1, PF_LASTCOL, pal & "_HDR", True, True
            r = r + 1
            firstR = r
            k = 0
            For i = 1 To n
                If sect(i) = sec Then
                    ProAccountRow ws, r, accts(i), sec, dec
                    If k Mod 2 = 0 Then
                        PaintRow ws, r, 1, PF_LASTCOL, pal & "_A", False, False
                    Else
                        PaintRow ws, r, 1, PF_LASTCOL, pal & "_B", False, False
                    End If
                    k = k + 1
                    r = r + 1
                End If
            Next i
            lastR = r - 1
            ws.Cells(r, 2).Value = SectionTotalTitle(sec)
            For c = 3 To 5
                ws.Cells(r, c).Formula = "=SUM(" & ColL(c) & firstR & ":" & ColL(c) & lastR & ")"
            Next c
            ProCalcCols ws, r, (sec = 1), dec
            PaintRow ws, r, 1, PF_LASTCOL, pal & "_HDR", True, True
            secTot(sec) = r
            r = r + 2
        End If
        If sec = 9 Then
            pl = ""
            For k = 4 To 9
                If secTot(k) > 0 Then pl = pl & IIf(pl = "", "", ",") & secTot(k)
            Next k
            ProSummaryRow ws, r, 2, "TOTAL EXPENSES", pl, "", "EXP_HDR", True, False, dec
            rTE = r
            r = r + 2
        End If
        If sec = 12 Then
            pl = CStr(rTE)
            For k = 10 To 12
                If secTot(k) > 0 Then pl = pl & "," & secTot(k)
            Next k
            ProSummaryRow ws, r, 1, "TOTAL OPERATING EXPENSES", pl, "", "GRAND_HDR", True, False, dec
            rTOE = r
            r = r + 2
            If secTot(1) > 0 Then pl = CStr(secTot(1)) Else pl = ""
            ProSummaryRow ws, r, 1, "N.O.I. BEFORE OTHER PROFESSIONAL FEES & DEBT SERVICE", pl, CStr(rTOE), "NOI_HDR", False, True, dec
            rNOI = r
            r = r + 2
        End If
        If sec = 14 Then
            If secTot(14) > 0 Then ml = CStr(secTot(14)) Else ml = ""
            ProSummaryRow ws, r, 1, "OPERATING INCOME BEFORE DEBT SERVICE", CStr(rNOI), ml, "NOI_HDR", False, True, dec
            rOI = r
            r = r + 2
        End If
        If sec = 13 Then
            If secTot(13) > 0 Then ml = CStr(secTot(13)) Else ml = ""
            ProSummaryRow ws, r, 1, "N.O.I. BEFORE CAPITAL", CStr(rOI), ml, "NOI_HDR", False, True, dec
            rNBC = r
            r = r + 2
        End If
    Next oi
    If secTot(3) > 0 Then ml = CStr(secTot(3)) Else ml = ""
    ProSummaryRow ws, r, 1, "NET CASH FLOW", CStr(rNBC), ml, "NOI_HDR", False, True, dec
    r = r + 2
    If n = 0 Then
        ws.Cells(r, 1).Value = "No amounts found yet."
        r = r + 2
    End If
    endSched = r - 2
    If endSched < 6 Then endSched = 6

    ' ---- status (fixed cell B3): rows must still match the source sheets; otherwise the DETAIL status
    ws.Range("B3").Formula = "=IF(ABS(" & AcctTotalTerm() & "-SUMIFS($C$6:$C$" & endSched & ",$A$6:$A$" & endSched & ","">=1000000""))>=0.5,""REBUILD NEEDED - click Build Schedules"",DETAIL!$B$3)"
    ws.Range("B3").Font.Bold = True
    ws.Range("B3").HorizontalAlignment = -4131
    ws.Cells(r, 1).Value = "Prior-year forecast = MRI actuals from January through the actuals-through month on INPUT plus the prior-year budget for the remaining months. " & _
        "Variance: + = favorable (income above budget, costs below budget). % and $/SF columns: budget year less prior year (+ = increase); n/a = no prior-year amount."
    ws.Cells(r, 1).Font.Italic = True
    ws.Cells(r + 1, 1).Value = "This sheet is generated by the BuildProforma macro - do not type on it. Change amounts on the input sheets, then click Build Schedules if accounts change."
    ws.Cells(r + 1, 1).Font.Italic = True
    LockSheet ws
End Sub

' Add accounts with prior-year actual or budget activity that DETAIL does not list (no budget-year amount),
' classified by their Financial Format group (accounts with no DETAIL section are skipped). Returns the new count; keeps the sort by account, then section.
Private Function AddPriorYearAccounts(ByRef accts() As Double, ByRef sect() As Long, ByVal n As Long) As Long
    Dim py As Long, s As Long, i As Long, j As Long, c As Long, sc As Long, ts As Long
    Dim v As Variant, a As Double, tot As Double, t As Double, known As Boolean
    Dim srcs(1 To 2) As String, acol(1 To 2) As Long, m1(1 To 2) As Long
    py = CLng(ThisWorkbook.Worksheets("INPUT").Range("C3").Value) - 1
    srcs(1) = "DATA_ACTUALS": acol(1) = 3: m1(1) = 4          ' Year in B, account in C, months D:O
    srcs(2) = "DATA_BUDGET_PY": acol(2) = 4: m1(2) = 5        ' Year in B, account in D, months E:P
    For s = 1 To 2
        v = ThisWorkbook.Worksheets(srcs(s)).Range("A2:P2000").Value
        For i = 1 To UBound(v, 1)
            If IsNumeric(v(i, 2)) And Not IsEmpty(v(i, 2)) Then
                If CLng(v(i, 2)) = py And Left(CStr(v(i, acol(s))), 2) = "MR" And IsNumeric(Mid(CStr(v(i, acol(s))), 3)) Then
                    a = CDbl(Mid(CStr(v(i, acol(s))), 3))
                    tot = 0
                    For c = m1(s) To m1(s) + 11
                        If IsNumeric(v(i, c)) And Not IsEmpty(v(i, c)) Then tot = tot + Abs(CDbl(v(i, c)))
                    Next c
                    If tot > 0.004 Then
                        known = False
                        For j = 1 To n
                            If accts(j) = a Then
                                known = True
                                Exit For
                            End If
                        Next j
                        If Not known Then
                            sc = RevSec(a)
                            If sc > 0 Then
                                n = n + 1
                                If n = 1 Then
                                    ReDim accts(1 To 1)
                                    ReDim sect(1 To 1)
                                Else
                                    ReDim Preserve accts(1 To n)
                                    ReDim Preserve sect(1 To n)
                                End If
                                accts(n) = a
                                sect(n) = sc
                            End If
                        End If
                    End If
                End If
            End If
        Next i
    Next s
    ' insertion sort ascending by account, then section (same order as DETAIL)
    For i = 2 To n
        t = accts(i)
        ts = sect(i)
        j = i - 1
        Do While j >= 1
            If accts(j) < t Then Exit Do
            If accts(j) = t And sect(j) <= ts Then Exit Do
            accts(j + 1) = accts(j)
            sect(j + 1) = sect(j)
            j = j - 1
        Loop
        accts(j + 1) = t
        sect(j + 1) = ts
    Next i
    AddPriorYearAccounts = n
End Function

' Prior-year forecast for one account: MRI actuals Jan..ActualsThru + prior-year budget for the later months (MRI signs)
Private Function PyForecastTerm(ByVal acctRef As String) As String
    PyForecastTerm = "SUMPRODUCT((DATA_ACTUALS!$C$2:$C$2000=""MR""&" & acctRef & ")*(DATA_ACTUALS!$B$2:$B$2000=BudgetYear-1)*(DATA_ACTUALS!$D$2:$O$2000)*((COLUMN(DATA_ACTUALS!$D$2:$O$2000)-3)<=ActualsThru))" & _
        "+SUMPRODUCT((DATA_BUDGET_PY!$D$2:$D$2000=""MR""&" & acctRef & ")*(DATA_BUDGET_PY!$B$2:$B$2000=BudgetYear-1)*(DATA_BUDGET_PY!$E$2:$P$2000)*((COLUMN(DATA_BUDGET_PY!$E$2:$P$2000)-4)>ActualsThru))"
End Function

' Prior-year full-year budget for one account (MRI signs)
Private Function PyBudgetTerm(ByVal acctRef As String) As String
    PyBudgetTerm = "SUMPRODUCT((DATA_BUDGET_PY!$D$2:$D$2000=""MR""&" & acctRef & ")*(DATA_BUDGET_PY!$B$2:$B$2000=BudgetYear-1)*(DATA_BUDGET_PY!$E$2:$P$2000))"
End Function

Private Sub ProAccountRow(ByVal ws As Worksheet, ByVal r As Long, ByVal acct As Double, ByVal sec As Long, ByVal dec As Boolean)
    Dim aRef As String, sg As String
    aRef = "$A" & r
    If sec = 1 Then sg = "-" Else sg = ""          ' MRI stores income as credits (negative); DETAIL shows income positive
    ws.Cells(r, 1).Value = acct
    ws.Cells(r, 1).NumberFormat = "0000\-000"
    ws.Cells(r, 1).HorizontalAlignment = -4131
    ws.Cells(r, 2).Formula = "=IFERROR(INDEX(COA_NAME,MATCH(" & aRef & ",COA_ACCTNO,0)),""(not in chart of accounts)"")"
    ws.Cells(r, 3).Formula = "=" & AcctMonthTerm(sec, "S", aRef)          ' column S = annual total on every source sheet (AN on REPAIRS_CAPITAL)
    ws.Cells(r, 4).Formula = "=" & sg & "(" & PyForecastTerm(aRef) & ")"
    ws.Cells(r, 5).Formula = "=" & sg & "(" & PyBudgetTerm(aRef) & ")"
    ProCalcCols ws, r, (sec = 1), dec
End Sub

' Summary line: columns C-E = sum of the plusRows less the minusRows (comma-separated row numbers), then the variance columns
Private Sub ProSummaryRow(ByVal ws As Worksheet, ByVal r As Long, ByVal lblCol As Long, ByVal lbl As String, _
                          ByVal plusRows As String, ByVal minusRows As String, ByVal palKey As String, _
                          ByVal whiteText As Boolean, ByVal incomeLike As Boolean, ByVal dec As Boolean)
    Dim c As Long, j As Long, f As String, p() As String
    ws.Cells(r, lblCol).Value = lbl
    For c = 3 To 5
        f = ""
        If plusRows <> "" Then
            p = Split(plusRows, ",")
            For j = LBound(p) To UBound(p)
                f = f & "+" & ColL(c) & p(j)
            Next j
        End If
        If minusRows <> "" Then
            p = Split(minusRows, ",")
            For j = LBound(p) To UBound(p)
                f = f & "-" & ColL(c) & p(j)
            Next j
        End If
        If f = "" Then f = "0"
        ws.Cells(r, c).Formula = "=" & f
    Next c
    ProCalcCols ws, r, incomeLike, dec
    PaintRow ws, r, 1, PF_LASTCOL, palKey, whiteText, True
End Sub

' Columns F-J for a row whose C (budget year), D (prior-year forecast) and E (prior-year budget) are in place.
' incomeLike = True for income and profit lines (higher forecast than budget is favorable).
Private Sub ProCalcCols(ByVal ws As Worksheet, ByVal r As Long, ByVal incomeLike As Boolean, ByVal dec As Boolean)
    If incomeLike Then
        ws.Cells(r, 6).Formula = "=D" & r & "-E" & r
    Else
        ws.Cells(r, 6).Formula = "=E" & r & "-D" & r
    End If
    ws.Cells(r, 7).Formula = "=IF(D" & r & "=0,""n/a"",(C" & r & "-D" & r & ")/ABS(D" & r & "))"
    ws.Cells(r, 8).Formula = "=IF(BillingSF=0,0,(C" & r & "-D" & r & ")/BillingSF)"
    ws.Cells(r, 9).Formula = "=IF(E" & r & "=0,""n/a"",(C" & r & "-E" & r & ")/ABS(E" & r & "))"
    ws.Cells(r, 10).Formula = "=IF(BillingSF=0,0,(C" & r & "-E" & r & ")/BillingSF)"
    ws.Range(ws.Cells(r, 3), ws.Cells(r, 6)).NumberFormat = MoneyFormat(dec)
    ws.Cells(r, 7).NumberFormat = PCTFMT
    ws.Cells(r, 9).NumberFormat = PCTFMT
    ws.Cells(r, 8).NumberFormat = MoneyFormat(True)          ' $/SF always shows cents
    ws.Cells(r, 10).NumberFormat = MoneyFormat(True)
    ws.Cells(r, 7).HorizontalAlignment = -4152
    ws.Cells(r, 9).HorizontalAlignment = -4152
End Sub
