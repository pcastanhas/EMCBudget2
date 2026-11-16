Attribute VB_Name = "modBuildDetail"
Option Explicit
'==============================================================================
' BuildDetail - creates the DETAIL budget schedule from REVENUE, OTHER_REVENUE, the six EXP_ sheets, TAXES, INSURANCE, PROF_FEES, OTHER_PROF_FEES and DEBT.
' One row per GL account that has amounts, grouped into colour-coded sections,
' with live SUMIFS formulas (by GL account). Formatting comes from modFormat.
' Other sheets reference only DETAIL!B3 (status) and DETAIL!B4 (last built).
'==============================================================================
Public Sub BuildDetail()
    Dim ws As Worksheet
    Dim accts() As Double, sect() As Long, n As Long
    Dim r As Long, sec As Long, oi As Long, i As Long, c As Long, cnt As Long, k As Long
    Dim secTot(1 To 14) As Long, rTE As Long, rTOE As Long, rNOI As Long, rOI As Long, rNBC As Long, rNCF As Long, rDist As Long, rRes As Long, rNCA As Long, rMR As Long, rOC As Long, pl As String, ml As String
    Dim firstR As Long, lastR As Long, endSched As Long
    Dim dec As Boolean, pal As String, wM As Double

    Set ws = ThisWorkbook.Worksheets("DETAIL")
    ResetSheet ws, FMT_FONT
    dec = ShowDecimals()
    n = CollectAccounts(accts, sect)

    ' ---- column widths (wider when cents are shown)
    If dec Then
        wM = 14
    Else
        wM = 12
    End If
    ws.Columns("A").ColumnWidth = 10
    ws.Columns("B").ColumnWidth = 44
    ws.Columns("C").ColumnWidth = 3
    For c = 4 To 15
        ws.Columns(c).ColumnWidth = wM
    Next c
    ws.Columns("P").ColumnWidth = wM + 2
    ws.Columns("Q").ColumnWidth = 11
    For c = 18 To 21
        ws.Columns(c).ColumnWidth = wM + 1
    Next c

    ' ---- title block (rows 1-4) and column headings (row 5)
    ws.Range("A1").Formula = "=BuildingName"
    ws.Range("A2").Formula = "=""BUDGET FOR THE YEAR ""&BudgetYear"
    ws.Range("A1:A2").Font.Bold = True
    ws.Range("A1:A2").Font.Color = PalColor("TITLE")
    ws.Range("A3").Value = "Status"
    ws.Range("A4").Value = "Last built"
    ws.Range("B4").Value = Now
    ws.Range("B4").NumberFormat = "yyyy-mm-dd hh:mm"
    ws.Range("B4").HorizontalAlignment = -4131
    ws.Range("A5").Value = "ACCT"
    ws.Range("B5").Value = "ACCOUNT NAME"
    For c = 4 To 15
        ws.Cells(5, c).NumberFormat = "mmm-yy"     ' format first (LibreOffice-safe), then formula
        ws.Cells(5, c).Formula = "=DATE(BudgetYear," & (c - 3) & ",1)"
    Next c
    ws.Range("P5").Value = "TOTAL"
    ws.Range("Q5").Value = "$/SF"
    ws.Range("R5").Value = "1ST QTR"
    ws.Range("S5").Value = "2ND QTR"
    ws.Range("T5").Value = "3RD QTR"
    ws.Range("U5").Value = "4TH QTR"
    PaintRow ws, 5, 1, 21, "COLHDR", True, True
    ws.Range("D5:U5").HorizontalAlignment = -4108

    ' ---- sections
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
            PaintRow ws, r, 1, 21, pal & "_HDR", True, True
            r = r + 1
            firstR = r
            k = 0
            For i = 1 To n
                If sect(i) = sec Then
                    DetailAccountRow ws, r, accts(i), dec, sec
                    If k Mod 2 = 0 Then
                        PaintRow ws, r, 1, 21, pal & "_A", False, False
                    Else
                        PaintRow ws, r, 1, 21, pal & "_B", False, False
                    End If
                    k = k + 1
                    r = r + 1
                End If
            Next i
            lastR = r - 1
            ws.Cells(r, 2).Value = SectionTotalTitle(sec)
            For c = 4 To 15
                ws.Cells(r, c).Formula = "=SUM(" & ColL(c) & firstR & ":" & ColL(c) & lastR & ")"
                ws.Cells(r, c).NumberFormat = MoneyFormat(dec)
            Next c
            DetailTotalsCols ws, r, dec
            PaintRow ws, r, 1, 21, pal & "_HDR", True, True
            secTot(sec) = r
            r = r + 2
        End If
        If sec = 9 Then                                ' after the last expense section: TOTAL EXPENSES
            pl = ""
            For k = 4 To 9
                If secTot(k) > 0 Then pl = pl & IIf(pl = "", "", ",") & secTot(k)
            Next k
            DetailSummaryRow ws, r, 2, "TOTAL EXPENSES", pl, "", "EXP_HDR", True, dec
            rTE = r
            r = r + 2
        End If
        If sec = 12 Then                               ' after professional fees: TOTAL OPERATING EXPENSES and N.O.I.
            pl = CStr(rTE)
            For k = 10 To 12
                If secTot(k) > 0 Then pl = pl & "," & secTot(k)
            Next k
            DetailSummaryRow ws, r, 1, "TOTAL OPERATING EXPENSES", pl, "", "GRAND_HDR", True, dec
            rTOE = r
            r = r + 2
            If secTot(1) > 0 Then pl = CStr(secTot(1)) Else pl = ""
            DetailSummaryRow ws, r, 1, "N.O.I. BEFORE OTHER PROFESSIONAL FEES & DEBT SERVICE", pl, CStr(rTOE), "NOI_HDR", False, dec
            rNOI = r
            r = r + 2
        End If
        If sec = 14 Then                               ' after other professional fees: OPERATING INCOME BEFORE DEBT SERVICE (N.O.I. less other professional fees)
            If secTot(14) > 0 Then ml = CStr(secTot(14)) Else ml = ""
            DetailSummaryRow ws, r, 1, "OPERATING INCOME BEFORE DEBT SERVICE", CStr(rNOI), ml, "NOI_HDR", False, dec
            rOI = r
            r = r + 2
        End If
        If sec = 13 Then                               ' after debt service: N.O.I. BEFORE CAPITAL (operating income before debt service less total debt service)
            If secTot(13) > 0 Then ml = CStr(secTot(13)) Else ml = ""
            DetailSummaryRow ws, r, 1, "N.O.I. BEFORE CAPITAL", CStr(rOI), ml, "NOI_HDR", False, dec
            rNBC = r
            r = r + 2
        End If
    Next oi
    ' ---- NET CASH FLOW (N.O.I. before capital less improvements & commissions), distributions and reserve additions (INPUT), then the reserve, operating cash and total cash balances
    If secTot(3) > 0 Then ml = CStr(secTot(3)) Else ml = ""
    DetailSummaryRow ws, r, 1, "NET CASH FLOW", CStr(rNBC), ml, "NOI_HDR", False, dec
    rNCF = r
    r = r + 2
    DetailFlowRow ws, r, "Less: partner distributions", "=IF(DistOn=""Yes"",SUMIFS(DistAmt,DistMonth,#),0)", dec
    rDist = r
    r = r + 1
    DetailFlowRow ws, r, "Less: additions to maintenance reserve", "=SUMIFS(ResAmt,ResMonth,#)", dec
    rRes = r
    r = r + 2
    DetailSummaryRow ws, r, 1, "NET CASH FLOW AFTER DISTRIBUTIONS & RESERVES", CStr(rNCF), rDist & "," & rRes, "NOI_HDR", False, dec
    rNCA = r
    r = r + 2
    DetailBalanceRow ws, r, "MAINTENANCE RESERVE BALANCE", "ResStart", rRes, dec
    rMR = r
    r = r + 1
    DetailBalanceRow ws, r, "PROJECTED OPERATING CASH BALANCE", "StartCash-ResStart", rNCA, dec
    rOC = r
    r = r + 2
    DetailTotalCashRow ws, r, rOC, rMR, dec
    r = r + 2
    If n = 0 Then
        ws.Cells(r, 1).Value = "No amounts found on REVENUE yet."
        r = r + 2
    End If
    endSched = r - 2
    If endSched < 6 Then endSched = 6

    ' ---- status (fixed cell B3 - referenced by INPUT); the checks live inside this formula
    ws.Range("B3").Formula = DetailStatusFormula(endSched, dec)
    ws.Range("B3").Font.Bold = True
    ws.Range("B3").HorizontalAlignment = -4131
    ws.Cells(r, 1).Value = "This sheet is generated by the BuildDetail macro - do not type on it. Change amounts on REVENUE / INPUT, then rebuild if accounts change."
    ws.Cells(r, 1).Font.Italic = True
    LockSheet ws
End Sub

' Status text for DETAIL!B3. Reports the first problem found, in this order:
'   1 the "show decimals" option changed since the build           -> rebuild
'   2 schedule rows no longer match REVENUE / OTHER_REVENUE        -> rebuild
'   3-16 data problems upstream of the schedules
Private Function DetailStatusFormula(ByVal endSched As Long, ByVal dec As Boolean) As String
    Dim sS As String, sD As String, sE As String, REVT As String
    Dim oS As String, oD As String, oE As String
    Dim conds(1 To 16) As String, msgs(1 To 16) As String
    Dim expT As String, j As Long
    Dim i As Long, f As String
    sS = "REVENUE!$S$" & REV_FIRST & ":$S$" & REV_LAST
    sD = "REVENUE!$D$" & REV_FIRST & ":$D$" & REV_LAST
    sE = "REVENUE!$E$" & REV_FIRST & ":$E$" & REV_LAST
    REVT = "REVENUE!$T$" & REV_FIRST & ":$T$" & REV_LAST
    oS = "OTHER_REVENUE!$S$" & OR_FIRST & ":$S$" & OR_LAST
    oD = "OTHER_REVENUE!$D$" & OR_FIRST & ":$D$" & OR_LAST
    oE = "OTHER_REVENUE!$E$" & OR_FIRST & ":$E$" & OR_LAST
    If dec Then
        conds(1) = "INPUT!$C$6<>""Yes"""
    Else
        conds(1) = "INPUT!$C$6=""Yes"""
    End If
    msgs(1) = "REBUILD NEEDED - display option changed. Click Build Schedules"
    conds(2) = "ABS(" & AcctTotalTerm() & "-SUMIFS($P$6:$P$" & endSched & ",$A$6:$A$" & endSched & ","">=1000000""))>=0.5"
    msgs(2) = "REBUILD NEEDED - click Build Schedules"
    conds(3) = "COUNTIF(" & sD & ",""CODE?"")+COUNTIF(TENANT_SCHEDULE!$K$2:$K$401,0)>0"
    msgs(3) = "CHECK DATA - charge code not in the MRI charge-code list"
    conds(4) = "ABS(SUMIFS(" & sS & "," & sD & ",""TOTAL"")-(SUMIFS(" & sS & "," & sD & ","">=1000000"")-SUMIFS(" & sS & "," & sE & ",""COST"")))>=0.5"
    msgs(4) = "CHECK DATA - amount on a REVENUE line with no GL account"
    conds(5) = "ABS(SUMIFS(" & oS & "," & oE & ",""INPUT"")-SUMIFS(" & oS & "," & oD & ","">=1000000""," & oE & ",""INPUT""))>=0.5"
    msgs(5) = "CHECK DATA - OTHER_REVENUE amount on a line with no GL account"
    conds(6) = "MAX(0,MAX(IA_RANK)-COUNT(OTHER_REVENUE!$U$" & OR_FIRST & ":$U$" & OR_LAST & "))>0"
    msgs(6) = "CHECK DATA - more income accounts than OTHER_REVENUE blocks"
    conds(7) = "MAX(0,MAX(TS_RANK)-COUNT(REVENUE!$U$" & REV_FIRST & ":$U$" & REV_LAST & "))>0"
    msgs(7) = "CHECK DATA - more MRI leases than lease blocks"
    conds(8) = "OR(COUNTIF(" & REVT & ",""MORE CODES*"")>0,SUMPRODUCT((DATA_TENANT_CHARGES!$K$2:$K$401<>"""")*(DATA_TENANT_CHARGES!$K$2:$K$401<>""M""))>0)"
    msgs(8) = "CHECK DATA - a lease has more MRI charge codes than lines (rebuild the workbook), or an MRI charge is not monthly"
    conds(9) = "COUNTIF(OTHER_REVENUE!$T$" & OR_FIRST & ":$T$" & OR_LAST & ",""DUPLICATE*"")+COUNTIF(OTHER_REVENUE!$T$" & OR_FIRST & ":$T$" & OR_LAST & ",""NOT AN INCOME*"")" & _
        "+COUNTIF(OTHER_PROF_FEES!$T$" & EXP_FIRST & ":$T$" & EXP_LAST & ",""DUPLICATE*"")+COUNTIF(OTHER_PROF_FEES!$T$" & EXP_FIRST & ":$T$" & EXP_LAST & ",""NOT AN OTHER*"")>0"
    msgs(9) = "CHECK DATA - a User block on OTHER_REVENUE or OTHER_PROF_FEES has a duplicate or invalid account"
    conds(10) = "SUM(EA_NOBLOCK)>0"
    msgs(10) = "CHECK DATA - an expense or debt service account has activity but no block on the input sheets (rebuild the workbook)"
    expT = ""
    For j = 1 To 9
        If j <> 7 Then
            If expT <> "" Then expT = expT & "+"
            expT = expT & "COUNTIF(" & ExpSheetName(j) & "!$T$" & EXP_FIRST & ":$T$" & EXP_LAST & ",""NEW VENDOR*"")"
        End If
    Next j
    conds(11) = expT & ">0"
    msgs(11) = "CHECK DATA - a vendor in the vendor extract has no line on the EXP_ sheets (rebuild the workbook)"
    conds(12) = "OR(COUNT(TAXES!$G$12:$G$15)<4,MIN(TAXES!$G$12:$G$15)<1,MAX(TAXES!$G$12:$G$15)>12,AND(COUNT(TAX_MAIN_LIST)>0,COUNTIF(TAX_MAIN_LIST,TaxMain)=0))"
    msgs(12) = "CHECK DATA - a tax payment month on TAXES is missing or not 1 to 12, or the main real estate tax account on INPUT is not a TAXES account"
    conds(13) = "COUNTIFS(REVENUE!$V$" & REV_FIRST & ":$V$" & REV_LAST & ",""IGNORED - NOT ON*"",REVENUE!$S$" & REV_FIRST & ":$S$" & REV_LAST & ",""<>0"")>0"
    msgs(13) = "CHECK DATA - a REVENUE charge posts to an account that is not on the MRI Financial Format (ignored) - see REVENUE column V"
    conds(14) = "COUNTIF(" & RC_SHEET & "!$AO$" & EXP_FIRST & ":$AO$" & EXP_LAST & ",""CHECK*"")>0"
    msgs(14) = "CHECK DATA - a line on REPAIRS_CAPITAL needs attention (account, start month or monthly amounts) - see the Line check column"
    conds(15) = "SUMPRODUCT((DistAmt<>0)*((DistMonth<1)+(DistMonth>12)+(DistMonth<>INT(DistMonth))))+SUMPRODUCT((ResAmt<>0)*((ResMonth<1)+(ResMonth>12)+(ResMonth<>INT(ResMonth))))>0"
    msgs(15) = "CHECK DATA - a distribution or reserve addition on INPUT has no valid month (1 to 12)"
    conds(16) = "COUNTIF(DEAL_CHECK,""CHECK*"")>0"
    msgs(16) = "CHECK DATA - a line on LEASE_DEALS needs attention - see its Line status column"
    f = """OK - schedules current"""
    For i = 16 To 1 Step -1
        f = "IF(" & conds(i) & ",""" & msgs(i) & """," & f & ")"
    Next i
    DetailStatusFormula = "=" & f
End Function

' Plain monthly line (no band) whose month formulas come from a template; "#" is replaced by the month number 1-12.
Private Sub DetailFlowRow(ByVal ws As Worksheet, ByVal r As Long, ByVal lbl As String, ByVal tmpl As String, ByVal dec As Boolean)
    Dim c As Long
    ws.Cells(r, 2).Value = lbl
    For c = 4 To 15
        ws.Cells(r, c).Formula = Replace(tmpl, "#", CStr(c - 3))
        ws.Cells(r, c).NumberFormat = MoneyFormat(dec)
    Next c
    DetailTotalsCols ws, r, dec
End Sub

' Running balance: start expression (INPUT names) plus the cumulative flow row. Total and quarter columns show the balance at the period end.
Private Sub DetailBalanceRow(ByVal ws As Worksheet, ByVal r As Long, ByVal lbl As String, ByVal startExpr As String, _
                             ByVal flowRow As Long, ByVal dec As Boolean)
    Dim c As Long
    ws.Cells(r, 2).Value = lbl
    ws.Cells(r, 4).Formula = "=" & startExpr & "+D" & flowRow
    For c = 5 To 15
        ws.Cells(r, c).Formula = "=" & ColL(c - 1) & r & "+" & ColL(c) & flowRow
    Next c
    BalanceEndCols ws, r, dec
    ws.Range(ws.Cells(r, 2), ws.Cells(r, 21)).Font.Bold = True
End Sub

' PROJECTED CASH BALANCE = operating cash + maintenance reserve (reserve additions are transfers, so total cash only moves with net cash flow and distributions)
Private Sub DetailTotalCashRow(ByVal ws As Worksheet, ByVal r As Long, ByVal rOC As Long, ByVal rMR As Long, ByVal dec As Boolean)
    Dim c As Long
    ws.Cells(r, 1).Value = "PROJECTED CASH BALANCE"
    For c = 4 To 15
        ws.Cells(r, c).Formula = "=" & ColL(c) & rOC & "+" & ColL(c) & rMR
    Next c
    BalanceEndCols ws, r, dec
    PaintRow ws, r, 1, 21, "GRAND_HDR", True, True
End Sub

' Month columns D:O are set by the caller; this sets the number format and the period-end total / quarter columns for a balance row
Private Sub BalanceEndCols(ByVal ws As Worksheet, ByVal r As Long, ByVal dec As Boolean)
    ws.Cells(r, 16).Formula = "=O" & r
    ws.Cells(r, 18).Formula = "=F" & r
    ws.Cells(r, 19).Formula = "=I" & r
    ws.Cells(r, 20).Formula = "=L" & r
    ws.Cells(r, 21).Formula = "=O" & r
    ws.Range(ws.Cells(r, 4), ws.Cells(r, 16)).NumberFormat = MoneyFormat(dec)
    ws.Range(ws.Cells(r, 18), ws.Cells(r, 21)).NumberFormat = MoneyFormat(dec)
End Sub

Private Sub DetailAccountRow(ByVal ws As Worksheet, ByVal r As Long, ByVal acct As Double, ByVal dec As Boolean, ByVal sec As Long)
    Dim c As Long, rc As String
    ws.Cells(r, 1).Value = acct
    ws.Cells(r, 1).NumberFormat = "0000\-000"
    ws.Cells(r, 1).HorizontalAlignment = -4131
    ws.Cells(r, 2).Formula = "=IFERROR(INDEX(COA_NAME,MATCH($A" & r & ",COA_ACCTNO,0)),""(not in chart of accounts)"")"
    For c = 4 To 15
        rc = ColL(c + 3)                       ' DETAIL col D <-> REVENUE / OTHER_REVENUE col G
        ws.Cells(r, c).Formula = "=" & AcctMonthTerm(sec, rc, "$A" & r)
        ws.Cells(r, c).NumberFormat = MoneyFormat(dec)
    Next c
    DetailTotalsCols ws, r, dec
End Sub

' Summary line on DETAIL: months = sum of the plusRows (comma-separated row numbers) less the minusRows.
' lblCol = column that holds the label (1 = A, 2 = B). Painted as a full-width band.
Private Sub DetailSummaryRow(ByVal ws As Worksheet, ByVal r As Long, ByVal lblCol As Long, ByVal lbl As String, _
                             ByVal plusRows As String, ByVal minusRows As String, ByVal palKey As String, _
                             ByVal whiteText As Boolean, ByVal dec As Boolean)
    Dim c As Long, j As Long, f As String, p() As String
    ws.Cells(r, lblCol).Value = lbl
    For c = 4 To 15
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
        If f = "" Then
            f = "=0"
        Else
            f = "=" & f
        End If
        ws.Cells(r, c).Formula = f
        ws.Cells(r, c).NumberFormat = MoneyFormat(dec)
    Next c
    DetailTotalsCols ws, r, dec
    PaintRow ws, r, 1, 21, palKey, whiteText, True
End Sub

' TOTAL, $/SF and quarter columns for a row (months D:O already in place)
Private Sub DetailTotalsCols(ByVal ws As Worksheet, ByVal r As Long, ByVal dec As Boolean)
    ws.Cells(r, 16).Formula = "=SUM(D" & r & ":O" & r & ")"
    ws.Cells(r, 16).NumberFormat = MoneyFormat(dec)
    ws.Cells(r, 17).Formula = "=IF(BillingSF=0,0,P" & r & "/BillingSF)"
    ws.Cells(r, 17).NumberFormat = MoneyFormat(True)           ' $/SF always shows cents
    ws.Cells(r, 18).Formula = "=SUM(D" & r & ":F" & r & ")"
    ws.Cells(r, 19).Formula = "=SUM(G" & r & ":I" & r & ")"
    ws.Cells(r, 20).Formula = "=SUM(J" & r & ":L" & r & ")"
    ws.Cells(r, 21).Formula = "=SUM(M" & r & ":O" & r & ")"
    ws.Range(ws.Cells(r, 18), ws.Cells(r, 21)).NumberFormat = MoneyFormat(dec)
End Sub
