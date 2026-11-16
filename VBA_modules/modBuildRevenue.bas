Attribute VB_Name = "modBuildRevenue"
Option Explicit
'==============================================================================
' BuildRevenue - creates the purple REVENUE schedule (read-only presentation) by lease:
'   IN-PLACE LEASES   one block per MRI lease on LEASE_CALC: its charge-code lines by month and a lease total,
'                     then any leasing costs typed on its LEASE_DEALS row
'   LEASE DEALS       one block per used LEASE_DEALS line (DEAL_SCHEDULE): income lines and total, then leasing costs and commissions
'   totals            in-place, lease deals, total tenant revenue, total leasing costs
' Every amount is a formula pointing at LEASE_CALC / DEAL_SCHEDULE (no values). DETAIL does not read this sheet.
' B3 = REBUILD NEEDED when the totals no longer match LEASE_CALC (new deal line, new cost, changed layout), otherwise DETAIL!B3.
'==============================================================================
Private Const RV_SHEET As String = "REVENUE"
Private Const RV_LASTCOL As Long = 17                ' A suite, B lease/code, C tenant/charge, D GL, E:P months, Q total
Private Const DS_SHEET As String = "DEAL_SCHEDULE"
Private Const DS_FIRSTROW As Long = 7
Private Const DS_LASTROW As Long = 2000

Public Sub BuildRevenue()
    Dim ws As Worksheet, wsC As Worksheet, wsD As Worksheet
    Dim vC As Variant, vD As Variant
    Dim r As Long, i As Long, c As Long, k As Long, src As Long
    Dim dec As Boolean, wM As Double
    Dim leaseFirst As Long, inLease As Boolean, inDeals As Boolean
    Dim totIn As String, totDeal As String, totCost As String
    Dim rIn As Long, rDeal As Long, rRev As Long, rCost As Long
    Dim lastDS As Long, dealNo As Long, hdrRow As Long

    Set ws = ThisWorkbook.Worksheets(RV_SHEET)
    Set wsC = ThisWorkbook.Worksheets(REV_SHEET)
    Set wsD = ThisWorkbook.Worksheets(DS_SHEET)
    ResetSheet ws, FMT_FONT
    dec = ShowDecimals()
    If dec Then wM = 13 Else wM = 11
    ws.Columns("A").ColumnWidth = 8
    ws.Columns("B").ColumnWidth = 12
    ws.Columns("C").ColumnWidth = 40
    ws.Columns("D").ColumnWidth = 10
    For c = 5 To 16
        ws.Columns(c).ColumnWidth = wM
    Next c
    ws.Columns("Q").ColumnWidth = wM + 2

    ' ---- title block and column headings
    ws.Range("A1").Formula = "=BuildingName&"" - REVENUE BY LEASE"""
    ws.Range("A2").Formula = "=""BUDGET FOR THE YEAR ""&BudgetYear"
    ws.Range("A1:A2").Font.Bold = True
    ws.Range("A1:A2").Font.Color = PalColor("TITLE")
    ws.Range("A3").Value = "Status"
    ws.Range("A4").Value = "Last built"
    ws.Range("B4").NumberFormat = "@"
    ws.Range("B4").Value = Format(Now, "yyyy-mm-dd hh:mm")   ' text so it spills into column C (column B is narrow)
    ws.Range("B4").HorizontalAlignment = -4131
    ws.Range("A5").Value = "SUITE"
    ws.Range("B5").Value = "LEASE / CODE"
    ws.Range("C5").Value = "TENANT / CHARGE"
    ws.Range("D5").Value = "GL"
    For c = 5 To 16
        ws.Cells(5, c).NumberFormat = "mmm-yy"
        ws.Cells(5, c).Formula = "=DATE(BudgetYear," & (c - 4) & ",1)"
    Next c
    ws.Range("Q5").Value = "TOTAL"
    PaintRow ws, 5, 1, RV_LASTCOL, "COLHDR", True, True
    ws.Range("E5:Q5").HorizontalAlignment = -4108

    ' ---- in-place leases (LEASE_CALC): header row = the only rows with a number in column U (block number); MRI lines; TOTAL row; COST rows
    vC = wsC.Range("A" & REV_FIRST & ":W" & REV_LAST).Value
    r = 7
    ws.Cells(r, 1).Value = "IN-PLACE LEASES (MRI)"
    PaintRow ws, r, 1, RV_LASTCOL, "INC_HDR", True, True
    r = r + 1
    inLease = False
    For i = 1 To UBound(vC, 1)
        src = REV_FIRST + i - 1
        If Left(CStr(vC(i, 1)), 11) = "LEASE DEALS" Then Exit For
        If VarType(vC(i, 21)) = vbDouble And CStr(vC(i, 2)) = "" Then
            If inLease Then r = r + 1
            inLease = False                                       ' empty block (fewer MRI leases than blocks): skip its rows
        ElseIf VarType(vC(i, 21)) = vbDouble And CStr(vC(i, 2)) <> "" Then
            ' lease header
            If inLease Then r = r + 1
            ws.Cells(r, 1).Formula = "=" & REV_SHEET & "!A" & src & "&"""""
            ws.Cells(r, 2).Formula = "=" & REV_SHEET & "!B" & src
            ws.Cells(r, 3).Formula = "=" & REV_SHEET & "!C" & src & "&IF(ISNUMBER(" & REV_SHEET & "!F" & src & "),""  - expires ""&TEXT(" & REV_SHEET & "!F" & src & ",""mm/dd/yyyy"")&IF(" & REV_SHEET & "!I" & src & "<>""Terminate"","" (""&" & REV_SHEET & "!I" & src & "&"")"",""""),"""")"
            PaintRow ws, r, 1, RV_LASTCOL, "INC_B", False, True
            r = r + 1
            leaseFirst = r
            inLease = True
        ElseIf inLease And CStr(vC(i, 5)) = "MRI" And CStr(vC(i, 2)) <> "" Then
            RvLine ws, r, REV_SHEET, src, "B", "C", "D", "G", dec
            PaintRow ws, r, 1, RV_LASTCOL, "INC_A", False, False
            r = r + 1
        ElseIf inLease And CStr(vC(i, 4)) = "TOTAL" Then
            RvTotal ws, r, "Total lease", leaseFirst, r - 1, dec
            totIn = totIn & IIf(totIn = "", "", ",") & r
            r = r + 1
        ElseIf inLease And CStr(vC(i, 5)) = "COST" Then
            If IsNumeric(vC(i, 19)) And Not IsEmpty(vC(i, 19)) Then
                If Abs(CDbl(vC(i, 19))) > 0.004 Then
                    RvLine ws, r, REV_SHEET, src, "", "C", "D", "G", dec
                    ws.Range(ws.Cells(r, 3), ws.Cells(r, RV_LASTCOL)).Font.Italic = True
                    totCost = totCost & IIf(totCost = "", "", ",") & r
                    r = r + 1
                End If
            End If
        End If
    Next i
    If inLease Then r = r + 1
    rIn = r
    RvSumRow ws, r, "TOTAL IN-PLACE LEASES", totIn, "INC_HDR", True, dec
    r = r + 2

    ' ---- lease deals (DEAL_SCHEDULE): IDX row starts a deal (column C = tenant or "(deal line n - empty)"); INC and COST role rows
    ws.Cells(r, 1).Value = "NEW LEASES / RENEWALS / EXPANSIONS (LEASE_DEALS)"
    PaintRow ws, r, 1, RV_LASTCOL, "INC_HDR", True, True
    r = r + 1
    lastDS = DS_LASTROW                                   ' fixed read range (no End(xlUp): not reliable under LibreOffice)
    vD = wsD.Range("A" & DS_FIRSTROW & ":S" & lastDS).Value
    inDeals = False
    hdrRow = 0
    For i = 1 To UBound(vD, 1)
        src = DS_FIRSTROW + i - 1
        If CStr(vD(i, 2)) = "IDX" Then
            If hdrRow > 0 Then RvCloseDeal ws, r, hdrRow, leaseFirst, totDeal, dec
            hdrRow = 0
            If Left(CStr(vD(i, 3)), 10) <> "(deal line" Then
                dealNo = CLng(vD(i, 1))
                ws.Cells(r, 2).Value = "Deal " & dealNo
                ws.Cells(r, 3).Formula = "=" & DS_SHEET & "!C" & src
                PaintRow ws, r, 1, RV_LASTCOL, "INC_B", False, True
                hdrRow = r
                r = r + 1
                leaseFirst = r
                inDeals = True
            End If
        ElseIf hdrRow > 0 And CStr(vD(i, 5)) = "INC" Then
            If IsNumeric(vD(i, 19)) And Not IsEmpty(vD(i, 19)) Then
                If Abs(CDbl(vD(i, 19))) > 0.004 Then
                    RvLine ws, r, DS_SHEET, src, "", "C", "D", "G", dec
                    PaintRow ws, r, 1, RV_LASTCOL, "INC_A", False, False
                    r = r + 1
                End If
            End If
        ElseIf hdrRow > 0 And CStr(vD(i, 5)) = "COST" Then
            If IsNumeric(vD(i, 19)) And Not IsEmpty(vD(i, 19)) Then
                If Abs(CDbl(vD(i, 19))) > 0.004 Then
                    If ws.Cells(r - 1, 3).Value <> "Total deal" And Not ws.Cells(r - 1, 3).Font.Italic Then
                        RvTotal ws, r, "Total deal", leaseFirst, r - 1, dec
                        totDeal = totDeal & IIf(totDeal = "", "", ",") & r
                        r = r + 1
                    End If
                    RvLine ws, r, DS_SHEET, src, "", "C", "D", "G", dec
                    ws.Range(ws.Cells(r, 3), ws.Cells(r, RV_LASTCOL)).Font.Italic = True
                    totCost = totCost & IIf(totCost = "", "", ",") & r
                    r = r + 1
                End If
            End If
        End If
    Next i
    If hdrRow > 0 Then RvCloseDeal ws, r, hdrRow, leaseFirst, totDeal, dec
    If Not inDeals Then
        ws.Cells(r, 3).Value = "No deals entered on LEASE_DEALS."
        ws.Cells(r, 3).Font.Italic = True
        r = r + 1
    End If
    r = r + 1
    rDeal = r
    RvSumRow ws, r, "TOTAL NEW LEASES / RENEWALS / EXPANSIONS", totDeal, "INC_HDR", True, dec
    r = r + 2
    rRev = r
    RvSumRow ws, r, "TOTAL TENANT REVENUE", rIn & "," & rDeal, "GRAND_HDR", True, dec
    r = r + 2
    rCost = r
    RvSumRow ws, r, "TOTAL LEASING COSTS (TI, commissions, legal, architect, permits - shown in italics above)", totCost, "TAX_HDR", True, dec
    r = r + 2

    ' ---- status (fixed cell B3)
    ws.Range("B3").Formula = "=IF(OR(ABS($Q$" & rRev & "-SUMIFS(" & REV_SHEET & "!$S$" & REV_FIRST & ":$S$" & REV_LAST & "," & REV_SHEET & "!$D$" & REV_FIRST & ":$D$" & REV_LAST & ",""TOTAL""))>=0.5," & _
        "ABS($Q$" & rCost & "-SUMIFS(" & REV_SHEET & "!$S$" & REV_FIRST & ":$S$" & REV_LAST & "," & REV_SHEET & "!$E$" & REV_FIRST & ":$E$" & REV_LAST & ",""COST""))>=0.5)," & _
        """REBUILD NEEDED - click Build Schedules"",DETAIL!$B$3)"
    ws.Range("B3").Font.Bold = True
    ws.Range("B3").HorizontalAlignment = -4131
    ws.Cells(r, 1).Value = "This sheet is generated by the BuildRevenue macro - do not type on it. Enter new leases, renewals, expansions, expiry overrides and leasing costs on LEASE_DEALS, then click Build Schedules."
    ws.Cells(r, 1).Font.Italic = True
    LockSheet ws
End Sub

' One line: labels from the source row (code column optional), months = the source month cells, total = sum
Private Sub RvLine(ByVal ws As Worksheet, ByVal r As Long, ByVal sh As String, ByVal src As Long, ByVal codeCol As String, _
                   ByVal lblCol As String, ByVal glCol As String, ByVal firstMonthCol As String, ByVal dec As Boolean)
    Dim c As Long, fm As Long
    If codeCol <> "" Then ws.Cells(r, 2).Formula = "=" & sh & "!" & codeCol & src & "&"""""
    ws.Cells(r, 3).Formula = "=""  ""&TRIM(" & sh & "!" & lblCol & src & ")"
    ws.Cells(r, 4).Formula = "=" & sh & "!" & glCol & src
    ws.Cells(r, 4).NumberFormat = "0"
    fm = ws.Range(firstMonthCol & "1").Column
    For c = 5 To 16
        ws.Cells(r, c).Formula = "=" & sh & "!" & ColL(fm + c - 5) & src
    Next c
    RvTotCol ws, r, dec
End Sub

' Lease / deal total: months = SUM of the lines above
Private Sub RvTotal(ByVal ws As Worksheet, ByVal r As Long, ByVal lbl As String, ByVal r1 As Long, ByVal r2 As Long, ByVal dec As Boolean)
    Dim c As Long
    ws.Cells(r, 3).Value = lbl
    For c = 5 To 16
        If r2 >= r1 Then
            ws.Cells(r, c).Formula = "=SUM(" & ColL(c) & r1 & ":" & ColL(c) & r2 & ")"
        Else
            ws.Cells(r, c).Formula = "=0"
        End If
    Next c
    RvTotCol ws, r, dec
    ws.Range(ws.Cells(r, 3), ws.Cells(r, RV_LASTCOL)).Font.Bold = True
    BorderTop ws, r, 3, RV_LASTCOL
End Sub

' After the last line of a deal: its total row (unless cost lines already closed it with one), then a blank row
Private Sub RvCloseDeal(ByVal ws As Worksheet, ByRef r As Long, ByVal hdrRow As Long, ByVal first As Long, ByRef totDeal As String, ByVal dec As Boolean)
    Dim k As Long, hasTotal As Boolean
    hasTotal = False
    For k = hdrRow + 1 To r - 1
        If ws.Cells(k, 3).Value = "Total deal" Then hasTotal = True
    Next k
    If Not hasTotal Then
        RvTotal ws, r, "Total deal", first, r - 1, dec
        totDeal = totDeal & IIf(totDeal = "", "", ",") & r
        r = r + 1
    End If
    r = r + 1
End Sub

' Summary row: months = sum of the listed rows (comma-separated)
Private Sub RvSumRow(ByVal ws As Worksheet, ByVal r As Long, ByVal lbl As String, ByVal rowsList As String, ByVal palKey As String, _
                     ByVal whiteText As Boolean, ByVal dec As Boolean)
    Dim c As Long, j As Long, f As String, p() As String
    ws.Cells(r, 1).Value = lbl
    For c = 5 To 16
        f = ""
        If rowsList <> "" Then
            p = Split(rowsList, ",")
            For j = LBound(p) To UBound(p)
                f = f & "+" & ColL(c) & p(j)
            Next j
        End If
        If f = "" Then f = "0"
        ws.Cells(r, c).Formula = "=" & f
    Next c
    RvTotCol ws, r, dec
    PaintRow ws, r, 1, RV_LASTCOL, palKey, whiteText, True
End Sub

Private Sub RvTotCol(ByVal ws As Worksheet, ByVal r As Long, ByVal dec As Boolean)
    ws.Cells(r, 17).Formula = "=SUM(E" & r & ":P" & r & ")"
    ws.Range(ws.Cells(r, 5), ws.Cells(r, 17)).NumberFormat = MoneyFormat(dec)
End Sub
