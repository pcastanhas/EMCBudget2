Attribute VB_Name = "modFormat"
Option Explicit
'==============================================================================
' Presentation formatting for the generated budget schedules.
' Budget schedules are the final presentation product, so they use the colour
' scheme of the original Allwood DETAIL sheet (not the white/blue input convention).
' Every section calls these routines, so new sections inherit the same look.
'==============================================================================
Public Const FMT_FONT As String = "Calibri"
Public Const FMT_SIZE As Long = 10

' Colour palette. Key = <PALETTE>_<ROLE>.  HDR = header/footer band (white text),
' A / B = alternating detail lines (black text).
Public Function PalColor(ByVal key As String) As Long
    Select Case key
        Case "INC_HDR": PalColor = RGB(79, 129, 189)       ' Income            #4F81BD
        Case "INC_A": PalColor = RGB(220, 230, 241)        '                   #DCE6F1
        Case "INC_B": PalColor = RGB(184, 204, 228)        '                   #B8CCE4
        Case "EXP_HDR": PalColor = RGB(192, 80, 77)        ' Expenses          #C0504D
        Case "EXP_A": PalColor = RGB(242, 220, 219)        '                   #F2DCDB
        Case "EXP_B": PalColor = RGB(230, 184, 183)        '                   #E6B8B7
        Case "EXPSUB_HDR": PalColor = RGB(218, 150, 148)   ' Expense subsection #DA9694
        Case "EXPSUB_A": PalColor = RGB(242, 220, 219)     '                   #F2DCDB
        Case "EXPSUB_B": PalColor = RGB(230, 184, 183)     '                   #E6B8B7
        Case "TAX_HDR": PalColor = RGB(128, 100, 162)      ' Taxes, Imprvmts & Commissions  #8064A2
        Case "TAX_A": PalColor = RGB(230, 224, 236)        '                   #E6E0EC
        Case "TAX_B": PalColor = RGB(204, 193, 218)        '                   #CCC1DA
        Case "INS_HDR": PalColor = RGB(75, 172, 198)       ' Insurance, Debt Service        #4BACC6
        Case "INS_A": PalColor = RGB(219, 238, 244)        '                   #DBEEF4
        Case "INS_B": PalColor = RGB(183, 222, 232)        '                   #B7DEE8
        Case "PRO_HDR": PalColor = RGB(31, 73, 125)        ' Professional fees, other fees  #1F497D
        Case "PRO_A": PalColor = RGB(198, 217, 241)        '                   #C6D9F1
        Case "PRO_B": PalColor = RGB(142, 180, 227)        '                   #8EB4E3
        Case "GRAND_HDR": PalColor = RGB(89, 89, 89)       ' Total operating expenses       #595959
        Case "NOI_HDR": PalColor = RGB(196, 189, 151)      ' NOI lines         #C4BD97
        Case "NET_HDR": PalColor = RGB(195, 214, 155)      ' Net cash flow     #C3D69B
        Case "CASH_HDR": PalColor = RGB(192, 80, 77)       ' Projected cash balance         #C0504D
        Case "COLHDR": PalColor = RGB(155, 187, 89)        ' Column headings   #9BBB59
        Case "TITLE": PalColor = RGB(31, 73, 125)          ' Sheet title text  #1F497D
        Case Else: PalColor = RGB(255, 255, 255)
    End Select
End Function

' Palette used by each DETAIL section (1 income, 2 tenant-charge offsets, 3 leasing costs, 4-9 expense sections, 10 taxes, 11 insurance, 12 professional fees, 13 debt service, 14 other professional fees)
Public Function SecPalette(ByVal sec As Long) As String
    Select Case sec
        Case 1: SecPalette = "INC"
        Case 2: SecPalette = "EXP"
        Case 3: SecPalette = "TAX"
        Case 4 To 9: SecPalette = "EXPSUB"
        Case 10: SecPalette = "TAX"
        Case 11: SecPalette = "INS"
        Case 12: SecPalette = "PRO"
        Case 13: SecPalette = "INS"
        Case 14: SecPalette = "PRO"
        Case Else: SecPalette = "INC"
    End Select
End Function

' "Show decimals" option on INPUT (cell C6)
Public Function ShowDecimals() As Boolean
    ShowDecimals = (UCase(Trim(CStr(ThisWorkbook.Worksheets("INPUT").Range("C6").Value))) = "YES")
End Function

' Accounting format, whole dollars (default) or with cents
Public Function MoneyFormat(ByVal withDecimals As Boolean) As String
    If withDecimals Then
        MoneyFormat = "_(""$""* #,##0.00_);_(""$""* \(#,##0.00\);_(""$""* ""-""??_);_(@_)"
    Else
        MoneyFormat = "_(""$""* #,##0_);_(""$""* \(#,##0\);_(""$""* ""-""??_);_(@_)"
    End If
End Function

' Paint a whole row band: fill colour from the palette, white or black text
Public Sub PaintRow(ByVal ws As Worksheet, ByVal r As Long, ByVal c1 As Long, ByVal c2 As Long, _
                    ByVal fillKey As String, ByVal whiteText As Boolean, ByVal isBold As Boolean)
    Dim rg As Range
    Set rg = ws.Range(ws.Cells(r, c1), ws.Cells(r, c2))
    rg.Interior.Color = PalColor(fillKey)
    If whiteText Then
        rg.Font.Color = RGB(255, 255, 255)
    Else
        rg.Font.Color = RGB(0, 0, 0)
    End If
    rg.Font.Bold = isBold
End Sub
