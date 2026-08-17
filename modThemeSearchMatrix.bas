Attribute VB_Name = "modThemeSearchMatrix"
Option Explicit
'=====================================================================
' modThemeSearchMatrix  -  restyle the "Search Matrix" sheet into the
' SAME Navy & Gold thematics used on Sheet1, so the two sheets match.
'
' Colours are pulled straight from Sheet1's own palette:
'   Navy   #223456   (navbar / panel fills)
'   Gold   #D9A441   (accent) + soft gold #E5C27A
'   Pearl  #F3F1EC   +  Warm White #FFFDF8   (zebra bands)
'   Text   #252525   +  secondary grey #6B6B6B   +  warm border #D9D4C8
'   Font   Segoe UI   (Sheet1's face; the sheet was Arial before)
'
'   StyleSearchMatrix    - apply the theme (idempotent, safe to re-run)
'   RemoveSearchMatrixTheme - strip back to a plain neutral look
'
' Styling ONLY: no buttons, no shapes, no macros added. Formulas,
' hyperlinks (col E) and layout are all left untouched. The data is
' laid out in 5-row blocks (one entity each: Customer, then CP1..CP6),
' so the zebra bands are drawn per-block, not per-row - each entity
' reads as one titled band.
'=====================================================================
Private Const SHEET_NAME As String = "Search Matrix"
Private Const BLOCK As Long = 5          ' rows per entity block

' --- palette (from Sheet1) ---
Private Function NAVY() As Long:       NAVY = RGB(34, 52, 86):        End Function   ' #223456
Private Function NAVY_DEEP() As Long:  NAVY_DEEP = RGB(28, 48, 86):   End Function   ' #1C3056
Private Function GOLD() As Long:       GOLD = RGB(217, 164, 65):      End Function   ' #D9A441
Private Function GOLD_SOFT() As Long:  GOLD_SOFT = RGB(229, 194, 122): End Function  ' #E5C27A
Private Function PEARL() As Long:      PEARL = RGB(243, 241, 236):    End Function   ' #F3F1EC
Private Function WARMWHITE() As Long:  WARMWHITE = RGB(255, 253, 248): End Function  ' #FFFDF8
Private Function TEXTCLR() As Long:    TEXTCLR = RGB(37, 37, 37):     End Function   ' #252525
Private Function TEXT2() As Long:      TEXT2 = RGB(107, 107, 107):    End Function   ' #6B6B6B
Private Function WARMBORDER() As Long: WARMBORDER = RGB(217, 212, 200): End Function ' #D9D4C8

Sub StyleSearchMatrix()
    Dim ws As Worksheet, r As Long, lastRow As Long, blk As Long
    On Error GoTo Fail

    On Error Resume Next
    Set ws = ThisWorkbook.Sheets(SHEET_NAME)
    On Error GoTo Fail
    If ws Is Nothing Then MsgBox SHEET_NAME & " sheet not found.", vbCritical: Exit Sub

    ' honour protection if it's ever turned on later
    Dim wasProt As Boolean
    On Error Resume Next
    wasProt = ws.ProtectContents
    ws.Unprotect Password:="p7ss"
    On Error GoTo Fail

    lastRow = ws.Cells(ws.Rows.count, "B").End(xlUp).row
    If lastRow < 2 Then lastRow = 1

    Application.ScreenUpdating = False

    ' clean the visible slate first (A:E only - helpers F.. stay hidden)
    With ws.Range("A1:E" & lastRow)
        .Interior.Pattern = xlSolid
        .Borders.LineStyle = xlNone
        .Font.Name = "Segoe UI"
        .Font.Color = TEXTCLR
        .Font.Size = 9
        .Font.Bold = False
        .Font.Italic = False
        .VerticalAlignment = xlCenter
        .WrapText = False
    End With

    ' ---- HEADER ROW (navy bar + gold text, like the navbar) ----
    With ws.Range("A1:E1")
        .Interior.Color = NAVY
        .Font.Name = "Segoe UI"
        .Font.Color = GOLD
        .Font.Bold = True
        .Font.Size = 10.5
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
    End With
    ws.Rows(1).RowHeight = 24
    ' thick gold underline beneath the header
    With ws.Range("A1:E1").Borders(xlEdgeBottom)
        .LineStyle = xlContinuous: .Weight = xlMedium: .Color = GOLD
    End With

    ' ---- DATA ROWS ----
    If lastRow >= 2 Then
        For r = 2 To lastRow
            blk = (r - 2) \ BLOCK                    ' 0 = Customer, 1 = CP1, ...
            ' zebra by ENTITY BLOCK (warm white / pearl)
            ws.Range("A" & r & ":E" & r).Interior.Color = _
                IIf(blk Mod 2 = 0, WARMWHITE, PEARL)

            ' A = entity name -> navy, bold (the block title)
            With ws.Range("A" & r).Font
                .Bold = True: .Color = NAVY_DEEP
            End With
            ws.Range("A" & r).HorizontalAlignment = xlLeft
            ws.Range("A" & r).IndentLevel = 1

            ' B = "Type of Search" -> italic secondary grey
            With ws.Range("B" & r).Font
                .Italic = True: .Color = TEXT2
            End With

            ' C = Raw URL -> small, muted (long strings, keep them quiet)
            With ws.Range("C" & r).Font
                .Name = "Segoe UI": .Size = 8: .Color = TEXT2: .Bold = False: .Italic = False
            End With

            ' D = Naming Convention -> plain body text; force Segoe UI explicitly
            ' (these formula cells kept an old typeface otherwise)
            With ws.Range("D" & r)
                .HorizontalAlignment = xlLeft
                With .Font
                    .Name = "Segoe UI": .Size = 9: .Color = TEXTCLR: .Bold = False: .Italic = False
                End With
            End With

            ' E = Open URL ("Link") -> centered, band kept, navy underlined link.
            ' Set the band LAST here so it wins over any leftover white "Hyperlink"
            ' cell style that Hyperlinks.Add stamped on before.
            With ws.Range("E" & r)
                .HorizontalAlignment = xlCenter
                .Interior.Color = IIf(blk Mod 2 = 0, WARMWHITE, PEARL)
                With .Font
                    .Name = "Segoe UI": .Size = 9: .Underline = xlUnderline: .Color = NAVY: .Bold = False
                End With
            End With

            ' thin warm separator under every row
            With ws.Range("A" & r & ":E" & r).Borders(xlEdgeBottom)
                .LineStyle = xlContinuous: .Weight = xlThin: .Color = WARMBORDER
            End With
            ' stronger gold-soft divider at the END of each entity block
            If (r - 1) Mod BLOCK = 0 Then
                With ws.Range("A" & r & ":E" & r).Borders(xlEdgeBottom)
                    .LineStyle = xlContinuous: .Weight = xlThin: .Color = GOLD_SOFT
                End With
            End If
        Next r
    End If

    ' ---- outer gold frame around the whole table ----
    With ws.Range("A1:E" & lastRow)
        .Borders(xlEdgeLeft).LineStyle = xlContinuous: .Borders(xlEdgeLeft).Weight = xlThin: .Borders(xlEdgeLeft).Color = GOLD
        .Borders(xlEdgeRight).LineStyle = xlContinuous: .Borders(xlEdgeRight).Weight = xlThin: .Borders(xlEdgeRight).Color = GOLD
        .Borders(xlEdgeTop).LineStyle = xlContinuous: .Borders(xlEdgeTop).Weight = xlMedium: .Borders(xlEdgeTop).Color = GOLD
        .Borders(xlEdgeBottom).LineStyle = xlContinuous: .Borders(xlEdgeBottom).Weight = xlMedium: .Borders(xlEdgeBottom).Color = GOLD
    End With

    ' gridlines off for the clean Sheet1 look
    On Error Resume Next
    ws.Activate
    ActiveWindow.DisplayGridlines = False
    On Error GoTo Fail

    If wasProt Then ws.Protect Password:="p7ss"
    Application.ScreenUpdating = True
    MsgBox "Search Matrix restyled to match Sheet1 (Navy & Gold).", vbInformation, "Theme applied"
    Exit Sub

Fail:
    Application.ScreenUpdating = True
    MsgBox "StyleSearchMatrix failed: " & Err.Description, vbCritical
End Sub

Sub RemoveSearchMatrixTheme()
    Dim ws As Worksheet, lastRow As Long
    On Error GoTo Fail
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets(SHEET_NAME)
    On Error GoTo Fail
    If ws Is Nothing Then Exit Sub

    Dim wasProt As Boolean
    On Error Resume Next
    wasProt = ws.ProtectContents
    ws.Unprotect Password:="p7ss"
    On Error GoTo Fail

    lastRow = ws.Cells(ws.Rows.count, "B").End(xlUp).row
    If lastRow < 1 Then lastRow = 1

    Application.ScreenUpdating = False
    With ws.Range("A1:E" & lastRow)
        .Interior.Pattern = xlNone
        .Borders.LineStyle = xlNone
        .Font.Name = "Calibri"
        .Font.Size = 11
        .Font.Bold = False
        .Font.Italic = False
        .Font.Color = RGB(0, 0, 0)
    End With
    On Error Resume Next
    ws.Activate
    ActiveWindow.DisplayGridlines = True
    On Error GoTo Fail

    If wasProt Then ws.Protect Password:="p7ss"
    Application.ScreenUpdating = True
    MsgBox "Search Matrix theme removed (reset to plain).", vbInformation
    Exit Sub
Fail:
    Application.ScreenUpdating = True
    MsgBox "RemoveSearchMatrixTheme failed: " & Err.Description, vbCritical
End Sub
