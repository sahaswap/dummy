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

' Silent when the theme manager is driving, chatty when you run these
' macros by hand.
'
' Switching themes used to raise a dialog from every routine involved -
' the two stylers here, the remover, plus the theme modules - and each
' one blocked the next until it was dismissed. These macros are still
' useful standalone, so the announcements are not deleted, only gated.
'
' Application.Run rather than a qualified call: if modThemeManager is not
' installed, or was pasted in under a different module name, this simply
' fails and returns False, and the macros behave exactly as they always
' did.
Private Function Quiet() As Boolean
    On Error Resume Next
    Quiet = CBool(Application.Run("ThemeIsQuiet"))
    On Error GoTo 0
End Function

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
    ' Fonts are set as WHOLE-COLUMN ops (not per row). Row-by-row Font writes
    ' were not reliably overriding the naming-convention formula cells, which
    ' clung to their old Arial. Whole-column writes flatten every cell at once.
    If lastRow >= 2 Then
        ' base body font for the entire data block (guarantees D is Segoe UI 9)
        With ws.Range("A2:E" & lastRow)
            .Font.Name = "Segoe UI": .Font.Size = 9: .Font.Color = TEXTCLR
            .Font.Bold = False: .Font.Italic = False
            .Font.Underline = xlUnderlineStyleNone
            .VerticalAlignment = xlCenter
        End With

        ' A = entity name -> navy bold, the block title
        With ws.Range("A2:A" & lastRow)
            .Font.Bold = True: .Font.Color = NAVY_DEEP
            .HorizontalAlignment = xlLeft: .IndentLevel = 1
        End With
        ' B = "Type of Search" -> italic secondary grey
        With ws.Range("B2:B" & lastRow)
            .Font.Italic = True: .Font.Color = TEXT2: .HorizontalAlignment = xlLeft
        End With
        ' C = Raw URL -> small, muted (long strings, keep them quiet)
        With ws.Range("C2:C" & lastRow)
            .Font.Size = 8: .Font.Color = TEXT2
        End With
        ' D = Naming Convention -> plain body text, left (force Segoe UI 9)
        With ws.Range("D2:D" & lastRow)
            .Font.Name = "Segoe UI": .Font.Size = 9: .Font.Color = TEXTCLR
            .HorizontalAlignment = xlLeft
        End With
        ' E = Open URL ("Link") -> centered, navy underlined link
        With ws.Range("E2:E" & lastRow)
            .HorizontalAlignment = xlCenter
            .Font.Underline = xlUnderlineStyleSingle: .Font.Color = NAVY
        End With

        ' zebra band by ENTITY BLOCK + row separators (per row)
        For r = 2 To lastRow
            blk = (r - 2) \ BLOCK                    ' 0 = Customer, 1 = CP1, ...
            ws.Range("A" & r & ":E" & r).Interior.Color = _
                IIf(blk Mod 2 = 0, WARMWHITE, PEARL)

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
    If Not Quiet() Then MsgBox "Search Matrix restyled to match Sheet1 (Navy & Gold).", vbInformation, "Theme applied"
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
    If Not Quiet() Then MsgBox "Search Matrix theme removed (reset to plain).", vbInformation
    Exit Sub
Fail:
    Application.ScreenUpdating = True
    MsgBox "RemoveSearchMatrixTheme failed: " & Err.Description, vbCritical
End Sub


'=====================================================================
' StyleBackendSettings - same Navy & Gold thematics for the tiny
' Backend_Settings config sheet, and hide the unused Flow_1 row.
'
' Flow_1 (row 3) is HIDDEN, not deleted: Module10 reads the PAD merge
' URL from B4, so deleting row 3 would shift Flow_Merge up to B3 and
' break the merge button. Hiding leaves B4 exactly where it is.
'=====================================================================
Sub StyleBackendSettings()
    Dim ws As Worksheet
    On Error GoTo Fail
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets("Backend_Settings")
    On Error GoTo Fail
    If ws Is Nothing Then MsgBox "Backend_Settings sheet not found.", vbCritical: Exit Sub

    Dim wasProt As Boolean
    On Error Resume Next
    wasProt = ws.ProtectContents
    ws.Unprotect Password:="p7ss"
    On Error GoTo Fail

    Application.ScreenUpdating = False

    ' Hide the unused Flow_1 row (keeps Flow_Merge at B4 - do NOT delete).
    ws.Rows(3).Hidden = True

    ' base look for the whole visible block
    With ws.Range("A1:B4")
        .Interior.Color = WARMWHITE
        .Borders.LineStyle = xlNone
        .Font.Name = "Segoe UI"
        .Font.Color = TEXTCLR
        .Font.Size = 10
        .Font.Bold = False
        .Font.Italic = False
        .VerticalAlignment = xlCenter
        .WrapText = False
    End With

    ' Title bar (row 1) - navy fill, gold text, like the other sheets.
    ws.Range("A1").Value = "Backend Settings"
    With ws.Range("A1:B1")
        .Interior.Color = NAVY
        .Font.Color = GOLD
        .Font.Bold = True
        .Font.Size = 11
        .HorizontalAlignment = xlLeft
    End With
    ws.Rows(1).RowHeight = 22
    With ws.Range("A1:B1").Borders(xlEdgeBottom)
        .LineStyle = xlContinuous: .Weight = xlMedium: .Color = GOLD
    End With

    ' Setting row (row 4): key = navy bold on pearl, value = body on warm white.
    With ws.Range("A4")
        .Font.Bold = True: .Font.Color = NAVY_DEEP
        .Interior.Color = PEARL
        .HorizontalAlignment = xlLeft
    End With
    ws.Range("B4").HorizontalAlignment = xlLeft

    ' gold frame around the visible block (hidden row 3 just doesn't show)
    With ws.Range("A1:B4")
        .Borders(xlEdgeLeft).LineStyle = xlContinuous: .Borders(xlEdgeLeft).Weight = xlThin: .Borders(xlEdgeLeft).Color = GOLD
        .Borders(xlEdgeRight).LineStyle = xlContinuous: .Borders(xlEdgeRight).Weight = xlThin: .Borders(xlEdgeRight).Color = GOLD
        .Borders(xlEdgeTop).LineStyle = xlContinuous: .Borders(xlEdgeTop).Weight = xlMedium: .Borders(xlEdgeTop).Color = GOLD
        .Borders(xlEdgeBottom).LineStyle = xlContinuous: .Borders(xlEdgeBottom).Weight = xlMedium: .Borders(xlEdgeBottom).Color = GOLD
    End With

    On Error Resume Next
    ws.Activate
    ActiveWindow.DisplayGridlines = False
    On Error GoTo Fail

    If wasProt Then ws.Protect Password:="p7ss"
    Application.ScreenUpdating = True
    If Not Quiet() Then MsgBox "Backend_Settings restyled (Navy & Gold); Flow_1 row hidden.", vbInformation, "Theme applied"
    Exit Sub
Fail:
    Application.ScreenUpdating = True
    MsgBox "StyleBackendSettings failed: " & Err.Description, vbCritical
End Sub
