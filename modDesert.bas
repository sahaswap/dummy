Attribute VB_Name = "modDesert"
Option Explicit
'=====================================================================
' modDesert - the Desert theme for the Beta dashboard.
'
' Named Desert, drawn from the production design of Villeneuve's Dune.
' The film references throughout these comments are the SOURCE of the
' design decisions, not the name of the theme - they are left in place
' because they are the reasons, and a reason with its origin stripped out
' is just an assertion.
'
' DO NOT call ApplyDesert directly from a button. Route through
' modThemeManager.ApplyThemeByKey, which guarantees any other theme is
' fully removed first. Applying one theme over another makes the second
' stash the FIRST theme's colours as "the original", and neither can be
' cleanly removed afterwards.
'
' ---------------------------------------------------------------------
' THE IDEA: Tron emits light. Dune blocks it.
'
' Nothing on the Grid casts a shadow - every object is a light source
' with a glowing edge. Everything in Arrakeen does the opposite: vast
' blunt slabs standing in hard sun, lit across their faces and throwing
' shadow underneath. So this theme uses NO glow anywhere, and gets its
' depth from real drop shadows instead. That one inversion is what stops
' it reading as "Tron in brown".
'
' EVERYTHING IS IN THE BUTTONS. An earlier pass had a dune ridge, a
' light shaft and horizontal rock strata down the sidebar. All three were
' removed: the strata in particular ran the full width of the slab and
' surfaced either side of every button as stray rules cutting through the
' stack, which is scenery actively fighting the controls it sits behind.
' A sidebar of eight macro buttons does not need a landscape. See the
' block above StyleButton for where the effort went instead.
'
' DELIBERATELY MONOCHROME. The film's desert has no saturated colour in
' it at all, so the hierarchy here is built from VALUE, not hue: Start is
' the brightest bone, ordinary buttons are mid sandstone, and only Reset
' breaks the warm family, in oxide rust. Adding a spice-blue accent was
' tempting and is canon, but a cool accent on a warm sheet would have
' pulled this back towards the Dark Blue theme it exists to be different
' from. Rust is also kept muted so it cannot be confused with the bright
' reds the risk-class conditional formatting owns.
'=====================================================================
Private Const SHEET_NAME As String = "Sheet1"
Private Const TAG As String = "DESERTORIG|"

' The theme used to be called Dune. A rename that only changed the new
' names would silently strand anything stashed under the old ones: shapes
' tagged DUNEORIG| would never be restored, DUNEADD_ decorations would
' never be swept, and a _DuneBak sheet would sit there holding the only
' copy of the original cell colours while the sheet stayed dark forever.
' So the old names are still recognised on the way OUT.
Private Const LEGACY_TAG As String = "DUNEORIG|"
Private Const LEGACY_ADD_PFX As String = "DUNEADD_"
Private Const LEGACY_BAK As String = "_DuneBak"
Private Const ADD_PFX As String = "DESERTADD_"
Private Const FOREIGN_PFX As String = "NGADD_"
Private Const BAK As String = "_DesertBak"
Private Const CANVAS As String = "A1:AC36"

' Shape types, numeric rather than named. A missing NAME fails the whole
' module at compile time; a number an Office build does not know simply
' fails under On Error and the shape keeps the type it had.
'   156 = msoShapeSnip2SameRectangle   both TOP corners cut - the lintel
'   1   = msoShapeRectangle            blunt monolith
Private Const SHP_LINTEL As Long = 156
Private Const SHP_SLAB As Long = 1
Private Const CUT As Single = 0.14          ' depth of the lintel cut
Private Const GAP As Single = 16            ' air between the two button groups

' Tracking. Dune's titles are set very wide - wider than Tron's, which is
' why these are larger than that theme's equivalents.
Private Const TRACK_TITLE As Single = 8#
Private Const TRACK_BADGE As Single = 3.5
Private Const TRACK_BANNER As Single = 2.5

' THE SUN IS NOT OVERHEAD.
'
' Every shadow on this sheet is thrown by the same low, raking light. The
' first pass offset shadows straight down, which is noon at the equator -
' the one lighting condition the desert never has, and the reason those
' shapes read as floating UI cards rather than as objects standing in
' sun. Offsetting X as well as Y puts a single light source somewhere off
' to the upper left, and because EVERY shape uses the same pair, the
' whole sheet agrees about where the sun is.
Private Const SUN_DX As Single = 3
Private Const SUN_DY As Single = 3

' --- palette ---
Private cNight As Long         ' ground - the canvas
Private cRock As Long          ' data rows
Private cSlab As Long          ' every shape face
Private cSeam As Long          ' unlit rule / border
Private cStone As Long         ' ordinary button edge
Private cSand As Long          ' secondary text, strata highs
Private cBone As Long          ' primary text, Start
Private cOxide As Long         ' Reset only


Private Sub InitPalette()
    ' Value-stepped, all one warm family. Three surface steps exactly as
    ' the Dark Blue theme uses, for the same reason: flat fills of a
    ' shared colour are what make a dark UI look deliberate.
    cNight = RGB(20, 18, 15)       ' #14120F  ground
    cRock = RGB(27, 26, 23)        ' #1B1A17  data rows - charcoal
    cSlab = RGB(42, 36, 29)        ' #2A241D  every shape face
    cSeam = RGB(58, 47, 37)        ' #3A2F25  taupe - all hairline rules
    cStone = RGB(140, 111, 82)     ' #8C6F52  sandstone - button edges
    cSand = RGB(201, 178, 138)     ' #C9B28A  secondary text
    cBone = RGB(237, 225, 200)     ' #EDE1C8  primary text, Start
    cOxide = RGB(163, 74, 42)      ' #A34A2A  Reset - muted, not risk-red

End Sub

'--------------------------------------------------------------------
Sub ApplyDesert()
    Dim ws As Worksheet, shp As Shape, n As Long
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets(SHEET_NAME)
    On Error GoTo 0
    If ws Is Nothing Then MsgBox SHEET_NAME & " not found.", vbCritical: Exit Sub

    InitPalette

    Dim wasProt As Boolean, wbProt As Boolean
    On Error Resume Next
    wasProt = ws.ProtectContents
    ws.Unprotect Password:="p7ss"
    wbProt = ThisWorkbook.ProtectStructure
    ThisWorkbook.Unprotect Password:="p7ss"
    On Error GoTo 0

    Application.ScreenUpdating = False
    ws.Activate

    ' Sheet1 has no title shape of its own - the only "Beta 3.6" on it is
    ' NGADD_TITLE_TXT, which Navy & Gold generates. Capture the caption
    ' before purging, or the sheet ends up with no title at all.
    Dim titleText As String
    titleText = CaptureTitleText(ws)

    PurgeAllSheetBackgroundImages ThisWorkbook
    UnhideBaselineShapes ws
    HideForeignGraphics ws
    PurgeForeignDecor ws
    RemoveAdded ws
    ApplyCells ws

    For Each shp In ws.Shapes
        If IsStylable(shp) Then
            StashOriginal shp
            Select Case ShapeRole(shp)
                Case "TITLE":     StyleTitle shp
                Case "PRIMARY":   StyleButton shp, cBone, 2       ' most finished
                Case "DANGER":    StyleButton shp, cOxide, 0      ' left raw
                Case "BUTTON":    StyleButton shp, cStone, 1
                Case "BADGE":     StyleBadge shp
                Case Else:        StyleBanner shp
            End Select
            n = n + 1
        End If
    Next shp

    ' Scenery removed on purpose. The strata bands ran the full width of
    ' the slab and surfaced either side of every button as stray rules
    ' cutting through the stack - decoration that fought the controls it
    ' was behind. The light well and the ridge went with them: all three
    ' were set dressing competing with the only things on this sidebar
    ' anyone actually uses. What is left is the buttons.
    TightenUtilityGap ws             ' close Navy & Gold's inherited dead space
    AddMoon ws                       ' sizes itself to the space left
    AddBadgeRules ws                 ' after the badges are styled
    AddBannerMarks ws                ' chapter ticks, after the banners are styled
    AddDesertTitle ws, titleText

    ApplySearchMatrix
    ApplyBackendSettings

    On Error Resume Next
    ActiveWindow.DisplayGridlines = False
    On Error GoTo 0

    If wasProt Then ws.Protect Password:="p7ss"
    If wbProt Then ThisWorkbook.Protect Password:="p7ss", Structure:=True
    Application.ScreenUpdating = True
End Sub

Sub RemoveDesert()
    Dim ws As Worksheet, shp As Shape, n As Long
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets(SHEET_NAME)
    On Error GoTo 0
    If ws Is Nothing Then Exit Sub

    Dim wasProt As Boolean, wbProt As Boolean
    On Error Resume Next
    wasProt = ws.ProtectContents
    ws.Unprotect Password:="p7ss"
    wbProt = ThisWorkbook.ProtectStructure
    ThisWorkbook.Unprotect Password:="p7ss"
    On Error GoTo 0

    Application.ScreenUpdating = False
    RemoveAdded ws
    For Each shp In ws.Shapes
        If Left$(GetAlt(shp), Len(TAG)) = TAG _
           Or Left$(GetAlt(shp), Len(LEGACY_TAG)) = LEGACY_TAG Then
            RestoreOriginal shp
            n = n + 1
        End If
    Next shp
    UndoShift ws                     ' before the backup is cleared
    RestoreAllCells

    On Error Resume Next
    ActiveWindow.DisplayGridlines = True
    ShowGridlines "Search Matrix"
    ShowGridlines "Backend_Settings"
    On Error GoTo 0

    If wasProt Then ws.Protect Password:="p7ss"
    If wbProt Then ThisWorkbook.Protect Password:="p7ss", Structure:=True
    Application.ScreenUpdating = True
End Sub

'---- INHERITED-ARTEFACT CLEANUP -------------------------------------
Private Sub PurgeAllSheetBackgroundImages(ByVal wb As Workbook)
    On Error Resume Next
    Dim sh As Worksheet, prev As Object
    Set prev = ActiveSheet
    For Each sh In wb.Worksheets
        Dim wasP As Boolean
        wasP = sh.ProtectContents
        sh.Unprotect Password:="p7ss"
        sh.Activate
        Application.CommandBars.ExecuteMso "SheetBackgroundDelete"
        If wasP Then sh.Protect Password:="p7ss"
    Next sh
    If Not prev Is Nothing Then prev.Activate
    On Error GoTo 0
End Sub

Private Function CaptureTitleText(ByVal ws As Worksheet) As String
    On Error Resume Next
    Dim shp As Shape, t As String, dot As Long
    CaptureTitleText = ThisWorkbook.Name
    dot = InStrRev(CaptureTitleText, ".")
    If dot > 1 Then CaptureTitleText = Left$(CaptureTitleText, dot - 1)
    For Each shp In ws.Shapes
        t = ""
        If shp.TextFrame.HasText Then t = Trim$(shp.TextFrame.Characters.Text)
        If Len(t) > 0 And Len(shp.OnAction) = 0 Then
            If InStr(1, t, "Beta", vbTextCompare) > 0 Then
                CaptureTitleText = t
                Exit For
            End If
        End If
    Next shp
    On Error GoTo 0
End Function

' A theme may HIDE a real shape and draw its own replacement rather than
' restyling the original - Navy & Gold does that with the title and never
' un-hides it. Recover anything hidden before purging, or the purge takes
' out the only visible copy.
Private Sub UnhideBaselineShapes(ByVal ws As Worksheet)
    On Error Resume Next
    Dim shp As Shape
    For Each shp In ws.Shapes
        If Left$(shp.Name, Len(ADD_PFX)) <> ADD_PFX _
           And Left$(shp.Name, Len(FOREIGN_PFX)) <> FOREIGN_PFX Then
            If shp.Visible = msoFalse Then shp.Visible = msoTrue
        End If
    Next shp
    On Error GoTo 0
End Sub

' The gold shield beside the title is an msoGraphic - not an AutoShape,
' so no theme can recolour it, and not NGADD_-prefixed, so purging leaves
' it stranded. Hidden, never deleted; the manager un-hides it.
Private Sub HideForeignGraphics(ByVal ws As Worksheet)
    On Error Resume Next
    Dim shp As Shape, sbRight As Single
    sbRight = ws.Range("F1").Left
    For Each shp In ws.Shapes
        If Left$(shp.Name, Len(ADD_PFX)) <> ADD_PFX Then
            If shp.Type <> msoAutoShape And shp.Type <> msoFreeform Then
                If shp.Left < sbRight Then shp.Visible = msoFalse
            End If
        End If
    Next shp
    On Error GoTo 0
End Sub

Private Sub PurgeForeignDecor(ByVal ws As Worksheet)
    On Error Resume Next
    Dim i As Long
    For i = ws.Shapes.count To 1 Step -1
        If Left$(ws.Shapes(i).Name, Len(FOREIGN_PFX)) = FOREIGN_PFX Then ws.Shapes(i).Delete
    Next i
    On Error GoTo 0
End Sub

' Sweeps this theme's own decorations under BOTH the current prefix and
' the one it used when it was called Dune.
Private Sub RemoveAdded(ByVal ws As Worksheet)
    On Error Resume Next
    Dim i As Long, nm As String
    For i = ws.Shapes.count To 1 Step -1
        nm = ws.Shapes(i).Name
        If Left$(nm, Len(ADD_PFX)) = ADD_PFX _
           Or Left$(nm, Len(LEGACY_ADD_PFX)) = LEGACY_ADD_PFX Then
            ws.Shapes(i).Delete
        End If
    Next i
    On Error GoTo 0
End Sub

'---- CELLS -----------------------------------------------------------
' A COMPLETE Interior reset before the colour.
'
' Setting .Color alone is not enough on this sheet, and that is why a
' near-black #342D24 sidebar rendered as pale khaki. Sheet1's own
' Worksheet_Change applies xlPatternCrissCross with ThemeColor
' xlThemeColorDark1 and TintAndShade -0.15 to ranges here, and Navy &
' Gold uses patterns too. Two separate leftovers then survive a plain
' .Color assignment:
'
'   PATTERN blends the new colour with the pattern colour - a dark fill
'   under a CrissCross comes out light, because most of the cell is still
'   showing the pattern's background.
'
'   TINTANDSHADE is applied AFTER the colour, so a stale -0.15 silently
'   shifts whatever you just set. It is not reset by setting .Color.
'
' Every cell paint in this module goes through here so neither can come
' back.
Private Sub PaintCells(ByVal rng As Range, ByVal clr As Long)
    On Error Resume Next
    With rng.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .PatternTintAndShade = 0
        .TintAndShade = 0
        .Color = clr
    End With
    On Error GoTo 0
End Sub

Private Sub ApplyCells(ByVal ws As Worksheet)
    BackupSheetRange ws, CANVAS

    PaintCells ws.Range(CANVAS), cNight
    With ws.Range(CANVAS)
        ' Pattern FIRST, every time. Setting .Color while a pattern is
        ' still active blends the new colour with the pattern colour, and
        ' a near-black fill under a leftover CrissCross comes out as pale
        ' khaki - which is exactly how this sheet rendered. Navy & Gold
        ' and Sheet1's own Worksheet_Change both apply xlPatternCrissCross
        ' to ranges here, so there is always a pattern to inherit.
        .Font.Color = cSand
        .Borders(xlEdgeBottom).LineStyle = xlNone
        .Borders(xlInsideHorizontal).LineStyle = xlNone
    End With

    ' the dune face, crest to slipface
    ' THE RAIL RECEDES. One flat colour, and DARKER than the data rows.
    '
    ' Two faults were stacked here. Four stepped fills were meant to read
    ' as a dune face turning into the light, but their boundaries fell at
    ' rows 6/7, 14/15 and 22/23 - fixed rows that line up with nothing, so
    ' on screen they were three hard seams cutting across the button stack
    ' at arbitrary points. A gradient you cannot align to the content is
    ' just banding.
    '
    ' Worse, the lightest of those steps (#342D24) sat at the TOP, which
    ' is where the eye lands first - so the sidebar read as a pale slab
    ' dominating the sheet. A probe of the live workbook confirmed the
    ' paint was landing exactly as specified; the value was simply wrong.
    '
    ' The rail now takes the ground colour. A navigation rail should sit
    ' BEHIND the content it launches, not in front of it, and here that
    ' means the darkest tone on the sheet, not a raised panel. The region
    ' is defined by the buttons and labels standing on it - which is also
    ' the more monolithic reading, and the data bands are then the only
    ' thing that lifts off the ground.
    PaintCells ws.Range("A1:E29"), cNight

    ' Data bands follow Sheet1's own merge map: every row is G:I label +
    ' J:T value. Labels sit back in sand, values come forward in bone.
    ' Rows 17 and 26 are the in-table column-header rows.
    Dim bands As Variant, addr As Variant, sec As Range, r As Long, rr As Range
    bands = Array("G5:T10", "G13:T14", "G17:T23", "G26:T27")
    For Each addr In bands
        Set sec = ws.Range(CStr(addr))
        For r = sec.Row To sec.Row + sec.Rows.count - 1
            Set rr = ws.Range("G" & r & ":T" & r)
            rr.Font.Name = "Segoe UI"
            rr.Font.Italic = False

            If r = 17 Or r = 26 Then
                PaintCells rr, cSlab
                rr.Font.Color = cSand
                rr.Font.Bold = True
            Else
                PaintCells rr, cRock
                ws.Range("G" & r & ":I" & r).Font.Color = cStone     ' label
                ws.Range("G" & r & ":I" & r).Font.Bold = False
                ws.Range("J" & r & ":T" & r).Font.Color = cBone      ' value
            End If

            With rr.Borders(xlEdgeBottom)
                .LineStyle = xlContinuous: .Weight = xlHairline: .Color = cSeam
            End With
        Next r
    Next addr

    ' The theme picker. It sits inside CANVAS, so without this the one
    ' control that switches themes takes the plain ground fill and
    ' becomes the least visible thing on the sheet.
    PaintCells ws.Range("U28"), cSlab
    With ws.Range("U28")
        .Font.Name = "Segoe UI": .Font.Size = 9
        .Font.Color = cSand: .Font.Bold = True: .Font.Italic = False
        .Borders.LineStyle = xlContinuous
        .Borders.Weight = xlThin
        .Borders.Color = cStone
        .Locked = False
    End With
End Sub

'---- COMPANION SHEETS ------------------------------------------------
' Alignment is deliberately NOT set in either of these: modThemeManager
' re-asserts centre/middle after every theme, so setting it here would
' only be overwritten, and would imply alignment is a theme decision.
Private Sub ApplySearchMatrix()
    On Error Resume Next
    Dim ws As Worksheet, lastRow As Long, r As Long, blk As Long
    Set ws = ThisWorkbook.Sheets("Search Matrix")
    If ws Is Nothing Then Exit Sub

    Dim wasProt As Boolean
    wasProt = ws.ProtectContents
    ws.Unprotect Password:="p7ss"

    lastRow = ws.Cells(ws.Rows.count, "B").End(xlUp).Row
    If lastRow < 1 Then lastRow = 1

    BackupSheetRange ws, "A1:E" & lastRow

    PaintCells ws.Range("A1:E" & lastRow), cRock
    With ws.Range("A1:E" & lastRow)
        .Borders.LineStyle = xlNone
        .Font.Name = "Segoe UI"
        .Font.Size = 9
        .Font.Color = cBone
        .Font.Bold = False
        .Font.Italic = False
        .WrapText = False
    End With

    PaintCells ws.Range("A1:E1"), cSlab
    With ws.Range("A1:E1")
        .Font.Color = cSand
        .Font.Bold = True
        .Font.Size = 10
    End With
    ws.Rows(1).RowHeight = 24
    With ws.Range("A1:E1").Borders(xlEdgeBottom)
        .LineStyle = xlContinuous: .Weight = xlThin: .Color = cStone
    End With

    If lastRow >= 2 Then
        With ws.Range("A2:A" & lastRow)
            .Font.Bold = True: .Font.Color = cBone
        End With
        ws.Range("B2:B" & lastRow).Font.Color = cStone
        With ws.Range("C2:C" & lastRow)
            .Font.Size = 8: .Font.Color = cStone
        End With
        ws.Range("D2:D" & lastRow).Font.Color = cBone
        With ws.Range("E2:E" & lastRow)
            .Font.Color = cSand
            .Font.Underline = xlUnderlineStyleSingle
        End With

        ' Banded per 5-row entity block, not per row - the data is one
        ' entity per block, so row striping would fight the grouping.
        For r = 2 To lastRow
            blk = (r - 2) \ 5
            PaintCells ws.Range("A" & r & ":E" & r), _
                IIf(blk Mod 2 = 0, cRock, cSlab)
            With ws.Range("A" & r & ":E" & r).Borders(xlEdgeBottom)
                .LineStyle = xlContinuous: .Weight = xlHairline: .Color = cSeam
            End With
            If (r - 1) Mod 5 = 0 Then
                With ws.Range("A" & r & ":E" & r).Borders(xlEdgeBottom)
                    .LineStyle = xlContinuous: .Weight = xlThin: .Color = cSeam
                End With
            End If
        Next r
    End If

    FrameRange ws.Range("A1:E" & lastRow)
    KillGridlines ws
    If wasProt Then ws.Protect Password:="p7ss"
    On Error GoTo 0
End Sub

Private Sub ApplyBackendSettings()
    On Error Resume Next
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("Backend_Settings")
    If ws Is Nothing Then Exit Sub

    Dim wasProt As Boolean
    wasProt = ws.ProtectContents
    ws.Unprotect Password:="p7ss"

    BackupSheetRange ws, "A1:B4"

    PaintCells ws.Range("A1:B4"), cRock
    With ws.Range("A1:B4")
        .Borders.LineStyle = xlNone
        .Font.Name = "Segoe UI"
        .Font.Color = cBone
        .Font.Size = 10
        .Font.Bold = False
        .Font.Italic = False
        .WrapText = False
    End With

    PaintCells ws.Range("A1:B1"), cSlab
    With ws.Range("A1:B1")
        .Font.Color = cSand
        .Font.Bold = True
        .Font.Size = 11
    End With
    ws.Rows(1).RowHeight = 22
    With ws.Range("A1:B1").Borders(xlEdgeBottom)
        .LineStyle = xlContinuous: .Weight = xlThin: .Color = cStone
    End With

    ws.Range("A4").Font.Bold = True
    ws.Range("A4").Font.Color = cStone
    ws.Range("B4").Font.Color = cBone

    FrameRange ws.Range("A1:B4")
    KillGridlines ws
    If wasProt Then ws.Protect Password:="p7ss"
    On Error GoTo 0
End Sub

Private Sub FrameRange(ByVal rng As Range)
    On Error Resume Next
    Dim e As Variant
    For Each e In Array(xlEdgeLeft, xlEdgeRight, xlEdgeTop, xlEdgeBottom)
        With rng.Borders(CLng(e))
            .LineStyle = xlContinuous: .Weight = xlThin: .Color = cStone
        End With
    Next e
    On Error GoTo 0
End Sub

Private Sub KillGridlines(ByVal ws As Worksheet)
    On Error Resume Next
    Dim prev As Object
    Set prev = ActiveSheet
    ws.Activate
    ActiveWindow.DisplayGridlines = False
    If Not prev Is Nothing Then prev.Activate
    On Error GoTo 0
End Sub

Private Sub ShowGridlines(ByVal nm As String)
    On Error Resume Next
    Dim prev As Object, ws As Worksheet
    Set ws = ThisWorkbook.Sheets(nm)
    If ws Is Nothing Then Exit Sub
    Set prev = ActiveSheet
    ws.Activate
    ActiveWindow.DisplayGridlines = True
    If Not prev Is Nothing Then prev.Activate
    On Error GoTo 0
End Sub

'=====================================================================
' STRATA - the sidebar's identity, and Tron's circuit spine inverted.
'
' Rock laid down in layers: horizontal bands at IRREGULAR intervals, in
' graduated sand tones, some running the full width of the slab and some
' stopping short. The irregularity is the whole point - evenly spaced
' lines read as a grid, and a grid is the one thing a desert is not.
'
' Sent to the back, so buttons and cards sit on top of the rock face
' rather than being interrupted by it.
'=====================================================================



'=====================================================================
' BANNER MARKS - a carved tick at the left edge of each section header.
'
' The film's title cards mark a chapter with a single small rule rather
' than a heading style. Four ticks, one per banner, positioned off each
' banner's own geometry so they follow wherever the banners sit.
'=====================================================================
Private Sub AddBannerMarks(ByVal ws As Worksheet)
    On Error Resume Next
    Dim shp As Shape, mk As Shape, seq As Long
    For Each shp In ws.Shapes
        If Left$(shp.Name, Len(ADD_PFX)) <> ADD_PFX Then
            If ShapeRole(shp) = "PANEL" And shp.Width > 80 Then
                seq = seq + 1
                Set mk = ws.Shapes.AddShape(SHP_SLAB, _
                            shp.Left + 7, shp.Top + shp.Height * 0.28, _
                            2, shp.Height * 0.44)
                If Not mk Is Nothing Then
                    mk.Name = ADD_PFX & "MARK" & Format$(seq, "00")
                    With mk.Fill
                        .Visible = msoTrue: .Solid
                        .ForeColor.RGB = cStone
                        .Transparency = 0.15
                    End With
                    mk.Line.Visible = msoFalse
                    mk.Shadow.Visible = msoFalse
                    mk.Placement = xlFreeFloating
                    mk.ZOrder msoBringToFront
                End If
            End If
        End If
    Next shp
    On Error GoTo 0
End Sub



' Finds a panel badge by its caption. Used by the title, which sizes
' itself to the clear space above the ACTION badge rather than guessing.
'
' Left as a Function on purpose: it returns a Shape, and the title needs
' the object, not just whether one exists.


'=====================================================================
' CLOSING THE UTILITY GAP.
'
' Navy & Gold anchors the UTILITY badge to the Counterparty Information
' banner - utlLbl.Top = cpBanner.Top - so the two panels line up across
' the sheet. That symmetry costs a large dead space between "Generate
' Narrative" and "UTILITY PANEL", because the action group is shorter
' than the alert section it sits beside. Desert does not inherit that
' trade: nothing in this theme lines the sidebar up with the main area,
' so the gap buys nothing and just reads as a hole.
'
' The shapes are moved by a single DELTA, and that delta is written to
' the backup sheet. Removal shifts the same shapes back by the same
' amount, so this is exactly reversible without storing a position per
' shape. One number, not twenty.
'
' Guarded three ways: only ever moves UP, never runs if the gap is
' already tight, and refuses a shift larger than 200pt in case the badge
' was not found where expected and the arithmetic went wrong.
'=====================================================================
Private Sub TightenUtilityGap(ByVal ws As Worksheet)
    On Error Resume Next
    Dim badge As Shape, shp As Shape
    Dim sbRight As Single, lastActionBottom As Single, delta As Single

    Set badge = FindBadge(ws, "UTILITY")
    If badge Is Nothing Then Exit Sub
    sbRight = ws.Range("F1").Left

    ' the lowest thing in the sidebar that sits ABOVE the utility badge
    For Each shp In ws.Shapes
        If Left$(shp.Name, Len(ADD_PFX)) <> ADD_PFX Then
            If shp.Left < sbRight And shp.Top < badge.Top Then
                If shp.Top + shp.Height > lastActionBottom Then
                    lastActionBottom = shp.Top + shp.Height
                End If
            End If
        End If
    Next shp
    If lastActionBottom = 0 Then Exit Sub

    delta = (lastActionBottom + GAP) - badge.Top
    If delta >= -2 Then Exit Sub             ' already tight, or would move down
    If delta < -200 Then Exit Sub            ' implausible - do not touch anything

    ShiftSidebarBelow ws, badge.Top - 1, delta
    StoreShift delta
    On Error GoTo 0
End Sub

' Moves every non-generated sidebar shape at or below fromTop by delta.
Private Sub ShiftSidebarBelow(ByVal ws As Worksheet, ByVal fromTop As Single, _
                              ByVal delta As Single)
    On Error Resume Next
    Dim shp As Shape, sbRight As Single
    sbRight = ws.Range("F1").Left
    For Each shp In ws.Shapes
        If Left$(shp.Name, Len(ADD_PFX)) <> ADD_PFX Then
            If shp.Left < sbRight And shp.Top >= fromTop Then
                shp.Top = shp.Top + delta
            End If
        End If
    Next shp
    On Error GoTo 0
End Sub

Private Sub StoreShift(ByVal delta As Single)
    On Error Resume Next
    Dim bak As Worksheet: Set bak = GetBak(True)
    If bak Is Nothing Then Exit Sub
    bak.Range("J1").Value = delta
    On Error GoTo 0
End Sub

' Read and undo the shift. Must run BEFORE RestoreAllCells, which clears
' the backup sheet.
Private Sub UndoShift(ByVal ws As Worksheet)
    On Error Resume Next
    Dim bak As Worksheet, delta As Single, badge As Shape
    Set bak = GetBak(False)
    If bak Is Nothing Then Exit Sub
    If Len(CStr(bak.Range("J1").Value)) = 0 Then Exit Sub
    delta = CSng(bak.Range("J1").Value)
    If delta = 0 Then Exit Sub

    Set badge = FindBadge(ws, "UTILITY")
    If Not badge Is Nothing Then ShiftSidebarBelow ws, badge.Top - 1, -delta
    bak.Range("J1").ClearContents
    On Error GoTo 0
End Sub

'=====================================================================
' THE MOON - two concentric rings, in whatever space is actually free.
'
' Twice placed badly. Behind the title it crossed "Beta 3.6.2", because a
' ring and a line of type in the same space is interference, not
' layering. Moved to the foot of the sidebar it was given a fixed 54pt
' diameter centred on row 29 - so half of it spilled past the rail onto
' the canvas and read as a clipped oval.
'
' Both failures were the same mistake: a fixed size dropped at a fixed
' point, with no regard for how much room was there. This measures the
' gap between the lowest button and the bottom of the rail and fits
' itself inside it. If the button stack grows and the gap closes, the
' moon simply is not drawn - which is the correct behaviour, and far
' better than drawing it half off the edge.
'
' Two rings at a step rather than one circle: the Carlo Scarpa detail
' from Vermette's reference list, the same profile repeated at an offset.
'=====================================================================
Private Sub AddMoon(ByVal ws As Worksheet)
    On Error Resume Next
    Dim shp As Shape, sbRight As Single
    Dim lowest As Single, railBottom As Single, band As Single
    Dim d As Single, cx As Single, cy As Single

    sbRight = ws.Range("F1").Left
    railBottom = ws.Range("A29").Top + ws.Range("A29").Height

    ' the bottom of the lowest real shape in the rail
    For Each shp In ws.Shapes
        If Left$(shp.Name, Len(ADD_PFX)) <> ADD_PFX Then
            If shp.Left < sbRight Then
                If shp.Top + shp.Height > lowest Then lowest = shp.Top + shp.Height
            End If
        End If
    Next shp
    If lowest = 0 Then Exit Sub

    band = railBottom - lowest
    If band < 26 Then Exit Sub            ' no room - draw nothing at all

    d = band - 10
    If d > 46 Then d = 46
    cx = ws.Range("A1").Left + ws.Range("A1:E1").Width * 0.62   ' off-axis
    cy = lowest + band / 2

    MoonRing ws, cx - d / 2, cy - d / 2, d, 1#, 0.55, "MOON1"
    MoonRing ws, cx - (d - 11) / 2, cy - (d - 11) / 2, d - 11, 0.75, 0.72, "MOON2"
    On Error GoTo 0
End Sub

Private Sub MoonRing(ByVal ws As Worksheet, ByVal x As Single, ByVal y As Single, _
                     ByVal d As Single, ByVal wt As Single, ByVal trans As Single, _
                     ByVal nm As String)
    On Error Resume Next
    Dim shp As Shape
    Set shp = ws.Shapes.AddShape(msoShapeOval, x, y, d, d)
    If shp Is Nothing Then Exit Sub
    shp.Name = ADD_PFX & nm
    shp.Fill.Visible = msoFalse
    With shp.Line
        .Visible = msoTrue
        .ForeColor.RGB = cStone
        .Weight = wt
        .Transparency = trans
    End With
    shp.Shadow.Visible = msoFalse
    shp.Glow.Radius = 0
    shp.Placement = xlFreeFloating
    shp.ZOrder msoSendToBack
    On Error GoTo 0
End Sub

Private Function FindBadge(ByVal ws As Worksheet, ByVal key As String) As Shape
    On Error Resume Next
    Dim shp As Shape, t As String
    For Each shp In ws.Shapes
        If Len(shp.OnAction) = 0 And Left$(shp.Name, Len(ADD_PFX)) <> ADD_PFX Then
            t = ""
            If shp.TextFrame.HasText Then t = shp.TextFrame.Characters.Text
            If InStr(1, t, key, vbTextCompare) > 0 Then Set FindBadge = shp: Exit For
        End If
    Next shp
    On Error GoTo 0
End Function

Private Sub AddDesertTitle(ByVal ws As Worksheet, ByVal caption As String)
    On Error Resume Next
    Dim shp As Shape, badge As Shape
    Dim x As Single, w As Single, y As Single, h As Single

    x = ws.Range("A1").Left + 14
    w = ws.Range("A1:E1").Width - 28
    y = ws.Range("A1").Top + 4

    Set badge = FindBadge(ws, "ACTION")
    If Not badge Is Nothing Then
        h = (badge.Top - y) - 10
    Else
        h = 22
    End If
    If h > 24 Then h = 24
    If h < 16 Then h = 16

    Set shp = ws.Shapes.AddTextbox(msoTextOrientationHorizontal, x, y, w, h)
    If shp Is Nothing Then Exit Sub
    shp.Name = ADD_PFX & "TITLE"
    shp.TextFrame.Characters.Text = caption
    With shp.TextFrame.Characters.Font
        .Name = "Segoe UI"
        .Size = 12
        .Bold = False                  ' Dune's titles are thin, not heavy
        .Color = cBone
    End With
    With shp.TextFrame2
        .TextRange.ParagraphFormat.Alignment = msoAlignCenter
        .VerticalAnchor = msoAnchorMiddle
        .MarginTop = 0: .MarginBottom = 0
        .WordWrap = msoFalse
        .TextRange.Font.Spacing = TRACK_TITLE
    End With
    shp.Fill.Visible = msoFalse
    shp.Line.Visible = msoFalse
    shp.Shadow.Visible = msoFalse
    shp.ZOrder msoBringToFront

    ' Rules ABOVE and below. The film's title cards bracket their type
    ' between two rules with a lot of air; one rule underneath reads as a
    ' heading with an underline, which is a different and much more
    ' ordinary thing. The upper rule is shorter than the lower one so the
    ' pair is not a symmetrical box.
    TitleRule ws, x + w * 0.32, y - 4, w * 0.36, "TITLERULETOP"
    TitleRule ws, x + w * 0.2, y + h + 3, w * 0.6, "TITLERULEBOT"
    On Error GoTo 0
End Sub

Private Sub TitleRule(ByVal ws As Worksheet, ByVal x As Single, ByVal y As Single, _
                      ByVal w As Single, ByVal nm As String)
    On Error Resume Next
    Dim ln As Shape
    Set ln = ws.Shapes.AddShape(SHP_SLAB, x, y, w, 1)
    If ln Is Nothing Then Exit Sub
    ln.Name = ADD_PFX & nm
    With ln.Fill
        .Visible = msoTrue: .Solid
        .ForeColor.RGB = cStone
        .Transparency = 0.3
    End With
    ln.Line.Visible = msoFalse
    ln.Shadow.Visible = msoFalse
    ln.Glow.Radius = 0
    ln.Placement = xlFreeFloating
    On Error GoTo 0
End Sub

'---- classification --------------------------------------------------
Private Function IsStylable(ByVal shp As Shape) As Boolean
    On Error Resume Next
    IsStylable = (shp.Type = msoAutoShape Or shp.Type = msoFreeform)
    If Left$(shp.Name, Len(ADD_PFX)) = ADD_PFX Then IsStylable = False
    On Error GoTo 0
End Function

' TITLE / PRIMARY / DANGER / BUTTON / BADGE / PANEL.
' The title test matches "Beta" ANYWHERE, not just at position 1 - Navy &
' Gold prepends a shield glyph to captions, so an anchored test silently
' fails on a themed workbook and the title falls through to the button
' branch.
Private Function ShapeRole(ByVal shp As Shape) As String
    On Error Resume Next
    Dim t As String, hasMac As Boolean
    ShapeRole = "PANEL"
    hasMac = (Len(shp.OnAction) > 0)
    If shp.TextFrame.HasText Then t = Trim$(shp.TextFrame.Characters.Text)

    If Len(t) > 0 Then
        If InStr(1, t, "Beta", vbTextCompare) > 0 And Not hasMac Then
            ShapeRole = "TITLE": Exit Function
        End If
    End If

    If hasMac Then
        If InStr(1, t, "Reset", vbTextCompare) > 0 Then
            ShapeRole = "DANGER"
        ElseIf InStr(1, t, "Start", vbTextCompare) > 0 Then
            ShapeRole = "PRIMARY"
        Else
            ShapeRole = "BUTTON"
        End If
    ElseIf InStr(1, t, "UTILITY", vbTextCompare) > 0 _
        Or InStr(1, t, "ACTION", vbTextCompare) > 0 Then
        ShapeRole = "BADGE"
    End If
    On Error GoTo 0
End Function

'---- styling ---------------------------------------------------------
' Buttons are battered lintels: both TOP corners cut back, square at the
' base. Flat fill - the same face colour as every other shape - and a
' real drop shadow rather than a glow. Depth here comes from mass
' blocking light, which is the whole conceit of the theme.
'=====================================================================
' THE BUTTONS - the only things on this sidebar anyone actually uses,
' so they get the whole design budget.
'
' Three ideas, each doing work beyond decoration:
'
' 1. THE CUT IS IN THE OUTLINE, NOT IN A BEVEL.
'    An earlier pass used Office's ThreeD bevel to give the block a
'    carved lip. Excel renders a bevel through its own 3-D lighting,
'    which softens and lightens the whole face - it came out looking like
'    moulded plastic. The snipped top corners carry the "cut stone" idea
'    on their own; a crisp flat fill inside a thin line is sharper than
'    anything the bevel adds.
'
' 2. THE CAPTION IS AN INSCRIPTION, NOT A LABEL.
'    Left-aligned on a deep margin with slight tracking, the way text is
'    cut into a lintel. It is also the more useful arrangement: eight
'    centred captions of different widths give eight different starting
'    points, and the eye re-finds the line on every row, where aligning
'    them left forms one vertical edge and the stack scans in a single
'    pass.
'
'    This is DESERT ONLY. Dark Blue keeps its captions centred - that
'    theme's buttons are chamfered on opposite corners, which is a
'    symmetrical figure, and a left-aligned caption inside a symmetrical
'    block reads as a mistake rather than a decision. The alignment
'    follows the shape, so the two themes differ on purpose.
'
' 3. HIERARCHY BY DEGREE OF FINISH.
'    Not just colour. Start is the most finished block - deepest cut,
'    fullest bevel, bone edge. The ordinary buttons are cut and lightly
'    bevelled. Reset is left RAW: same stone, no bevel at all, the one
'    unworked block in the wall. So the destructive control is set apart
'    by form as well as by its oxide edge, and it stays distinguishable
'    even to someone who cannot separate the two warm colours.
'=====================================================================
Private Sub StyleButton(ByVal shp As Shape, ByVal edge As Long, ByVal finish As Long)
    On Error Resume Next
    shp.AutoShapeType = SHP_LINTEL
    ' the most finished block carries the deepest cut
    shp.Adjustments(1) = IIf(finish >= 2, CUT * 1.5, CUT)

    With shp.Fill
        .Visible = msoTrue: .Solid
        .ForeColor.RGB = cSlab
        .Transparency = 0
    End With
    With shp.Line
        .Visible = msoTrue
        .ForeColor.RGB = edge
        .Weight = IIf(finish >= 2, 2#, 1#)
        .Transparency = 0
    End With

    ' NO BEVEL. It was here to give the block a carved lip, and I flagged
    ' the risk when adding it: Excel renders a ThreeD bevel through its own
    ' 3-D lighting, which softens and lightens the whole face. On screen it
    ' read as moulded plastic - the Office 2007 look - which is the exact
    ' opposite of cut stone, and it was what changed the buttons' character
    ' for the worse. A crisp 1pt line on a flat fill is sharper and more
    ' like masonry than any bevel Excel can draw.
    shp.ThreeD.BevelTopType = msoBevelNone

    shp.Glow.Radius = 0                 ' explicitly none
    With shp.Shadow
        .Type = msoShadow25
        .Visible = msoTrue
        .ForeColor.RGB = vbBlack
        .Transparency = IIf(finish >= 2, 0.45, 0.6)
        .Blur = 5
        .OffsetX = SUN_DX
        .OffsetY = SUN_DY
    End With
    shp.SoftEdge.Type = 0

    If shp.TextFrame.HasText Then
        With shp.TextFrame.Characters.Font
            .Name = "Segoe UI"
            .Bold = (finish >= 2)
            .Size = 9
            .Color = IIf(finish >= 2, edge, cSand)
        End With
        With shp.TextFrame2
            .TextRange.ParagraphFormat.Alignment = msoAlignLeft
            .VerticalAnchor = msoAnchorMiddle
            .MarginLeft = 11
            .MarginRight = 4
            .WordWrap = msoFalse
            .TextRange.Font.Spacing = 1.2
        End With
        ClearTextGlow shp
    End If
    On Error GoTo 0
End Sub

' Badges taper - wider at the base than the top, the way Arrakeen's walls
' are battered. The taper is kept shallow so the caption still clears the
' cut corners.
' The badges are LABELS, not plates.
'
' They were trapezoids - battered walls, wider at the base. That was
' wrong three ways. A battered wall tapers because it is load-bearing
' masonry resisting lateral force; a label plate carries nothing, so the
' form was borrowed without its reason. It was also the only tapered
' shape on a sheet of otherwise orthogonal ones, which reads as a mistake
' rather than a motif - a signature needs repetition or an obvious cause,
' and this had neither. And its sides sloped inward directly above a card
' of full-width vertical buttons, so it disagreed with the group it was
' labelling.
'
' So the plate is gone entirely: no fill, no border, no shadow. Just
' tracked caps over a hairline rule, which is exactly how the title is
' set. That gives the sheet ONE typographic idea used twice at two
' scales, instead of a second competing object - and a label can no
' longer be mistaken for a control you could click.
'
' AutoShapeType is reset to a plain rectangle even though nothing is
' drawn: a shape's text frame is fitted to its geometry, so a leftover
' trapezoid would keep pinching the caption at the corners even with the
' fill switched off.
Private Sub StyleBadge(ByVal shp As Shape)
    On Error Resume Next
    shp.AutoShapeType = SHP_SLAB
    shp.Fill.Visible = msoFalse
    shp.Line.Visible = msoFalse
    shp.Glow.Radius = 0
    shp.Shadow.Visible = msoFalse
    shp.SoftEdge.Type = 0
    If shp.TextFrame.HasText Then
        With shp.TextFrame.Characters.Font
            .Name = "Segoe UI"
            .Color = cSand
            .Bold = False              ' the tracking carries it, not weight
            .Italic = False
            .Size = 9
        End With
        ClearTextGlow shp
        SetTracking shp, TRACK_BADGE
    End If
    On Error GoTo 0
End Sub

' The hairline each badge sits on. Drawn off the badge's own geometry, so
' it follows wherever Navy & Gold last repositioned them.
Private Sub AddBadgeRules(ByVal ws As Worksheet)
    On Error Resume Next
    Dim shp As Shape, ln As Shape, seq As Long
    For Each shp In ws.Shapes
        If Left$(shp.Name, Len(ADD_PFX)) <> ADD_PFX Then
            If ShapeRole(shp) = "BADGE" Then
                seq = seq + 1
                Set ln = ws.Shapes.AddShape(SHP_SLAB, _
                            shp.Left + shp.Width * 0.05, _
                            shp.Top + shp.Height - 1, _
                            shp.Width * 0.9, 1)
                If Not ln Is Nothing Then
                    ln.Name = ADD_PFX & "BADGERULE" & Format$(seq, "00")
                    With ln.Fill
                        .Visible = msoTrue: .Solid
                        .ForeColor.RGB = cStone
                        .Transparency = 0.2
                    End With
                    ln.Line.Visible = msoFalse
                    ln.Shadow.Visible = msoFalse
                    ln.Placement = xlFreeFloating
                End If
            End If
        End If
    Next shp
    On Error GoTo 0
End Sub

' Section banners are blunt monoliths: square, flat, a heavy sandstone
' rule along the bottom edge like the shadow line where a slab meets the
' sand. No shadow of their own - four full-width objects all casting
' shadows would make the sheet look like it was peeling.
Private Sub StyleBanner(ByVal shp As Shape)
    On Error Resume Next
    shp.AutoShapeType = SHP_SLAB
    With shp.Fill
        .Visible = msoTrue: .Solid
        .ForeColor.RGB = cSlab
        .Transparency = 0
    End With
    With shp.Line
        .Visible = msoTrue
        .ForeColor.RGB = cSeam
        .Weight = 0.75
    End With
    shp.Glow.Radius = 0
    shp.Shadow.Visible = msoFalse
    shp.SoftEdge.Type = 0
    If shp.TextFrame.HasText Then
        With shp.TextFrame.Characters.Font
            .Name = "Segoe UI"
            .Color = cSand
            .Bold = True
            .Italic = False
        End With
        ClearTextGlow shp
        SetTracking shp, TRACK_BANNER
    End If
    On Error GoTo 0
End Sub

' Any leftover title shape that is not ours - stripped bare so our own
' AddDuneTitle is the only one visible.
Private Sub StyleTitle(ByVal shp As Shape)
    On Error Resume Next
    shp.Fill.Visible = msoFalse
    shp.Line.Visible = msoFalse
    shp.Glow.Radius = 0
    shp.Shadow.Visible = msoFalse
    shp.SoftEdge.Type = 0
    If shp.TextFrame.HasText Then
        With shp.TextFrame.Characters.Font
            .Name = "Segoe UI"
            .Color = cBone
            .Bold = False
            .Italic = False
        End With
        ClearTextGlow shp
        SetTracking shp, TRACK_TITLE
    End If
    On Error GoTo 0
End Sub

Private Sub ClearTextGlow(ByVal shp As Shape)
    On Error Resume Next
    shp.TextFrame2.TextRange.Font.Glow.Radius = 0
    On Error GoTo 0
End Sub

Private Sub SetTracking(ByVal shp As Shape, ByVal pts As Single)
    On Error Resume Next
    shp.TextFrame2.TextRange.Font.Spacing = pts
    On Error GoTo 0
End Sub

'=====================================================================
' BACKUP / RESTORE - three sheets, keyed by sheet as well as address.
'   A sheet | B address | C ColorIndex | D Color | E Font.Color | F Italic
' A row whose address is "#TAB" carries that sheet's tab colour instead.
'=====================================================================
Private Sub BackupSheetRange(ByVal ws As Worksheet, ByVal addr As String)
    On Error Resume Next
    Dim bak As Worksheet: Set bak = GetBak(True)
    If bak Is Nothing Then Exit Sub
    bak.Visible = xlSheetVeryHidden
    If SheetIsBackedUp(bak, ws.Name) Then Exit Sub

    Dim rng As Range: Set rng = ws.Range(addr)
    Dim n As Long: n = rng.Cells.count
    If n = 0 Then Exit Sub

    Dim arr() As Variant
    ReDim arr(1 To n + 1, 1 To 6)

    Dim c As Range, i As Long
    i = 0
    For Each c In rng.Cells
        i = i + 1
        arr(i, 1) = ws.Name
        arr(i, 2) = c.Address
        arr(i, 3) = c.Interior.ColorIndex
        arr(i, 4) = c.Interior.Color
        arr(i, 5) = c.Font.Color
        arr(i, 6) = c.Font.Italic
    Next c

    i = i + 1
    arr(i, 1) = ws.Name
    arr(i, 2) = "#TAB"
    arr(i, 3) = ws.Tab.ColorIndex
    arr(i, 4) = ws.Tab.Color
    arr(i, 5) = 0
    arr(i, 6) = False

    bak.Cells(NextBakRow(bak), 1).Resize(i, 6).Value = arr
    On Error GoTo 0
End Sub

Private Function NextBakRow(ByVal bak As Worksheet) As Long
    NextBakRow = 1
    On Error Resume Next
    If Len(CStr(bak.Range("A1").Value)) > 0 Then
        NextBakRow = bak.Cells(bak.Rows.count, 1).End(xlUp).Row + 1
    End If
    On Error GoTo 0
End Function

Private Function SheetIsBackedUp(ByVal bak As Worksheet, ByVal nm As String) As Boolean
    On Error Resume Next
    Dim last As Long
    If Len(CStr(bak.Range("A1").Value)) = 0 Then Exit Function
    last = bak.Cells(bak.Rows.count, 1).End(xlUp).Row
    SheetIsBackedUp = Not IsError(Application.Match(nm, bak.Range(bak.Cells(1, 1), bak.Cells(last, 1)), 0))
    On Error GoTo 0
End Function

Private Sub RestoreAllCells()
    On Error Resume Next
    Dim bak As Worksheet: Set bak = GetBak(False)
    If bak Is Nothing Then Exit Sub
    If Len(CStr(bak.Range("A1").Value)) = 0 Then Exit Sub

    Dim last As Long, data As Variant, i As Long
    last = bak.Cells(bak.Rows.count, 1).End(xlUp).Row
    data = bak.Range(bak.Cells(1, 1), bak.Cells(last, 6)).Value

    Dim ws As Worksheet, curName As String
    For i = 1 To last
        If CStr(data(i, 1)) <> curName Then
            curName = CStr(data(i, 1))
            Set ws = Nothing
            Set ws = ThisWorkbook.Sheets(curName)
            If Not ws Is Nothing Then
                ws.Unprotect Password:="p7ss"
                ws.Cells.Borders(xlEdgeBottom).LineStyle = xlNone
                ws.Cells.Borders(xlEdgeLeft).LineStyle = xlNone
                ws.Cells.Borders(xlEdgeRight).LineStyle = xlNone
                ws.Cells.Borders(xlEdgeTop).LineStyle = xlNone
                ws.Cells.Borders(xlInsideHorizontal).LineStyle = xlNone
                ws.Cells.Borders(xlInsideVertical).LineStyle = xlNone
            End If
        End If
        If ws Is Nothing Then GoTo NextRow

        If CStr(data(i, 2)) = "#TAB" Then
            If data(i, 3) = xlColorIndexNone Then
                ws.Tab.ColorIndex = xlColorIndexNone
            Else
                ws.Tab.Color = data(i, 4)
            End If
        Else
            With ws.Range(CStr(data(i, 2)))
                If data(i, 3) = xlNone Then
                    .Interior.ColorIndex = xlNone
                Else
                    .Interior.Color = data(i, 4)
                End If
                .Font.Color = data(i, 5)
                .Font.Italic = data(i, 6)
            End With
        End If
NextRow:
    Next i

    bak.Cells.Clear
    bak.Visible = xlSheetVeryHidden
    On Error GoTo 0
End Sub

'---- stash / restore (per-shape, in its own AltText) ------------------
Private Sub StashOriginal(ByVal shp As Shape)
    On Error Resume Next
    If Left$(GetAlt(shp), Len(TAG)) = TAG Then Exit Sub

    Dim ht As Long, fn As String, fs As Single, fb As Long, fi As Long, fc As Long
    ht = 0: fn = "": fs = 11: fb = 0: fi = 0: fc = 0
    If shp.TextFrame.HasText Then
        ht = -1
        With shp.TextFrame.Characters.Font
            fn = .Name: fs = .Size: fb = .Bold: fi = .Italic: fc = .Color
        End With
    End If

    ' Field 14 is the paragraph alignment. Desert is the only theme that
    ' CHANGES alignment - it left-aligns button captions - so it is the
    ' only one that has to be able to put it back. Without this, removing
    ' Desert or switching to Dark Blue would leave every caption stranded
    ' on the left, because nothing else on the sheet ever sets it.
    '
    ' Appending rather than inserting keeps older stashes readable: the
    ' restore below only reads field 14 when it is actually present.
    Dim al As Long
    al = msoAlignCenter
    On Error Resume Next
    al = shp.TextFrame2.TextRange.ParagraphFormat.Alignment
    On Error Resume Next

    Dim s As String
    s = TAG & shp.AutoShapeType & "|" & shp.Fill.Visible & "|" & _
        shp.Fill.ForeColor.RGB & "|" & CLng(shp.Fill.Transparency * 1000) & "|" & _
        shp.Line.Visible & "|" & shp.Line.ForeColor.RGB & "|" & CLng(shp.Line.Weight * 100) & "|" & _
        ht & "|" & fn & "|" & CLng(fs * 10) & "|" & fb & "|" & fi & "|" & fc & "|" & al
    SetAlt shp, s
    On Error GoTo 0
End Sub

Private Sub RestoreOriginal(ByVal shp As Shape)
    On Error Resume Next
    Dim p() As String
    p = Split(GetAlt(shp), "|")
    If UBound(p) < 7 Then Exit Sub

    shp.Glow.Radius = 0
    shp.TextFrame2.TextRange.Font.Glow.Radius = 0
    shp.TextFrame2.TextRange.Font.Spacing = 0
    shp.SoftEdge.Type = 0
    shp.Shadow.Visible = msoFalse

    If CLng(p(1)) > 0 Then shp.AutoShapeType = CLng(p(1))
    If CLng(p(2)) = msoTrue Then
        shp.Fill.Visible = msoTrue: shp.Fill.Solid
        shp.Fill.ForeColor.RGB = CLng(p(3))
        shp.Fill.Transparency = CLng(p(4)) / 1000
    Else
        shp.Fill.Visible = msoFalse
    End If
    If CLng(p(5)) = msoTrue Then
        shp.Line.Visible = msoTrue
        shp.Line.ForeColor.RGB = CLng(p(6))
        shp.Line.Weight = CLng(p(7)) / 100
    Else
        shp.Line.Visible = msoFalse
    End If

    If UBound(p) >= 13 Then
        If CLng(p(8)) <> 0 And Len(p(9)) > 0 Then
            With shp.TextFrame.Characters.Font
                .Name = p(9)
                .Size = CLng(p(10)) / 10
                .Bold = (CLng(p(11)) <> 0)
                .Italic = (CLng(p(12)) <> 0)
                .Color = CLng(p(13))
            End With
        End If
    End If

    ' field 14 - only present on stashes written since alignment was added
    If UBound(p) >= 14 Then
        On Error Resume Next
        shp.TextFrame2.TextRange.ParagraphFormat.Alignment = CLng(p(14))
        On Error Resume Next
    End If

    SetAlt shp, ""
    On Error GoTo 0
End Sub

'=====================================================================
' COLOUR PROBE - reads back what is ACTUALLY on the sheet.
'
' The sidebar keeps rendering as pale khaki when it is being painted
' #2A241D, which is nearly black. I have guessed at the cause twice from
' photographs and fixed two real bugs without the symptom going away, so
' this stops guessing: run it after applying Desert and it reports what
' Excel says the cells and shapes are, not what the code asked for.
'
' If the reported values ARE the palette, the paint is landing and the
' problem is display - screen colour profile, or a phone camera's white
' balance on a dark room. If they are not, the number it reports says
' what is overriding them.
'
' Alt+F8 > DesertColourProbe.
'=====================================================================
Public Sub DesertColourProbe()
    Dim ws As Worksheet, m As String, shp As Shape, n As Long
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets(SHEET_NAME)
    If ws Is Nothing Then Exit Sub
    InitPalette

    m = "EXPECTED" & vbCrLf & _
        "  sidebar   " & Hx(cSlab) & vbCrLf & _
        "  canvas    " & Hx(cNight) & vbCrLf & _
        "  data row  " & Hx(cRock) & vbCrLf & vbCrLf & "ACTUAL" & vbCrLf

    m = m & "  A5  fill " & Hx(ws.Range("A5").Interior.Color) & _
            "   pattern " & ws.Range("A5").Interior.Pattern & _
            "   tint " & Format$(ws.Range("A5").Interior.TintAndShade, "0.00") & vbCrLf
    m = m & "  C20 fill " & Hx(ws.Range("C20").Interior.Color) & _
            "   pattern " & ws.Range("C20").Interior.Pattern & _
            "   tint " & Format$(ws.Range("C20").Interior.TintAndShade, "0.00") & vbCrLf
    m = m & "  W3  fill " & Hx(ws.Range("W3").Interior.Color) & "   (canvas)" & vbCrLf
    m = m & "  H6  fill " & Hx(ws.Range("H6").Interior.Color) & "   (data row)" & vbCrLf & vbCrLf

    For Each shp In ws.Shapes
        If Len(shp.OnAction) > 0 And n < 2 Then
            n = n + 1
            m = m & "  btn '" & Left$(shp.TextFrame.Characters.Text, 12) & "' fill " & _
                Hx(shp.Fill.ForeColor.RGB) & "  line " & Hx(shp.Line.ForeColor.RGB) & _
                "  bevel " & shp.ThreeD.BevelTopType & vbCrLf
        End If
    Next shp

    MsgBox m, vbInformation, "Desert colour probe"
    On Error GoTo 0
End Sub

' VBA colours are &HBBGGRR - flipped here to the #RRGGBB people read.
Private Function Hx(ByVal c As Long) As String
    On Error Resume Next
    Hx = "#" & Right$("0" & Hex$(c Mod 256), 2) & _
               Right$("0" & Hex$((c \ 256) Mod 256), 2) & _
               Right$("0" & Hex$((c \ 65536) Mod 256), 2)
    On Error GoTo 0
End Function

'---- helpers ---------------------------------------------------------
Private Function GetAlt(ByVal shp As Shape) As String
    On Error Resume Next
    GetAlt = shp.AlternativeText
    On Error GoTo 0
End Function

Private Sub SetAlt(ByVal shp As Shape, ByVal s As String)
    On Error Resume Next
    shp.AlternativeText = s
    On Error GoTo 0
End Sub

Private Function GetBak(ByVal createIfMissing As Boolean) As Worksheet
    On Error Resume Next
    Set GetBak = ThisWorkbook.Sheets(BAK)
    ' A workbook themed before the rename holds its only copy of the
    ' original cell colours in _DuneBak. Without this fallback the theme
    ' would create an empty _DesertBak, restore nothing, and leave the
    ' sheet permanently dark with the real colours stranded next door.
    If GetBak Is Nothing Then Set GetBak = ThisWorkbook.Sheets(LEGACY_BAK)
    On Error GoTo 0
    If GetBak Is Nothing And createIfMissing Then
        On Error Resume Next
        ThisWorkbook.Unprotect Password:="p7ss"
        Set GetBak = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.count))
        GetBak.Name = BAK
        GetBak.Visible = xlSheetVeryHidden
        On Error GoTo 0
    End If
End Function
