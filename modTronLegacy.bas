Attribute VB_Name = "modTronLegacy"
Option Explicit
'=====================================================================
' modTronLegacy - "Tron: Legacy" restyle for the Sheet1 dashboard.
'
' DO NOT call ApplyTronLegacy directly from a button. Route through
' modThemeManager.ApplyTron, which guarantees any other theme is fully
' removed first. Applying one theme on top of another makes the second
' theme stash the FIRST theme's colours as "the original", and neither
' can be cleanly removed afterwards.
'
'   ApplyTronLegacy   - Grid-black canvas, unlit sidebar slab, circuit
'                       traces, native Office Glow on lit objects
'   RemoveTronLegacy  - restores cells from the hidden backup sheet and
'                       every shape from its own AltText stash
'
' PALETTE: the cyan is #6FC3DF, the screen-accurate Tron Legacy blue -
' cold and desaturated, NOT #00FFFF. Pure neon cyan is what makes a
' "Tron" theme read as a gaming skin. The brightness comes from the
' white-hot core (#F2FEFF) on text and the bloom (#A8E8F9) in the Glow,
' layered over an almost-black ground.
'
' LIT vs UNLIT is the organising rule: only things you can click get a
' lit cyan edge. Panels, banners and rules use the dim trace colour, so
' the eye lands on the controls.
'=====================================================================
Private Const SHEET_NAME As String = "Sheet1"
Private Const TAG As String = "TRONORIG|"
Private Const ADD_PFX As String = "TRONADD_"
Private Const FOREIGN_PFX As String = "NGADD_"     ' modNavyGold's decorations
Private Const BAK As String = "_TronBak"
' Sheet1's merge map runs past row 30 - there are merged blocks at rows
' 29-34 and a full-width A35:T35. Navy & Gold's A1:AC30 canvas stops
' short of them, so those rows keep the default white and show as a bright
' strip under a dark dashboard. This canvas covers them.
Private Const CANVAS As String = "A1:AC36"
Private Const CORNER As Single = 0.08               ' size of the corner cut

' CHAMFERED CORNERS - the single most Tron thing available here.
'
' The film's interfaces (and its hardware, and its architecture) are built
' from rectangles with CUT corners, never rounded ones. A rounded button
' reads as a generic dark UI no matter what colour it glows; a snipped
' corner reads as the Grid immediately. It costs one property per shape.
'
' Numeric rather than the named MsoAutoShapeType constants on purpose: if
' an Office build's type library lacks a name the module fails to COMPILE,
' whereas an unsupported number just fails under On Error and the shape
' keeps the type it already had. Wrong-looking beats not running.
'   155 = msoShapeSnip1Rectangle       one cut corner, top-right
'   157 = msoShapeSnip2DiagRectangle   two cut corners, opposite
Private Const SHP_SNIP1 As Long = 155
Private Const SHP_SNIP2DIAG As Long = 157

' Tron's UI lettering is wide-tracked. Excel cannot letter-space cell
' text, but TextFrame2 exposes Font.Spacing on SHAPES, which is where all
' the labels live - so the banners and badges can carry it. Tracking is
' what makes uppercase read as an interface label rather than a heading.
Private Const TRACK_BANNER As Single = 2#
Private Const TRACK_BADGE As Single = 2.5

' --- palette (set by InitPalette) ---
Private cVoid As Long          ' Grid floor - near-black with a blue bias
Private cPanel As Long         ' raised panel / data cell
Private cPanelLift As Long     ' gradient partner for cPanel - gives buttons depth
Private cSlab As Long          ' sidebar slab - the unlit block light sits on
Private cTrace As Long         ' unlit circuit trace
Private cCyan As Long          ' lit circuit - the icy Tron blue
Private cBloom As Long         ' glow bloom / primary action edge
Private cCore As Long          ' white-hot core (text)
Private cDim As Long           ' secondary / de-emphasised text
Private cOrange As Long        ' Reset button ONLY - the sheet's one accent
Private cAmber As Long         ' its halo

' Glow.Transparency runs 0 (solid halo) to 1 (invisible), so a stronger
' glow is a LOWER number.
'
' THE RULE: on a dark sheet, glow is a scarce resource. Roughly ten
' objects on Sheet1 should emit light - the title, eight buttons, and the
' two panel badges. Everything else is dark structure. When banners, table
' borders and body text all glow too, the contrast that made the buttons
' findable is gone and the sheet becomes an even wall of cyan. That is
' what happened at radius 24 with transparency 0.05 on everything, and it
' is why these values are now conservative and reserved.
'
' Font.Glow is deliberately NOT used below 12pt. A halo on 8.5pt text
' does not make it glow, it smears it - the halo is wider than the stroke,
' so the letterforms fill in and the caption turns to mush. Only the
' title, which is large, gets lit text.
'
' These live in the declarations section because VBA accepts module-level
' Const only before the first procedure; one placed after an End Sub fails
' the whole module with "Only comments may appear after End Sub".
Private Const GLOW_TITLE_R As Single = 16       ' the title is the sheet's brightest object
Private Const GLOW_TITLE_T As Single = 0.3
Private Const GLOW_TITLE_TEXT_R As Single = 8   ' the ONLY lit text on the sheet
Private Const GLOW_HERO_R As Single = 12        ' Start / Reset
Private Const GLOW_HERO_T As Single = 0.35
Private Const GLOW_BTN_R As Single = 7
Private Const GLOW_BTN_T As Single = 0.55
Private Const GLOW_BADGE_R As Single = 6        ' ACTION / UTILITY badges
Private Const GLOW_BADGE_T As Single = 0.6

Private Sub InitPalette()
    ' SURFACES - one blue, three steps, nothing else.
    '
    ' Everything that is a surface uses one of exactly these three values,
    ' and they are the same hue at different lightness. The previous pass
    ' also put a GRADIENT on every button and banner, so each object ran
    ' from #162834 at its top to #04070A at its bottom - meaning no two
    ' objects on the sheet ever shared a colour, and the whole dashboard
    ' read as a muddled blue wash. Flat fills are what make it consistent.
    cVoid = RGB(6, 16, 24)          ' #061018  ground / canvas
    cPanel = RGB(11, 24, 35)        ' #0B1823  data rows - one step up
    cSlab = RGB(14, 30, 44)         ' #0E1E2C  EVERY shape: banners, buttons, badges
    cPanelLift = cSlab              ' kept for compatibility; no gradients any more

    ' Light sources. These are emitters, so unlike the surfaces they are
    ' bright and cold - the icy near-white blue of a lit circuit.
    ' Unlit circuit. This is a SEPARATOR, not a light - it draws every
    ' table rule on the sheet, so at #26829C it turned the data into a
    ' glowing neon grid that competed with the buttons. Dark enough to
    ' read as structure.
    cTrace = RGB(24, 56, 72)        ' #183848  unlit circuit
    cCyan = RGB(122, 214, 245)      ' #7AD6F5  lit circuit
    cBloom = RGB(180, 240, 255)     ' #B4F0FF  its halo
    cCore = RGB(240, 253, 255)      ' #F0FDFF  white-hot core
    cDim = RGB(118, 162, 178)       ' #76A2B2  secondary text

    ' ORANGE IS FOR RESET AND NOTHING ELSE.
    '
    ' It was previously also the Utility faction colour - the UTILITY
    ' badge and the lower half of the sidebar light run. That gave the
    ' sheet two competing accents and made Reset stop reading as the one
    ' destructive control, which is the only job the colour has here. One
    ' orange object on the sheet means it can never be missed.
    ' The edge runs hot red-orange, the halo around it runs yellower, so
    ' the single object still reads as a light rather than a flat swatch.
    cOrange = RGB(255, 78, 26)      ' #FF4E1A  Reset edge
    cAmber = RGB(255, 160, 51)      ' #FFA033  its halo
End Sub

'--------------------------------------------------------------------
Sub ApplyTronLegacy()
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

    ' Order matters. Clear every inherited artefact BEFORE stashing any
    ' shape, or the stash records another theme's colours as "original".
    ' Sheet1 has NO title shape of its own. Reading the drawing XML, the
    ' only "Beta 3.6" on the sheet is NGADD_TITLE_TXT - a shape Navy &
    ' Gold generates. Purging NGADD_* therefore deletes the sheet's only
    ' title outright, and nothing un-hides it because there is nothing
    ' hidden. So capture the caption first and re-issue it as our own.
    Dim titleText As String
    titleText = CaptureTitleText(ws)

    PurgeAllSheetBackgroundImages ThisWorkbook   ' the mottled wash behind the data
    UnhideBaselineShapes ws                      ' recover anything a theme hid
    HideForeignGraphics ws                       ' the gold shield - see below
    PurgeForeignDecor ws                         ' NavyGold folds / cards / title
    RemoveAdded ws                               ' our own additions from a previous run
    ApplyCells ws                                ' Grid-black canvas + sidebar slab

    For Each shp In ws.Shapes
        If IsStylable(shp) Then
            StashOriginal shp
            Select Case ShapeRole(shp)
                Case "TITLE":     StyleTitle shp
                Case "PRIMARY":   StyleButton shp, cBloom, cBloom, True    ' Start
                Case "DANGER":    StyleButton shp, cOrange, cAmber, True   ' Reset
                Case "BUTTON":    StyleButton shp, cCyan, cBloom, False
                ' Both badges are cyan. Orange is reserved for Reset.
                Case "BADGE":     StyleBadge shp, cCyan, cBloom            ' ACTION
                Case "BADGE_ALT": StyleBadge shp, cCyan, cBloom            ' UTILITY
                Case Else:        StyleBanner shp                          ' section headers
            End Select
            n = n + 1
        End If
    Next shp

    AddGroupCards ws                 ' behind the buttons - restores sidebar structure
    AddCircuitTraces ws
    AddCircuitTaps ws                ' the spine feeds each panel - see below
    AddTronTitle ws, titleText       ' after the cards, so it sits above them

    ' The theme is not a Sheet1 theme. An analyst moves between these
    ' three tabs constantly, and leaving two of them on the stock white
    ' grid made the dashboard look like the odd one out rather than the
    ' product looking themed.
    ApplySearchMatrix
    ApplyBackendSettings

    On Error Resume Next
    ActiveWindow.DisplayGridlines = False
    On Error GoTo 0

    If wasProt Then ws.Protect Password:="p7ss"
    If wbProt Then ThisWorkbook.Protect Password:="p7ss", Structure:=True
    Application.ScreenUpdating = True
    MsgBox "Tron Legacy applied - canvas, " & n & " shape(s) and circuit traces.", _
           vbInformation, "Tron Legacy"
End Sub

Sub RemoveTronLegacy()
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
        If Left$(GetAlt(shp), Len(TAG)) = TAG Then
            RestoreOriginal shp
            n = n + 1
        End If
    Next shp
    RestoreAllCells                  ' all three sheets, from one table

    On Error Resume Next
    ActiveWindow.DisplayGridlines = True
    ShowGridlines "Search Matrix"
    ShowGridlines "Backend_Settings"
    On Error GoTo 0

    If wasProt Then ws.Protect Password:="p7ss"
    If wbProt Then ThisWorkbook.Protect Password:="p7ss", Structure:=True
    Application.ScreenUpdating = True
    MsgBox "Tron Legacy removed; canvas and " & n & " shape(s) restored.", _
           vbInformation, "Tron Legacy"
End Sub

'---- INHERITED-ARTEFACT CLEANUP -------------------------------------
' A sheet background image set by an earlier theme survives every cell
' fill - Interior.Color paints OVER it only where cells are opaque, so
' it shows through as a mottled wash. It must be deleted, not covered.
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

' A theme may HIDE a real dashboard shape and draw its own replacement
' rather than restyling the original. modNavyGold does exactly that with
' the Beta title: RepositionBeta sets the workbook's own title shape to
' Visible = msoFalse, then AddTitleBeta draws NGADD_SHIELD_OUTER,
' NGADD_SHIELD_CHK and NGADD_TITLE_TXT in its place.
'
' RemoveNavyGold never sets that title back to visible - it deletes its
' own NGADD_ shapes and stops - so the real title stays hidden even after
' Navy & Gold is removed. Purging NGADD_ shapes without this step would
' therefore delete the ONLY visible title on the sheet and leave the nav
' bar with no heading at all.
'
' Any shape hidden here is either that title or a leftover of the same
' kind. Un-hiding one shape too many is visible and trivially undone;
' silently losing the title is neither. So this un-hides broadly, then
' lets the normal styling pass treat the recovered shape as the title.
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

' Reads the current title caption off whatever shape is carrying it, so
' the version number follows the workbook instead of being hardcoded here
' and going stale at the next bump.
Private Function CaptureTitleText(ByVal ws As Worksheet) As String
    On Error Resume Next
    Dim shp As Shape, t As String

    ' Fall back to the workbook's own filename, not the bare word "Beta".
    ' Once a previous run has deleted the title shape there is no caption
    ' left on the sheet to read, and every run after that would render a
    ' title reading just "Beta" - which is exactly what happened. The
    ' filename always carries the real version.
    CaptureTitleText = ThisWorkbook.Name
    Dim dot As Long
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

' Sheet1 carries one msoGraphic ("Graphic 2") - the gold shield that sits
' beside the title. It is not an AutoShape, so no theme can recolour it,
' and it is not NGADD_-prefixed, so purging Navy & Gold's decorations
' leaves it behind: a gold icon stranded on a black Grid.
'
' It is hidden rather than deleted, because deleting is irreversible and
' the icon belongs to the workbook. The theme manager's un-hide pass
' brings it straight back when a theme is removed.
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

' modNavyGold generates its own decorations - corner folds, panel cards,
' the shield-and-title lockup, inline icon glyphs - all named NGADD_*.
' They are Navy & Gold objects, not dashboard content, so Tron deletes
' them rather than restyling them. ApplyNavyGold regenerates every one of
' them, so this is not destructive - but it is only safe AFTER
' UnhideBaselineShapes has brought the real title back.
Private Sub PurgeForeignDecor(ByVal ws As Worksheet)
    On Error Resume Next
    Dim i As Long
    For i = ws.Shapes.count To 1 Step -1
        If Left$(ws.Shapes(i).Name, Len(FOREIGN_PFX)) = FOREIGN_PFX Then ws.Shapes(i).Delete
    Next i
    On Error GoTo 0
End Sub

'---- CELLS -----------------------------------------------------------
Private Sub ApplyCells(ByVal ws As Worksheet)
    BackupSheetRange ws, CANVAS

    ' 1. The Grid floor - the whole canvas goes black. Light needs dark.
    With ws.Range(CANVAS)
        .Interior.Color = cVoid
        .Font.Color = cDim
        .Borders(xlEdgeBottom).LineStyle = xlNone
        .Borders(xlInsideHorizontal).LineStyle = xlNone
    End With

    ' 2. Sidebar slab - a shade off the floor, so it reads as a solid
    '    object rather than a hole in the background.
    ws.Range("A1:E29").Interior.Color = cSlab

    ' 3. Data bands.
    '
    ' Read from Sheet1's own merge map rather than assumed: every data row
    ' is merged as G:I + J:T (or J:S). That is a real label/value split
    ' the sheet already encodes, so the theme uses it instead of painting
    ' one flat band across G:T - which is exactly why the earlier pass
    ' looked uniform and dead. Labels sit back in the secondary tone,
    ' values come forward white-hot. That contrast, not more colour, is
    ' what makes a dark UI readable.
    '
    ' Rows 17 and 26 are the two in-table COLUMN HEADER rows (CounterParty
    ' #s / Name / Address, and Country Name / ISOCode2 / ...). They are
    ' not data and are styled as headers.
    '
    ' Italic is cleared throughout: Grid typography is upright, and the
    ' inherited italic labels were a Navy & Gold mannerism that read as a
    ' rendering glitch against this palette.
    Dim bands As Variant, addr As Variant, sec As Range, r As Long, rr As Range
    bands = Array("G5:T10", "G13:T14", "G17:T23", "G26:T27")
    For Each addr In bands
        Set sec = ws.Range(CStr(addr))
        For r = sec.Row To sec.Row + sec.Rows.count - 1
            Set rr = ws.Range("G" & r & ":T" & r)
            rr.Font.Name = "Segoe UI"
            rr.Font.Italic = False

            If r = 17 Or r = 26 Then
                ' column-header row - lit, so the table reads as a table
                rr.Interior.Color = cSlab
                rr.Font.Color = cCyan
                rr.Font.Bold = True
            Else
                rr.Interior.Color = cPanel
                ws.Range("G" & r & ":I" & r).Font.Color = cDim     ' label
                ws.Range("G" & r & ":I" & r).Font.Bold = False
                ws.Range("J" & r & ":T" & r).Font.Color = cCore    ' value
            End If

            ' A bottom rule only. Boxing every row on three sides drew a
            ' bright cage around each line of data and made the tables the
            ' loudest thing on the sheet. One horizontal separator is
            ' enough to group rows; the panel fill already defines the
            ' block's edges.
            With rr.Borders(xlEdgeBottom)
                .LineStyle = xlContinuous: .Weight = xlHairline: .Color = cTrace
            End With
        Next r
    Next addr

    ' 4. The theme picker (U28). It sits inside CANVAS, so without this it
    '    takes the plain ground fill and dim text and the one control that
    '    switches themes becomes the least visible thing on the sheet.
    '    Styled as a control: surface colour, lit caption.
    With ws.Range("U28")
        .Interior.Color = cSlab
        .Font.Name = "Segoe UI": .Font.Size = 9
        .Font.Color = cCyan: .Font.Bold = True: .Font.Italic = False
        .HorizontalAlignment = xlCenter
        .Borders.LineStyle = xlContinuous
        .Borders.Weight = xlThin
        .Borders.Color = cCyan
        .Locked = False                 ' the sheet is re-protected on exit
    End With
End Sub

'=====================================================================
' SEARCH MATRIX
'
' Laid out in 5-row blocks, one entity each (Customer, then CP1..CP6), so
' the banding is drawn per BLOCK rather than per row - each entity reads
' as one group instead of the rows striping against the grouping the data
' actually has.
'
' Columns carry different weights: A is the entity title, B and C are
' reference detail that should sit back, D is the naming convention an
' analyst reads, E is the link. Same label/value discipline as Sheet1.
'=====================================================================
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

    ' clean slate - A:E only, the helper columns past F stay hidden.
    ' Alignment is deliberately NOT set anywhere in here: modThemeManager
    ' re-asserts centre/middle across this sheet after every theme, so
    ' setting it would only be overwritten, and would imply alignment is a
    ' theme decision when it is a property of the sheet.
    With ws.Range("A1:E" & lastRow)
        .Interior.Pattern = xlSolid
        .Interior.Color = cPanel
        .Borders.LineStyle = xlNone
        .Font.Name = "Segoe UI"
        .Font.Size = 9
        .Font.Color = cCore
        .Font.Bold = False
        .Font.Italic = False
        .WrapText = False
    End With

    ' header
    With ws.Range("A1:E1")
        .Interior.Color = cSlab
        .Font.Color = cCyan
        .Font.Bold = True
        .Font.Size = 10
    End With
    ws.Rows(1).RowHeight = 24
    With ws.Range("A1:E1").Borders(xlEdgeBottom)
        .LineStyle = xlContinuous: .Weight = xlThin: .Color = cCyan
    End With

    If lastRow >= 2 Then
        With ws.Range("A2:A" & lastRow)
            .Font.Bold = True: .Font.Color = cCore
        End With
        ws.Range("B2:B" & lastRow).Font.Color = cDim
        With ws.Range("C2:C" & lastRow)
            .Font.Size = 8: .Font.Color = cDim
        End With
        ws.Range("D2:D" & lastRow).Font.Color = cCore
        With ws.Range("E2:E" & lastRow)
            .Font.Color = cCyan
            .Font.Underline = xlUnderlineStyleSingle
        End With

        For r = 2 To lastRow
            blk = (r - 2) \ 5                       ' 0 = Customer, 1 = CP1, ...
            ws.Range("A" & r & ":E" & r).Interior.Color = _
                IIf(blk Mod 2 = 0, cPanel, cSlab)
            With ws.Range("A" & r & ":E" & r).Borders(xlEdgeBottom)
                .LineStyle = xlContinuous: .Weight = xlHairline: .Color = cTrace
            End With
            ' the last row of each entity block gets a lit divider, so the
            ' groups read even where two same-shade blocks meet
            If (r - 1) Mod 5 = 0 Then
                With ws.Range("A" & r & ":E" & r).Borders(xlEdgeBottom)
                    .LineStyle = xlContinuous: .Weight = xlThin: .Color = cTrace
                End With
            End If
        Next r
    End If

    FrameRange ws.Range("A1:E" & lastRow)
    KillGridlines ws
    If wasProt Then ws.Protect Password:="p7ss"
    On Error GoTo 0
End Sub

'=====================================================================
' BACKEND_SETTINGS - four visible cells, so it is all title bar and one
' key/value row. Row 3 (Flow_1) is left exactly as it is: Module10 reads
' the PAD merge URL from B4, and hiding or unhiding rows here would shift
' that reference. The theme colours, it does not restructure.
'=====================================================================
Private Sub ApplyBackendSettings()
    On Error Resume Next
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("Backend_Settings")
    If ws Is Nothing Then Exit Sub

    Dim wasProt As Boolean
    wasProt = ws.ProtectContents
    ws.Unprotect Password:="p7ss"

    BackupSheetRange ws, "A1:B4"

    With ws.Range("A1:B4")
        .Interior.Color = cPanel
        .Borders.LineStyle = xlNone
        .Font.Name = "Segoe UI"
        .Font.Color = cCore
        .Font.Size = 10
        .Font.Bold = False
        .Font.Italic = False
        .WrapText = False
    End With

    ' Alignment is deliberately NOT set here. modThemeManager re-asserts
    ' centre/middle on this block after every theme, so setting it in the
    ' theme would only be overwritten - and would give the false
    ' impression that alignment is a theme decision.
    With ws.Range("A1:B1")
        .Interior.Color = cSlab
        .Font.Color = cCyan
        .Font.Bold = True
        .Font.Size = 11
    End With
    ws.Rows(1).RowHeight = 22
    With ws.Range("A1:B1").Borders(xlEdgeBottom)
        .LineStyle = xlContinuous: .Weight = xlThin: .Color = cCyan
    End With

    ws.Range("A4").Font.Bold = True          ' key
    ws.Range("A4").Font.Color = cDim
    ws.Range("B4").Font.Color = cCore        ' value

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
            .LineStyle = xlContinuous: .Weight = xlThin: .Color = cCyan
        End With
    Next e
    On Error GoTo 0
End Sub

' Gridlines are a WINDOW property, so the sheet has to be active to set
' them - hence the activate-and-return dance.
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
' BACKUP / RESTORE - now covers THREE sheets, not one.
'
' The backup table is keyed by sheet as well as address, because the
' theme no longer touches Sheet1 alone:
'
'   A = sheet name   B = cell address   C = Interior.ColorIndex
'   D = Interior.Color   E = Font.Color   F = Font.Italic
'
' A row whose address is "#TAB" carries that sheet's tab colour instead
' of a cell's - it rides in the same table so there is one thing to read
' back and one thing to clear, rather than a cell block plus a separate
' corner of the sheet holding tab state.
'
' Cells are captured into an array and written in ONE go. Row-by-row
' writes of ~1200 cells x 6 columns was the slowest part of applying the
' theme, and it is exactly the kind of thing that crawls on VDI.
'=====================================================================
Private Sub BackupSheetRange(ByVal ws As Worksheet, ByVal addr As String)
    On Error Resume Next
    Dim bak As Worksheet: Set bak = GetBak(True)
    If bak Is Nothing Then Exit Sub
    bak.Visible = xlSheetVeryHidden

    ' Already captured for this sheet - never re-capture, or a second
    ' apply would record the FIRST theme's colours as the original.
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

    ' the tab-colour row
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

' Replays every backed-up sheet, then empties the table.
Private Sub RestoreAllCells()
    On Error Resume Next
    Dim bak As Worksheet: Set bak = GetBak(False)
    If bak Is Nothing Then Exit Sub
    If Len(CStr(bak.Range("A1").Value)) = 0 Then Exit Sub

    Dim last As Long, data As Variant, i As Long
    last = bak.Cells(bak.Rows.count, 1).End(xlUp).Row
    data = bak.Range(bak.Cells(1, 1), bak.Cells(last, 6)).Value

    ' OLD-FORMAT BACKUPS.
    ' The first version of this table was Sheet1-only and had no sheet
    ' column: A held the cell address. Reading one of those with the new
    ' layout would look up a worksheet named "$A$1", find nothing, restore
    ' nothing, and then clear the table - leaving the sheet permanently
    ' dark with its original colours gone. An address always starts with
    ' "$", and a sheet name never can, so the two are trivially told apart.
    Dim legacy As Boolean
    legacy = (Left$(CStr(data(1, 1)), 1) = "$")

    Dim ws As Worksheet, curName As String
    For i = 1 To last
        If legacy Then
            ' shift the row right by one into the new shape
            data(i, 6) = data(i, 5)
            data(i, 5) = data(i, 4)
            data(i, 4) = data(i, 3)
            data(i, 3) = data(i, 2)
            data(i, 2) = data(i, 1)
            data(i, 1) = SHEET_NAME
        End If

        If CStr(data(i, 1)) <> curName Then
            curName = CStr(data(i, 1))
            Set ws = Nothing
            Set ws = ThisWorkbook.Sheets(curName)
            If Not ws Is Nothing Then
                ws.Unprotect Password:="p7ss"
                ' Clear every border the theme could have drawn, across
                ' the whole used area - cheaper and more thorough than
                ' tracking which edges were set.
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

'---- CIRCUIT TRACES --------------------------------------------------
' Thin lit lines with a glow, anchored to cell geometry (not to shapes,
' whose positions we cannot assume). These are what separate "Tron" from
' "someone turned the lights off": the Grid is defined by light running
' along the edges of dark slabs.
Private Sub AddCircuitTraces(ByVal ws As Worksheet)
    On Error Resume Next
    Dim x As Single, ln As Shape

    ' The sidebar's outer edge is the sheet's one continuous light run.
    '
    ' It is a single cyan ribbon now. It used to be cyan over the Action
    ' half and orange over the Utility half, but orange belongs to Reset
    ' alone - a second accent running half the sidebar was competing with
    ' the one control that actually needs to stand out.
    '
    ' It still fades at BOTH ends rather than starting and stopping at
    ' full strength. A light that begins abruptly reads as a drawn line;
    ' one that fades in reads as light. Same principle as before, applied
    ' to the ends instead of to a mid-point join.
    x = ws.Range("F1").Left
    AddLightRibbon ws, x, ws.Range("A1").Top, ws.Range("A30").Top, cCyan, cBloom

    ' No horizontal rule at row 4 any more - it ran straight through the
    ' title's glow and read as a strike-through.
    On Error GoTo 0
End Sub

' Tron's own title. Navy & Gold pairs its title with a vector shield; the
' Grid's own convention is lettering with nothing containing it, so this
' is text alone - white-hot, with the widest halo on the sheet.
Private Sub AddTronTitle(ByVal ws As Worksheet, ByVal caption As String)
    On Error Resume Next
    Dim shp As Shape, badge As Shape
    Dim x As Single, w As Single, y As Single, h As Single

    x = ws.Range("A1").Left + 14
    w = ws.Range("A1:E1").Width - 28
    y = ws.Range("A1").Top + 4

    ' Fit the title into the clear space ABOVE the ACTION badge, the same
    ' way Navy & Gold sizes its own header. Sizing it blind made it
    ' overlap the badge and the rule beneath it.
    Set badge = FindBadge(ws, "ACTION")
    If Not badge Is Nothing Then
        h = (badge.Top - y) - 8              ' 8pt of clear air above the badge
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
        .Size = 13
        .Bold = True
        .Color = cCore
    End With
    With shp.TextFrame2
        .TextRange.ParagraphFormat.Alignment = msoAlignCenter
        .VerticalAnchor = msoAnchorMiddle
        .MarginTop = 0: .MarginBottom = 0
        .WordWrap = msoFalse
    End With
    StyleTitle shp
    shp.ZOrder msoBringToFront
    On Error GoTo 0
End Sub

' Navy & Gold groups the buttons inside two gold cards. Purging those
' leaves the eight buttons floating on a flat slab with nothing tying each
' group together, so Tron draws its own - dark, hairline-edged, unlit, and
' sent to the back. Structure, not decoration: they say which buttons
' belong to ACTION and which to UTILITY.
Private Sub AddGroupCards(ByVal ws As Worksheet)
    On Error Resume Next
    AddOneCard ws, FindBadge(ws, "ACTION"), "A5", "A14", "ACT"
    AddOneCard ws, FindBadge(ws, "UTILITY"), "A18", "A27", "UTL"
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

' The card must start BELOW its badge, not at the badge's own row.
'
' Anchoring it to a hardcoded row was wrong because the badge sits on that
' same row, so the card wrapped up and around the badge instead of
' enclosing only the buttons. Navy & Gold anchors to the badge SHAPE -
' top = badge.Top + badge.Height + 5 - and that is correct, because the
' badge is repositioned at runtime and its row is not fixed. Matching it
' exactly, including the transparent fill: the card is a boundary line,
' not a filled panel, so the sidebar slab shows through and the sheet
' keeps one surface colour.
Private Sub AddOneCard(ByVal ws As Worksheet, ByVal badge As Shape, _
                       ByVal fallbackTop As String, ByVal botCell As String, _
                       ByVal tag As String)
    On Error Resume Next
    Dim shp As Shape, x As Single, w As Single, y As Single, h As Single
    x = ws.Range("A1").Left + 10
    w = ws.Range("A1:E1").Width - 20

    If Not badge Is Nothing Then
        y = badge.Top + badge.Height + 5
    Else
        y = ws.Range(fallbackTop).Top - 1
    End If
    h = (ws.Range(botCell).Top + ws.Range(botCell).Height) - y + 1
    If h < 10 Then Exit Sub                    ' badge not found where expected

    Set shp = ws.Shapes.AddShape(msoShapeRoundedRectangle, x, y, w, h)
    If shp Is Nothing Then Exit Sub
    shp.Name = ADD_PFX & "CARD_" & tag
    shp.Adjustments(1) = 0.06
    shp.Fill.Visible = msoFalse                ' boundary only - slab shows through
    With shp.Line
        .Visible = msoTrue
        .ForeColor.RGB = cTrace
        .Weight = 1#
    End With
    shp.Glow.Radius = 0
    shp.Shadow.Visible = msoFalse
    shp.ZOrder msoSendToBack
    On Error GoTo 0
End Sub

' CIRCUIT TAPS - what turns a dark sidebar into a Grid.
'
' In Tron, light is never just an outline round a thing. It is POWER, and
' it runs along routed paths: a spine, taps branching off it, and a bright
' node at every junction. That reading - "this panel is connected to that
' line, and the junction is live" - is the film's actual visual grammar,
' and no amount of extra glow substitutes for it.
'
' So the sidebar ribbon becomes a spine, and each panel card is tapped off
' it by a short trace ending in a node square sitting ON the spine. Two
' taps, four small shapes. It costs almost nothing and it is the detail
' that stops this reading as a generic dark theme.
Private Sub AddCircuitTaps(ByVal ws As Worksheet)
    On Error Resume Next
    TapToCard ws, ADD_PFX & "CARD_ACT"
    TapToCard ws, ADD_PFX & "CARD_UTL"
    On Error GoTo 0
End Sub

Private Sub TapToCard(ByVal ws As Worksheet, ByVal cardName As String)
    On Error Resume Next
    Dim card As Shape, ln As Shape, node As Shape
    Dim spineX As Single, y As Single, xEnd As Single

    Set card = Nothing
    Set card = ws.Shapes(cardName)
    If card Is Nothing Then Exit Sub

    spineX = ws.Range("F1").Left
    y = card.Top + card.Height / 2
    xEnd = card.Left + card.Width
    If xEnd >= spineX Then Exit Sub          ' card overlaps the spine - no room

    Static seq As Long
    seq = seq + 1

    ' the trace
    Set ln = ws.Shapes.AddLine(xEnd, y, spineX, y)
    If Not ln Is Nothing Then
        ln.Name = ADD_PFX & "TAP" & Format$(seq, "00")
        With ln.Line
            .Visible = msoTrue
            .ForeColor.RGB = cTrace
            .Weight = 1#
        End With
        ln.Placement = xlFreeFloating
    End If

    ' the junction node, centred on the spine and lit - this is the bit
    ' that reads as a live connection rather than a stray line
    Const N As Single = 5
    Set node = ws.Shapes.AddShape(msoShapeRectangle, spineX - N / 2, y - N / 2, N, N)
    If Not node Is Nothing Then
        node.Name = ADD_PFX & "NODE" & Format$(seq, "00")
        With node.Fill
            .Visible = msoTrue: .Solid
            .ForeColor.RGB = cCyan
        End With
        node.Line.Visible = msoFalse
        With node.Glow
            .Color.RGB = cBloom
            .Radius = 6
            .Transparency = 0.35
        End With
        node.Placement = xlFreeFloating
    End If
    On Error GoTo 0
End Sub

' The sidebar light run: a thin filled strip that fades in from
' transparent at the top, holds, and fades back out at the bottom.
'
' A gradient FILL on a narrow rectangle is used rather than a line,
' because Excel's line format cannot carry a gradient - a line is one flat
' colour, so it can only ever start and stop abruptly. Only the ALPHA
' ramps here; the hue is constant, so the ribbon never passes through a
' muddy intermediate colour on its way to transparent.
Private Sub AddLightRibbon(ByVal ws As Worksheet, ByVal x As Single, _
                           ByVal yTop As Single, ByVal yBot As Single, _
                           ByVal clr As Long, ByVal halo As Long)
    On Error Resume Next
    Const W As Single = 2.5
    Dim shp As Shape
    Set shp = ws.Shapes.AddShape(msoShapeRectangle, x - W / 2, yTop, W, yBot - yTop)
    If shp Is Nothing Then Exit Sub

    shp.Name = ADD_PFX & "RIBBON"
    shp.Line.Visible = msoFalse

    With shp.Fill
        .Visible = msoTrue
        .TwoColorGradient msoGradientVertical, 1
        .ForeColor.RGB = clr
        .BackColor.RGB = clr
        .GradientStops.Insert clr, 0#, 1        ' transparent at the top
        .GradientStops.Insert clr, 0.14, 0      ' full strength
        .GradientStops.Insert clr, 0.88, 0
        .GradientStops.Insert clr, 1#, 1        ' transparent at the bottom
    End With

    With shp.Glow
        .Color.RGB = halo
        .Radius = 9
        .Transparency = 0.5
    End With
    shp.Shadow.Visible = msoFalse
    shp.Placement = xlFreeFloating
    On Error GoTo 0
End Sub

Private Sub LightTrace(ByVal shp As Shape, ByVal w As Single, ByVal clr As Long, _
                       ByVal halo As Long, ByVal glowR As Single)
    On Error Resume Next
    Static seq As Long
    seq = seq + 1
    shp.Name = ADD_PFX & Format$(seq, "000")
    With shp.Line
        .Visible = msoTrue
        .ForeColor.RGB = clr
        .Weight = w
        .Transparency = 0
    End With
    With shp.Glow
        .Color.RGB = halo
        .Radius = glowR
        .Transparency = 0.25
    End With
    shp.Placement = xlFreeFloating
    On Error GoTo 0
End Sub

Private Sub RemoveAdded(ByVal ws As Worksheet)
    On Error Resume Next
    Dim i As Long
    For i = ws.Shapes.count To 1 Step -1
        If Left$(ws.Shapes(i).Name, Len(ADD_PFX)) = ADD_PFX Then ws.Shapes(i).Delete
    Next i
    On Error GoTo 0
End Sub

'---- classification --------------------------------------------------
Private Function IsStylable(ByVal shp As Shape) As Boolean
    On Error Resume Next
    IsStylable = (shp.Type = msoAutoShape Or shp.Type = msoFreeform)
    If Left$(shp.Name, Len(ADD_PFX)) = ADD_PFX Then IsStylable = False
    On Error GoTo 0
End Function

' Returns TITLE / PRIMARY / DANGER / BUTTON / PANEL.
'
' The title test matches "Beta" ANYWHERE in the text, not just at
' position 1. modNavyGold prepends an icon glyph to shape captions, so
' an anchored test silently failed on a themed workbook and the title
' fell through to the button branch - which is exactly why "Beta 3.6"
' came out looking like a control instead of a title.
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
    ElseIf InStr(1, t, "UTILITY", vbTextCompare) > 0 Then
        ShapeRole = "BADGE_ALT"        ' the orange half of the Grid
    ElseIf InStr(1, t, "ACTION", vbTextCompare) > 0 Then
        ShapeRole = "BADGE"
    End If
    On Error GoTo 0
End Function

'---- styling ---------------------------------------------------------
' One button routine, three intensities. Start and Reset are the two
' consequential controls on the sheet, so both are lit harder than the
' rest: Start in the bright bloom cyan, Reset in Rinzler orange. The
' neutral buttons stay at standard cyan and recede.
Private Sub StyleButton(ByVal shp As Shape, ByVal edge As Long, ByVal halo As Long, _
                        ByVal emphasise As Boolean)
    On Error Resume Next
    SetChamfer shp, SHP_SNIP2DIAG          ' cut corners, not rounded

    ' Flat, and the same surface colour as every other shape on the sheet.
    ' A gradient here meant each button ran through two different blues,
    ' so no two objects matched and the dashboard lost its base colour.
    With shp.Fill
        .Visible = msoTrue: .Solid
        .ForeColor.RGB = cSlab
        .Transparency = 0
    End With

    With shp.Line
        .Visible = msoTrue
        .ForeColor.RGB = edge
        .Weight = IIf(emphasise, 2.5, 1.5)
        .Transparency = 0
    End With

    With shp.Glow
        .Color.RGB = halo
        .Radius = IIf(emphasise, GLOW_HERO_R, GLOW_BTN_R)
        .Transparency = IIf(emphasise, GLOW_HERO_T, GLOW_BTN_T)
    End With

    shp.Shadow.Visible = msoFalse
    shp.SoftEdge.Type = 0              ' Grid edges are hard
    If shp.TextFrame.HasText Then
        With shp.TextFrame.Characters.Font
            .Name = "Segoe UI"
            .Bold = emphasise
            .Size = 9
            .Color = cCore
        End With
        ' No text glow at 9pt - see the note on the GLOW_ constants. The
        ' caption has to stay readable; the halo around the shape is what
        ' makes the button look lit.
        ClearTextGlow shp
    End If
    On Error GoTo 0
End Sub

' Font.Glow is a separate object from Shape.Glow, so it has to be cleared
' separately too - otherwise a caption keeps a halo left over from an
' earlier run even after the shape's own glow is gone.
Private Sub LightText(ByVal shp As Shape, ByVal halo As Long, ByVal radius As Single)
    On Error Resume Next
    With shp.TextFrame2.TextRange.Font.Glow
        .Color.RGB = halo
        .Radius = radius
        .Transparency = 0.35
    End With
    On Error GoTo 0
End Sub

Private Sub ClearTextGlow(ByVal shp As Shape)
    On Error Resume Next
    shp.TextFrame2.TextRange.Font.Glow.Radius = 0
    On Error GoTo 0
End Sub

' Swaps a shape to a cut-corner rectangle. Silently leaves the shape as it
' was if this Office build does not know the type.
Private Sub SetChamfer(ByVal shp As Shape, ByVal kind As Long)
    On Error Resume Next
    shp.AutoShapeType = kind
    shp.Adjustments(1) = CORNER
    On Error GoTo 0
End Sub

' Wide-tracked lettering. Set to 0 to clear it on removal.
Private Sub SetTracking(ByVal shp As Shape, ByVal pts As Single)
    On Error Resume Next
    shp.TextFrame2.TextRange.Font.Spacing = pts
    On Error GoTo 0
End Sub

' Banners / section headers / panel badges. Takes its accent so the two
' faction colours can both appear: ACTION and the section banners run
' cyan, UTILITY runs orange. That is what gives the orange real presence
' on the sheet instead of it existing only on the Reset button - and it
' maps onto a split the dashboard already has, so it carries meaning
' rather than being decoration.
' Section banners - Alert Related Information, Customer Information, and
' so on. These are LABELS, not controls. There are four of them spanning
' the full width of the sheet, so lighting them was most of what turned
' the dashboard into a wall of cyan: they are the largest objects present
' and they were glowing as hard as the buttons.
'
' Dark slab, hairline unlit edge, crisp cyan caption, NO glow at all. They
' read as headers because of the fill and the colour, not because they
' emit light.
Private Sub StyleBanner(ByVal shp As Shape)
    On Error Resume Next
    ' One cut corner, top-right - the same position Navy & Gold put its
    ' paper fold, so the eye lands where it expects a corner detail, but
    ' the language is now a chamfer instead of a dog-ear.
    SetChamfer shp, SHP_SNIP1
    With shp.Fill
        .Visible = msoTrue: .Solid
        .ForeColor.RGB = cSlab
        .Transparency = 0
    End With
    With shp.Line
        .Visible = msoTrue
        .ForeColor.RGB = cTrace
        .Weight = 0.75
        .Transparency = 0
    End With
    shp.Glow.Radius = 0
    shp.Shadow.Visible = msoFalse
    shp.SoftEdge.Type = 0
    If shp.TextFrame.HasText Then
        With shp.TextFrame.Characters.Font
            .Name = "Segoe UI"
            .Color = cCyan
            .Bold = True
            .Italic = False
        End With
        ClearTextGlow shp
        SetTracking shp, TRACK_BANNER
    End If
    On Error GoTo 0
End Sub

' ACTION / UTILITY badges. Only two of them, and they head the sidebar, so
' they carry a small halo - enough to tie them to the buttons below
' without competing with them.
Private Sub StyleBadge(ByVal shp As Shape, ByVal edge As Long, ByVal halo As Long)
    On Error Resume Next
    SetChamfer shp, SHP_SNIP2DIAG
    With shp.Fill
        .Visible = msoTrue: .Solid
        .ForeColor.RGB = cSlab
        .Transparency = 0
    End With
    With shp.Line
        .Visible = msoTrue
        .ForeColor.RGB = edge
        .Weight = 1#
        .Transparency = 0
    End With
    With shp.Glow
        .Color.RGB = halo
        .Radius = GLOW_BADGE_R
        .Transparency = GLOW_BADGE_T
    End With
    shp.Shadow.Visible = msoFalse
    shp.SoftEdge.Type = 0
    If shp.TextFrame.HasText Then
        With shp.TextFrame.Characters.Font
            .Name = "Segoe UI"
            .Color = edge
            .Bold = True
            .Italic = False
        End With
        ClearTextGlow shp
        SetTracking shp, TRACK_BADGE
    End If
    On Error GoTo 0
End Sub

' Title - no box at all. White-hot text with a wide cyan bloom, the way
' the film sets its titles: light with nothing containing it.
Private Sub StyleTitle(ByVal shp As Shape)
    On Error Resume Next
    shp.Fill.Visible = msoFalse
    shp.Line.Visible = msoFalse
    shp.SoftEdge.Type = 0
    shp.Shadow.Visible = msoFalse
    With shp.Glow
        .Color.RGB = cBloom
        .Radius = GLOW_TITLE_R
        .Transparency = GLOW_TITLE_T
    End With
    If shp.TextFrame.HasText Then
        With shp.TextFrame.Characters.Font
            .Name = "Segoe UI"
            .Color = cCore
            .Bold = True
            .Italic = False
        End With
        ' The only lit text on the sheet. It can carry a halo because it
        ' is large - at 15pt the glow reads as a glow, where the same
        ' treatment on a 9pt button caption just fills in the letterforms.
        LightText shp, cBloom, GLOW_TITLE_TEXT_R
    End If
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

    Dim s As String
    s = TAG & shp.AutoShapeType & "|" & shp.Fill.Visible & "|" & _
        shp.Fill.ForeColor.RGB & "|" & CLng(shp.Fill.Transparency * 1000) & "|" & _
        shp.Line.Visible & "|" & shp.Line.ForeColor.RGB & "|" & CLng(shp.Line.Weight * 100) & "|" & _
        ht & "|" & fn & "|" & CLng(fs * 10) & "|" & fb & "|" & fi & "|" & fc
    SetAlt shp, s
    On Error GoTo 0
End Sub

Private Sub RestoreOriginal(ByVal shp As Shape)
    On Error Resume Next
    Dim p() As String
    p = Split(GetAlt(shp), "|")
    If UBound(p) < 7 Then Exit Sub

    ' Kill both halos first. Font.Glow is a separate object from
    ' Shape.Glow, so clearing the shape's glow alone leaves the caption
    ' still emitting light after the theme is supposedly gone.
    shp.Glow.Radius = 0
    shp.TextFrame2.TextRange.Font.Glow.Radius = 0
    shp.TextFrame2.TextRange.Font.Spacing = 0    ' undo the wide tracking
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

    SetAlt shp, ""
    On Error GoTo 0
End Sub

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
