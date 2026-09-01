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
Private Const CORNER As Single = 0.12               ' hard-edged Grid panel, not a soft card

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
Private cOrange As Long        ' Rinzler - Reset, Utility faction
Private cAmber As Long         ' the orange side's bloom

' Excel's Glow.Transparency runs 0 (solid halo) to 1 (invisible), so a
' stronger glow means a LOWER number, not a higher one. The first pass
' used 0.45-0.75 with radii of 4-8, which is why nothing appeared to be
' emitting light. These are the values that actually bloom.
'
' These live here, in the module's declarations section, because VBA only
' accepts module-level Const and Dim before the first procedure. Put one
' after an End Sub and the whole module fails to compile with "Only
' comments may appear after End Sub, End Function, or End Property".
Private Const GLOW_TITLE_R As Single = 28
Private Const GLOW_TITLE_T As Single = 0.05
Private Const GLOW_HERO_R As Single = 24        ' Start / Reset
Private Const GLOW_HERO_T As Single = 0.05
Private Const GLOW_BTN_R As Single = 14
Private Const GLOW_BTN_T As Single = 0.25
Private Const GLOW_PANEL_R As Single = 10
Private Const GLOW_PANEL_T As Single = 0.5
Private Const GLOW_TEXT_R As Single = 7         ' Font.Glow - the big win
Private Const GLOW_TEXT_T As Single = 0.15

Private Sub InitPalette()
    ' Surfaces. Frame analysis of the film describes the Grid not as black
    ' but as "very desaturated metallic light blue" - dark, reflective
    ' structure rather than void. So the panels carry a blue-grey cast
    ' instead of being pure near-black, which is also what gives the glow
    ' something to reflect off.
    cVoid = RGB(4, 7, 10)           ' #04070A  deepest ground
    cSlab = RGB(8, 14, 20)          ' #080E14  sidebar slab
    cPanel = RGB(13, 22, 30)        ' #0D161E  metallic panel
    cPanelLift = RGB(22, 40, 52)    ' #162834  top of the button gradient

    ' Light sources. These are emitters, so unlike the surfaces they are
    ' bright and cold - the icy near-white blue of a lit circuit.
    cTrace = RGB(38, 130, 156)      ' #26829C  unlit circuit
    cCyan = RGB(122, 214, 245)      ' #7AD6F5  lit circuit
    cBloom = RGB(180, 240, 255)     ' #B4F0FF  its halo
    cCore = RGB(240, 253, 255)      ' #F0FDFF  white-hot core
    cDim = RGB(118, 162, 178)       ' #76A2B2  secondary text

    ' Orange is deliberately NOT one hue. The film's own rule is that its
    ' orange light sources "span from saturated yellow to a darker reddish
    ' hue" - the variation is what stops the orange side reading as a
    ' single flat warning colour. So the edge runs hot red-orange and the
    ' halo around it runs yellower, and the pair reads as one light.
    cOrange = RGB(255, 78, 26)      ' #FF4E1A  Rinzler - the reddish end
    cAmber = RGB(255, 160, 51)      ' #FFA033  its halo - the yellow end
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
    PurgeAllSheetBackgroundImages ThisWorkbook   ' the mottled wash behind the data
    UnhideBaselineShapes ws                      ' MUST run before PurgeForeignDecor - see below
    PurgeForeignDecor ws                         ' NavyGold folds / cards / icon shapes
    RemoveAdded ws                               ' our own traces from a previous run
    ApplyCells ws                                ' Grid-black canvas + sidebar slab

    For Each shp In ws.Shapes
        If IsStylable(shp) Then
            StashOriginal shp
            Select Case ShapeRole(shp)
                Case "TITLE":     StyleTitle shp
                Case "PRIMARY":   StyleButton shp, cBloom, cBloom, True    ' Start
                Case "DANGER":    StyleButton shp, cOrange, cAmber, True   ' Reset
                Case "BUTTON":    StyleButton shp, cCyan, cBloom, False
                Case "PANEL_ALT": StylePanel shp, cOrange, cAmber          ' UTILITY badge
                Case Else:        StylePanel shp, cCyan, cBloom
            End Select
            n = n + 1
        End If
    Next shp

    AddCircuitTraces ws
    ApplyTabColour ws

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
    RestoreCells ws
    RestoreTabColour ws

    On Error Resume Next
    ActiveWindow.DisplayGridlines = True
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
    Dim bak As Worksheet: Set bak = GetBak(True)
    If bak Is Nothing Then Exit Sub

    If Len(CStr(bak.Cells(1, 1).Value)) = 0 Then
        Dim c As Range, i As Long: i = 0
        For Each c In ws.Range(CANVAS).Cells
            i = i + 1
            bak.Cells(i, 1).Value = c.Address
            bak.Cells(i, 2).Value = c.Interior.ColorIndex
            bak.Cells(i, 3).Value = c.Interior.Color
            bak.Cells(i, 4).Value = c.Font.Color
            bak.Cells(i, 5).Value = c.Font.Italic
        Next c
    End If
    bak.Visible = xlSheetVeryHidden

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

            With rr.Borders(xlEdgeBottom)
                .LineStyle = xlContinuous: .Weight = xlThin: .Color = cTrace
            End With
            With rr.Borders(xlEdgeLeft)
                .LineStyle = xlContinuous: .Weight = xlThin: .Color = cTrace
            End With
            With rr.Borders(xlEdgeRight)
                .LineStyle = xlContinuous: .Weight = xlThin: .Color = cTrace
            End With
        Next r
    Next addr
End Sub

Private Sub RestoreCells(ByVal ws As Worksheet)
    Dim bak As Worksheet: Set bak = GetBak(False)
    If bak Is Nothing Then Exit Sub
    Dim i As Long, addr As String
    With ws.Range(CANVAS)
        .Borders(xlEdgeBottom).LineStyle = xlNone
        .Borders(xlEdgeLeft).LineStyle = xlNone
        .Borders(xlEdgeRight).LineStyle = xlNone
        .Borders(xlInsideHorizontal).LineStyle = xlNone
        .Borders(xlInsideVertical).LineStyle = xlNone
    End With
    i = 0
    Do
        i = i + 1
        addr = CStr(bak.Cells(i, 1).Value)
        If Len(addr) = 0 Then Exit Do
        If bak.Cells(i, 2).Value = xlNone Then
            ws.Range(addr).Interior.ColorIndex = xlNone
        Else
            ws.Range(addr).Interior.Color = bak.Cells(i, 3).Value
        End If
        ws.Range(addr).Font.Color = bak.Cells(i, 4).Value
        ws.Range(addr).Font.Italic = bak.Cells(i, 5).Value
    Loop
    bak.Cells.Clear
    On Error Resume Next
    bak.Visible = xlSheetVeryHidden
    On Error GoTo 0
End Sub

' Tab colour is stored in the backup sheet's own header cells (row 1 of
' columns G/H), which the per-cell replay above never reads.
Private Sub ApplyTabColour(ByVal ws As Worksheet)
    On Error Resume Next
    Dim bak As Worksheet: Set bak = GetBak(True)
    If bak Is Nothing Then Exit Sub
    If Len(CStr(bak.Range("G1").Value)) = 0 Then
        bak.Range("G1").Value = ws.Tab.ColorIndex
        bak.Range("H1").Value = ws.Tab.Color
    End If
    ws.Tab.Color = cTrace
    On Error GoTo 0
End Sub

Private Sub RestoreTabColour(ByVal ws As Worksheet)
    On Error Resume Next
    Dim bak As Worksheet: Set bak = GetBak(False)
    If bak Is Nothing Then Exit Sub
    If Len(CStr(bak.Range("G1").Value)) = 0 Then Exit Sub
    If bak.Range("G1").Value = xlColorIndexNone Then
        ws.Tab.ColorIndex = xlColorIndexNone
    Else
        ws.Tab.Color = bak.Range("H1").Value
    End If
    bak.Range("G1:H1").ClearContents
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

    ' The sidebar's outer edge is the sheet's main light run, and it is
    ' split at the Action/Utility boundary: cyan down the action half,
    ' orange down the utility half. Two circuits of opposing colour on
    ' one dark slab is the Grid's whole visual signature, and here the
    ' split lands on a division the dashboard already has.
    x = ws.Range("F1").Left
    Set ln = ws.Shapes.AddLine(x, ws.Range("A1").Top, x, ws.Range("A16").Top)
    LightTrace ln, 2#, cCyan, cBloom, 18

    Set ln = ws.Shapes.AddLine(x, ws.Range("A16").Top, x, ws.Range("A30").Top)
    LightTrace ln, 2#, cOrange, cAmber, 18

    ' Horizontal rule capping the sidebar header block.
    Set ln = ws.Shapes.AddLine(ws.Range("A4").Left + 8, ws.Range("A4").Top, _
                               ws.Range("F4").Left - 8, ws.Range("A4").Top)
    LightTrace ln, 1.25, cCyan, cBloom, 12

    ' Baseline closing the slab - orange, since it sits under the
    ' utility half and completes that circuit.
    Set ln = ws.Shapes.AddLine(ws.Range("A30").Left, ws.Range("A30").Top, _
                               ws.Range("F30").Left, ws.Range("A30").Top)
    LightTrace ln, 1.25, cOrange, cAmber, 12
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
        ShapeRole = "PANEL_ALT"        ' the orange half of the Grid
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
    shp.AutoShapeType = msoShapeRoundedRectangle
    shp.Adjustments(1) = CORNER

    ' A flat fill reads as a sticker. A dark vertical gradient makes the
    ' slab look lit from its own edges, which is where the depth comes
    ' from once a strong glow sits around it.
    With shp.Fill
        .Visible = msoTrue
        .TwoColorGradient msoGradientVertical, 1
        .ForeColor.RGB = cPanelLift
        .BackColor.RGB = cVoid
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
            .Color = IIf(emphasise, cCore, cCore)
        End With
        LightText shp, halo, IIf(emphasise, GLOW_TEXT_R + 3, GLOW_TEXT_R)
    End If
    On Error GoTo 0
End Sub

' Font.Glow is a SEPARATE object from Shape.Glow - putting a halo on the
' shape does nothing to its caption. This is the single biggest reason
' the earlier passes read as "dark mode with borders" rather than Tron:
' the lettering was inert. Office 2010+ only, hence the guard.
Private Sub LightText(ByVal shp As Shape, ByVal halo As Long, ByVal radius As Single)
    On Error Resume Next
    With shp.TextFrame2.TextRange.Font.Glow
        .Color.RGB = halo
        .Radius = radius
        .Transparency = GLOW_TEXT_T
    End With
    On Error GoTo 0
End Sub

' Banners / section headers / panel badges. Takes its accent so the two
' faction colours can both appear: ACTION and the section banners run
' cyan, UTILITY runs orange. That is what gives the orange real presence
' on the sheet instead of it existing only on the Reset button - and it
' maps onto a split the dashboard already has, so it carries meaning
' rather than being decoration.
Private Sub StylePanel(ByVal shp As Shape, ByVal edge As Long, ByVal halo As Long)
    On Error Resume Next
    With shp.Fill
        .Visible = msoTrue
        .TwoColorGradient msoGradientVertical, 1
        .ForeColor.RGB = cSlab
        .BackColor.RGB = cVoid
        .Transparency = 0
    End With
    With shp.Line
        .Visible = msoTrue
        .ForeColor.RGB = edge
        .Weight = 1.25
        .Transparency = 0
    End With
    With shp.Glow
        .Color.RGB = halo
        .Radius = GLOW_PANEL_R
        .Transparency = GLOW_PANEL_T
    End With
    shp.Shadow.Visible = msoFalse
    shp.SoftEdge.Type = 0
    If shp.TextFrame.HasText Then
        With shp.TextFrame.Characters.Font
            .Name = "Segoe UI"
            .Color = cCore
            .Bold = True
            .Italic = False
        End With
        LightText shp, halo, GLOW_TEXT_R
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
        ' The title is pure light with no box around it, so its halo is
        ' the widest on the sheet.
        LightText shp, cBloom, 18
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
