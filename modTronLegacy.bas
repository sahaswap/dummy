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
Private Const CANVAS As String = "A1:AC30"
Private Const CORNER As Single = 0.12               ' hard-edged Grid panel, not a soft card

' --- palette (set by InitPalette) ---
Private cVoid As Long          ' Grid floor - near-black with a blue bias
Private cPanel As Long         ' raised panel / data cell
Private cSlab As Long          ' sidebar slab - the unlit block light sits on
Private cTrace As Long         ' unlit circuit trace
Private cCyan As Long          ' lit circuit - THE Tron Legacy blue
Private cBloom As Long         ' glow bloom / primary action edge
Private cCore As Long          ' white-hot core (text)
Private cDim As Long           ' secondary / de-emphasised text
Private cOrange As Long        ' Rinzler - Reset only

Private Sub InitPalette()
    cVoid = RGB(4, 7, 10)          ' #04070A
    cPanel = RGB(9, 15, 21)        ' #090F15
    cSlab = RGB(6, 11, 16)         ' #060B10
    cTrace = RGB(27, 108, 127)     ' #1B6C7F
    cCyan = RGB(111, 195, 223)     ' #6FC3DF  <- the canonical one
    cBloom = RGB(168, 232, 249)    ' #A8E8F9
    cCore = RGB(242, 254, 255)     ' #F2FEFF
    cDim = RGB(106, 145, 158)      ' #6A919E
    cOrange = RGB(242, 111, 33)    ' #F26F21
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
    PurgeForeignDecor ws                         ' NavyGold folds / cards / icon shapes
    RemoveAdded ws                               ' our own traces from a previous run
    ApplyCells ws                                ' Grid-black canvas + sidebar slab

    For Each shp In ws.Shapes
        If IsStylable(shp) Then
            StashOriginal shp
            Select Case ShapeRole(shp)
                Case "TITLE":   StyleTitle shp
                Case "PRIMARY": StyleButton shp, cBloom, True     ' Start
                Case "DANGER":  StyleButton shp, cOrange, True    ' Reset
                Case "BUTTON":  StyleButton shp, cCyan, False
                Case Else:      StylePanel shp
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

' modNavyGold generates its own decorations - corner folds, panel cards,
' inline icon glyphs - all named NGADD_*. They are Navy & Gold objects,
' not dashboard content, so Tron deletes them rather than restyling them.
' ApplyNavyGold regenerates every one of them, so this is not destructive.
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

    ' 3. Data bands - raised panel, white-hot text, unlit hairline rules.
    '    Italic is cleared: Grid typography is upright, and the inherited
    '    italic labels were a Navy & Gold mannerism that read as a glitch
    '    against this palette. Bold is left alone - it carries the real
    '    label/value distinction in the sheet.
    Dim bands As Variant, addr As Variant, sec As Range, r As Long, rr As Range
    bands = Array("G5:T10", "G13:T14", "G17:T23", "G26:T27")
    For Each addr In bands
        Set sec = ws.Range(CStr(addr))
        For r = sec.Row To sec.Row + sec.Rows.count - 1
            Set rr = ws.Range("G" & r & ":T" & r)
            rr.Interior.Color = cPanel
            rr.Font.Color = cCore
            rr.Font.Name = "Segoe UI"
            rr.Font.Italic = False
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

    ' Vertical ribbon down the sidebar's outer edge, full slab height.
    x = ws.Range("F1").Left
    Set ln = ws.Shapes.AddLine(x, ws.Range("A1").Top, x, ws.Range("A30").Top)
    LightTrace ln, 1.5, cCyan, 10

    ' Horizontal rule capping the sidebar header block.
    Set ln = ws.Shapes.AddLine(ws.Range("A4").Left + 8, ws.Range("A4").Top, _
                               ws.Range("F4").Left - 8, ws.Range("A4").Top)
    LightTrace ln, 1#, cTrace, 6

    ' Baseline closing the slab.
    Set ln = ws.Shapes.AddLine(ws.Range("A30").Left, ws.Range("A30").Top, _
                               ws.Range("F30").Left, ws.Range("A30").Top)
    LightTrace ln, 1#, cTrace, 6
    On Error GoTo 0
End Sub

Private Sub LightTrace(ByVal shp As Shape, ByVal w As Single, _
                       ByVal clr As Long, ByVal glowR As Single)
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
        .Color.RGB = cBloom
        .Radius = glowR
        .Transparency = 0.55
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
    End If
    On Error GoTo 0
End Function

'---- styling ---------------------------------------------------------
' One button routine, three intensities. Start and Reset are the two
' consequential controls on the sheet, so both are lit harder than the
' rest: Start in the bright bloom cyan, Reset in Rinzler orange. The
' neutral buttons stay at standard cyan and recede.
Private Sub StyleButton(ByVal shp As Shape, ByVal edge As Long, ByVal emphasise As Boolean)
    On Error Resume Next
    shp.AutoShapeType = msoShapeRoundedRectangle
    shp.Adjustments(1) = CORNER
    With shp.Fill
        .Visible = msoTrue: .Solid
        .ForeColor.RGB = cPanel
        .Transparency = 0.05           ' near-solid: a slab, not frosted glass
    End With
    With shp.Line
        .Visible = msoTrue
        .ForeColor.RGB = edge
        .Weight = IIf(emphasise, 2#, 1.25)
        .Transparency = 0
    End With
    ' Native Office glow - this is the bloom. Shadow cannot do it.
    With shp.Glow
        .Color.RGB = edge
        .Radius = IIf(emphasise, 16, 8)
        .Transparency = IIf(emphasise, 0.2, 0.45)
    End With
    shp.Shadow.Visible = msoFalse
    shp.SoftEdge.Type = 0              ' Grid edges are hard
    If shp.TextFrame.HasText Then
        With shp.TextFrame.Characters.Font
            .Name = "Segoe UI"
            .Bold = emphasise
            .Size = 9
            .Color = IIf(emphasise, edge, cCore)
        End With
    End If
    On Error GoTo 0
End Sub

' Banners / section headers - dark slab, unlit edge, cyan label.
Private Sub StylePanel(ByVal shp As Shape)
    On Error Resume Next
    With shp.Fill
        .Visible = msoTrue: .Solid
        .ForeColor.RGB = cSlab
        .Transparency = 0.05
    End With
    With shp.Line
        .Visible = msoTrue
        .ForeColor.RGB = cTrace
        .Weight = 0.75
        .Transparency = 0.2
    End With
    With shp.Glow
        .Color.RGB = cBloom
        .Radius = 4
        .Transparency = 0.75           ' barely there - panels sit back
    End With
    shp.Shadow.Visible = msoFalse
    shp.SoftEdge.Type = 0
    If shp.TextFrame.HasText Then
        With shp.TextFrame.Characters.Font
            .Name = "Segoe UI"
            .Color = cCyan
            .Bold = True
            .Italic = False
        End With
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
        .Color.RGB = cCyan
        .Radius = 20
        .Transparency = 0.25
    End With
    If shp.TextFrame.HasText Then
        With shp.TextFrame.Characters.Font
            .Name = "Segoe UI"
            .Color = cCore
            .Bold = True
            .Italic = False
        End With
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

    shp.Glow.Radius = 0                ' kill the bloom first
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
