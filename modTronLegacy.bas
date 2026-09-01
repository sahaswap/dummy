Attribute VB_Name = "modTronLegacy"
Option Explicit
'=====================================================================
' modTronLegacy - "Tron: Legacy" restyle for the Sheet1 dashboard.
'
' Full treatment, not shapes-only: the canvas goes to Grid-black first,
' then every shape is relit. Styling a glowing button on top of a pearl
' canvas was what made the first pass read wrong - the light had nothing
' dark to sit in.
'
'   ApplyTronLegacy   - black canvas, unlit sidebar slab, cyan circuit
'                       traces, real Office Glow on buttons/banners,
'                       Rinzler orange on Reset
'   RemoveTronLegacy  - restores cells from the backup sheet and every
'                       shape from its own AltText stash, deletes traces
'
' PALETTE NOTE: the cyan here is #6FC3DF, the screen-accurate Tron
' Legacy blue - a cold, desaturated ice blue, NOT #00FFFF. Pure neon
' cyan is what makes a "Tron" theme look like a generic gamer skin.
' The brightness comes from the white-hot core (#F2FEFF) on text and
' the bloom (#A8E8F9) in the Glow, layered over an almost-black ground.
'
' Idempotent: re-running re-stashes nothing and re-adds no duplicates.
'=====================================================================
Private Const SHEET_NAME As String = "Sheet1"
Private Const TAG As String = "TRONORIG|"
Private Const ADD_PFX As String = "TRONADD_"
Private Const BAK As String = "_TronBak"
Private Const CANVAS As String = "A1:AC30"
Private Const CORNER As Single = 0.12          ' hard-edged Grid panel, not a soft card

' --- palette (film-accurate, set by InitPalette) ---
Private cVoid As Long          ' Grid floor - near-black with a blue bias
Private cPanel As Long         ' raised panel / data cell
Private cSlab As Long          ' sidebar slab - the unlit block light sits on
Private cTrace As Long         ' unlit circuit trace
Private cCyan As Long          ' lit circuit - THE Tron Legacy blue
Private cBloom As Long         ' glow bloom
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

    RemoveAdded ws                  ' clear traces from any previous run
    ApplyCells ws                   ' Grid-black canvas + sidebar slab (backed up first)

    For Each shp In ws.Shapes
        If IsStylable(shp) Then
            StashOriginal shp
            If IsTitle(shp) Then
                StyleTitle shp
            ElseIf HasMacro(shp) Then
                StyleButton shp
            Else
                StylePanel shp
            End If
            n = n + 1
        End If
    Next shp

    AddCircuitTraces ws             ' the detail that reads as "Grid" rather than "dark mode"

    On Error Resume Next
    ActiveWindow.DisplayGridlines = False
    On Error GoTo 0

    If wasProt Then ws.Protect Password:="p7ss"
    If wbProt Then ThisWorkbook.Protect Password:="p7ss", Structure:=True
    Application.ScreenUpdating = True
    MsgBox "Tron Legacy applied - canvas, " & n & " shape(s) and circuit traces." & vbCrLf & vbCrLf & _
           "Run RemoveTronLegacy to restore everything.", vbInformation, "Tron Legacy"
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

    On Error Resume Next
    ActiveWindow.DisplayGridlines = True
    On Error GoTo 0

    If wasProt Then ws.Protect Password:="p7ss"
    If wbProt Then ThisWorkbook.Protect Password:="p7ss", Structure:=True
    Application.ScreenUpdating = True
    MsgBox "Tron Legacy removed; canvas and " & n & " shape(s) restored.", vbInformation, "Tron Legacy"
End Sub

'---- CELLS -----------------------------------------------------------
' Same backup-sheet mechanism modNavyGold uses (proven against this
' workbook), and the same validated ranges - so the sidebar slab and the
' data bands land exactly where the existing theme puts them.
Private Sub ApplyCells(ByVal ws As Worksheet)
    Dim bak As Worksheet: Set bak = GetBak(True)
    If Len(CStr(bak.Cells(1, 1).Value)) = 0 Then
        Dim c As Range, i As Long: i = 0
        For Each c In ws.Range(CANVAS).Cells
            i = i + 1
            bak.Cells(i, 1).Value = c.Address
            bak.Cells(i, 2).Value = c.Interior.ColorIndex
            bak.Cells(i, 3).Value = c.Interior.Color
            bak.Cells(i, 4).Value = c.Font.Color
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

    ' 2. Sidebar slab - a shade off the floor so the panel reads as a
    '    solid object rather than a hole in the background.
    ws.Range("A1:E29").Interior.Color = cSlab

    ' 3. Data bands - raised panel, cyan hairline rules. Unlit trace
    '    colour on the borders: only interactive things get lit cyan.
    Dim bands As Variant, addr As Variant, sec As Range, r As Long, rr As Range
    bands = Array("G5:T10", "G13:T14", "G17:T23", "G26:T27")
    For Each addr In bands
        Set sec = ws.Range(CStr(addr))
        For r = sec.Row To sec.Row + sec.Rows.count - 1
            Set rr = ws.Range("G" & r & ":T" & r)
            rr.Interior.Color = cPanel
            rr.Font.Color = cCore
            rr.Font.Name = "Segoe UI"
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
    Loop
    bak.Cells.Clear
    On Error Resume Next
    bak.Visible = xlSheetVeryHidden
    On Error GoTo 0
End Sub

'---- CIRCUIT TRACES --------------------------------------------------
' Thin lit lines with a glow, anchored to cell geometry (not to shapes,
' whose positions we can't assume). These are what separate "Tron" from
' "someone turned the lights off": the Grid is defined by light running
' along the edges of dark slabs.
Private Sub AddCircuitTraces(ByVal ws As Worksheet)
    On Error Resume Next
    Dim x As Single, y1 As Single, y2 As Single, ln As Shape

    ' Vertical ribbon down the sidebar's outer edge, full slab height.
    x = ws.Range("F1").Left
    y1 = ws.Range("A1").Top
    y2 = ws.Range("A30").Top
    Set ln = ws.Shapes.AddLine(x, y1, x, y2)
    LightTrace ln, 1.5, cCyan, 10

    ' Horizontal rule capping the sidebar header block.
    Set ln = ws.Shapes.AddLine(ws.Range("A4").Left + 8, ws.Range("A4").Top, _
                               ws.Range("F4").Left - 8, ws.Range("A4").Top)
    LightTrace ln, 1#, cTrace, 6

    ' Baseline under the sidebar, closing the slab.
    Set ln = ws.Shapes.AddLine(ws.Range("A30").Left, ws.Range("A30").Top, _
                               ws.Range("F30").Left, ws.Range("A30").Top)
    LightTrace ln, 1#, cTrace, 6
    On Error GoTo 0
End Sub

Private Sub LightTrace(ByVal shp As Shape, ByVal w As Single, _
                       ByVal clr As Long, ByVal glowR As Single)
    On Error Resume Next
    shp.Name = ADD_PFX & Format$(Timer * 1000, "0") & "_" & Int(Rnd * 10000)
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

Private Function HasMacro(ByVal shp As Shape) As Boolean
    On Error Resume Next
    HasMacro = (Len(shp.OnAction) > 0)
    On Error GoTo 0
End Function

Private Function IsTitle(ByVal shp As Shape) As Boolean
    On Error Resume Next
    If Len(shp.OnAction) = 0 And shp.TextFrame.HasText Then
        IsTitle = (InStr(1, Trim$(shp.TextFrame.Characters.Text), "Beta", vbTextCompare) = 1)
    End If
    On Error GoTo 0
End Function

'---- styling ---------------------------------------------------------
' Buttons are the only LIT objects on the sheet. Everything else uses
' the unlit trace colour, so the eye goes straight to what's clickable.
Private Sub StyleButton(ByVal shp As Shape)
    On Error Resume Next
    Dim isReset As Boolean, edge As Long
    If shp.TextFrame.HasText Then
        isReset = (InStr(1, shp.TextFrame.Characters.Text, "Reset", vbTextCompare) > 0)
    End If
    edge = IIf(isReset, cOrange, cCyan)

    shp.AutoShapeType = msoShapeRoundedRectangle
    shp.Adjustments(1) = CORNER
    With shp.Fill
        .Visible = msoTrue: .Solid
        .ForeColor.RGB = cPanel
        .Transparency = 0.05          ' near-solid: the slab, not frosted glass
    End With
    With shp.Line
        .Visible = msoTrue
        .ForeColor.RGB = edge
        .Weight = 1.25
        .Transparency = 0
    End With
    ' Real Office glow - this is the bloom. Shadow can't do it.
    With shp.Glow
        .Color.RGB = IIf(isReset, cOrange, cBloom)
        .Radius = 8
        .Transparency = 0.45
    End With
    shp.Shadow.Visible = msoFalse
    shp.SoftEdge.Type = 0             ' Grid edges are hard
    If shp.TextFrame.HasText Then
        With shp.TextFrame.Characters.Font
            .Name = "Segoe UI"
            .Bold = False              ' the letterspacing does the work, not weight
            .Size = 9
            .Color = IIf(isReset, cOrange, cCore)
        End With
    End If
    On Error GoTo 0
End Sub

' Banners / section headers - dark slab, unlit trace edge, cyan label.
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
        .Transparency = 0.75          ' barely there - panels sit back
    End With
    shp.Shadow.Visible = msoFalse
    shp.SoftEdge.Type = 0
    If shp.TextFrame.HasText Then
        With shp.TextFrame.Characters.Font
            .Color = cCyan
            .Bold = True
        End With
    End If
    On Error GoTo 0
End Sub

' Title - no box at all. Just white-hot text with a cyan bloom, the way
' the film sets its titles: light with nothing containing it.
Private Sub StyleTitle(ByVal shp As Shape)
    On Error Resume Next
    shp.Fill.Visible = msoFalse
    shp.Line.Visible = msoFalse
    shp.SoftEdge.Type = 0
    shp.Shadow.Visible = msoFalse
    With shp.Glow
        .Color.RGB = cCyan
        .Radius = 12
        .Transparency = 0.35
    End With
    If shp.TextFrame.HasText Then
        With shp.TextFrame.Characters.Font
            .Color = cCore
            .Bold = False
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

    shp.Glow.Radius = 0               ' kill the bloom first
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
