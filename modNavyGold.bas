Option Explicit
'=====================================================================
' modNavyGold - premium NAVY & GOLD theme for the Sheet1 dashboard.
'
'   ApplyNavyGold   - pearl canvas, navy sidebar/panels, gold-bordered
'                     buttons (Start = gold, Reset = crimson), navy+gold
'                     banners with gold corner folds, warm-white zebra cards
'   RemoveNavyGold  - restores everything (shapes, cells, removes folds)
'
' FULLY REVERSIBLE: shape styles are stashed in each shape's AltText;
' cell fills/fonts are backed up to a very-hidden sheet (_NGBak); the
' added gold folds are named NGADD_* and deleted on remove.
'
' Icons (Segoe MDL2) come in the next pass. RUN ON A COPY first.
'=====================================================================
Private Const SHEET_NAME As String = "Sheet1"
Private Const TAG As String = "NGORIG|"
Private Const ADD_PFX As String = "NGADD_"
Private Const BAK As String = "_NGBak"
Private Const CANVAS As String = "A1:AC30"       ' pearl canvas
Private Const CONTENT_FIRST As Long = 4
Private Const CONTENT_LAST As Long = 27
Private Const CONTENT_COLS As String = "G:T"

' palette (set by InitPalette; RGB() is used so no hand-computed Longs)
Private cNavy As Long, cSoftNavy As Long, cPearl As Long, cWarmWhite As Long
Private cGold As Long, cSoftGold As Long, cWarmBorder As Long
Private cText As Long, cText2 As Long, cMutedNavy As Long
Private cCrimson As Long, cCream As Long, cWhite As Long

Private Sub InitPalette()
    cNavy = RGB(17, 25, 54)          ' #111936 Deep Navy
    cSoftNavy = RGB(24, 32, 63)      ' #18203F Soft Navy
    cPearl = RGB(243, 241, 236)      ' #F3F1EC Pearl
    cWarmWhite = RGB(255, 253, 248)  ' #FFFDF8 Warm White
    cGold = RGB(217, 164, 65)        ' #D9A441 Champagne Gold
    cSoftGold = RGB(229, 194, 122)   ' #E5C27A Soft Gold
    cWarmBorder = RGB(217, 212, 200) ' #D9D4C8 Warm Border
    cText = RGB(37, 37, 37)          ' #252525 Primary Text
    cText2 = RGB(107, 107, 107)      ' #6B6B6B Secondary Text
    cMutedNavy = RGB(58, 66, 97)     ' #3A4261 Muted Navy
    cCrimson = RGB(158, 69, 60)      ' Reset danger
    cCream = RGB(251, 237, 234)      ' light text on dark buttons
    cWhite = RGB(244, 241, 234)      ' banner/title text
End Sub

'--------------------------------------------------------------------
Sub ApplyNavyGold()
    Dim ws As Worksheet
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
    On Error Resume Next
    Application.CommandBars.ExecuteMso "SheetBackgroundDelete"   ' clear any leftover aurora/pearl bg image
    On Error GoTo 0
    ApplyCells ws
    StyleShapes ws
    RepositionBeta ws
    AddFolds ws
    AddPanelDecor ws
    AddContentFrames ws
    ' AddIcons ws   ' re-enabled once IconDiagnostic confirms the glyph codes
    On Error Resume Next
    ActiveWindow.DisplayGridlines = False
    On Error GoTo 0

    If wasProt Then ws.Protect Password:="p7ss"
    If wbProt Then ThisWorkbook.Protect Password:="p7ss", Structure:=True
    Application.ScreenUpdating = True
    MsgBox "Navy & Gold theme applied. Run RemoveNavyGold to undo." & vbCrLf & _
           "(Icons come in the next pass.)", vbInformation, "Navy & Gold"
End Sub

Sub RemoveNavyGold()
    Dim ws As Worksheet
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
    RestorePositions ws
    RestoreShapes ws
    RestoreCells ws
    On Error Resume Next
    ActiveWindow.DisplayGridlines = True
    On Error GoTo 0

    If wasProt Then ws.Protect Password:="p7ss"
    If wbProt Then ThisWorkbook.Protect Password:="p7ss", Structure:=True
    Application.ScreenUpdating = True
    MsgBox "Navy & Gold theme removed; everything restored.", vbInformation, "Navy & Gold"
End Sub

' One-shot: strip the leftover aurora/pearl BACKGROUND PICTURE off every
' sheet (Sheet1 and ConsolidatedData still had one, which the Export Trx
' File macro copies into the exported workbook). Cosmetic only - no data
' is touched. Run this once.
Sub ClearAllSheetBackgrounds()
    Dim ws As Worksheet, prev As Object
    Set prev = ActiveSheet
    Application.ScreenUpdating = False
    For Each ws In ThisWorkbook.Worksheets
        On Error Resume Next
        ws.Activate
        Application.CommandBars.ExecuteMso "SheetBackgroundDelete"
        On Error GoTo 0
    Next ws
    On Error Resume Next
    prev.Activate
    On Error GoTo 0
    Application.ScreenUpdating = True
    MsgBox "Removed leftover background pictures from all sheets." & vbCrLf & _
           "Re-run Export Trx File - it'll come out clean now.", vbInformation, "Backgrounds cleared"
End Sub

'---- CELLS: pearl canvas + warm-white zebra content ----------------
Private Sub ApplyCells(ByVal ws As Worksheet)
    Dim bak As Worksheet: Set bak = GetBak(True)
    ' back up canvas cells (skip if already backed up)
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

    ws.Range(CANVAS).Interior.Color = cPearl                 ' pearl base
    ws.Range("A1:E29").Interior.Color = cNavy                ' navy sidebar strip (cols A-E)

    Dim r As Long
    For r = CONTENT_FIRST To CONTENT_LAST
        Dim rr As Range
        Set rr = ws.Range(Split(CONTENT_COLS, ":")(0) & r & ":" & Split(CONTENT_COLS, ":")(1) & r)
        If r Mod 2 = 1 Then rr.Interior.Color = cWarmWhite Else rr.Interior.Color = cPearl  ' zebra
        rr.Font.Color = cText
        rr.Borders(xlEdgeBottom).Color = cWarmBorder
        rr.Borders(xlEdgeBottom).Weight = xlThin
    Next r
End Sub

Private Sub RestoreCells(ByVal ws As Worksheet)
    Dim bak As Worksheet: Set bak = GetBak(False)
    If bak Is Nothing Then Exit Sub
    Dim i As Long, addr As String
    ws.Range(CANVAS).Borders(xlEdgeBottom).LineStyle = xlNone
    ws.Range(CANVAS).Borders(xlInsideHorizontal).LineStyle = xlNone
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

'---- SHAPES: restyle existing --------------------------------------
Private Sub StyleShapes(ByVal ws As Worksheet)
    Dim shp As Shape, k As String
    For Each shp In ws.Shapes
        k = ShapeKind(shp)
        If k <> "skip" Then
            StashOriginal shp
            Select Case k
                Case "button":     StyleButton shp
                Case "title":      StyleTitle shp
                Case "panelcard":  StylePanelCard shp
                Case "panellabel": StylePanelLabel shp
                Case "banner":     StyleBanner shp
            End Select
        End If
    Next shp
End Sub

Private Sub RestoreShapes(ByVal ws As Worksheet)
    Dim shp As Shape
    For Each shp In ws.Shapes
        If Left$(GetAlt(shp), Len(TAG)) = TAG Then RestoreOriginal shp
    Next shp
End Sub

Private Function ShapeKind(ByVal shp As Shape) As String
    On Error Resume Next
    ShapeKind = "skip"
    If shp.Type <> msoAutoShape And shp.Type <> msoFreeform Then Exit Function
    If Left$(shp.Name, Len(ADD_PFX)) = ADD_PFX Then Exit Function
    Dim hasMac As Boolean, txt As String
    hasMac = (Len(shp.OnAction) > 0)
    If shp.TextFrame.HasText Then txt = Trim$(shp.TextFrame.Characters.Text)
    If hasMac Then
        ShapeKind = "button"
    ElseIf InStr(1, txt, "Beta", vbTextCompare) = 1 Then
        ShapeKind = "title"
    ElseIf Len(txt) = 0 Then
        ShapeKind = "panelcard"
    ElseIf InStr(1, txt, "Panel", vbTextCompare) > 0 Then
        ShapeKind = "panellabel"
    Else
        ShapeKind = "banner"
    End If
    On Error GoTo 0
End Function

Private Sub StyleButton(ByVal shp As Shape)
    On Error Resume Next
    shp.AutoShapeType = msoShapeRoundedRectangle
    shp.Adjustments(1) = 0.28
    Dim txt As String: txt = ""
    If shp.TextFrame.HasText Then txt = shp.TextFrame.Characters.Text
    With shp.Fill: .Visible = msoTrue: .Solid: End With
    If InStr(1, txt, "Start", vbTextCompare) > 0 Then
        shp.Fill.ForeColor.RGB = cGold
        shp.Line.ForeColor.RGB = cSoftGold
        SetText shp, cNavy, True
    ElseIf InStr(1, txt, "Reset", vbTextCompare) > 0 Then
        shp.Fill.ForeColor.RGB = cCrimson
        shp.Line.ForeColor.RGB = RGB(196, 106, 96)
        SetText shp, cCream, False
    Else
        shp.Fill.ForeColor.RGB = cSoftNavy
        shp.Line.ForeColor.RGB = cGold
        SetText shp, cCream, False
    End If
    shp.Line.Weight = 1
    shp.Line.Transparency = 0.3
    On Error GoTo 0
End Sub

Private Sub StyleTitle(ByVal shp As Shape)
    On Error Resume Next
    shp.AutoShapeType = msoShapeRoundedRectangle
    With shp.Fill: .Visible = msoTrue: .Solid: .ForeColor.RGB = cNavy: End With
    With shp.Line: .Visible = msoTrue: .ForeColor.RGB = cGold: .Weight = 1.25: End With
    SetText shp, cWhite, True
    On Error GoTo 0
End Sub

Private Sub StylePanelCard(ByVal shp As Shape)
    On Error Resume Next
    ' transparent so the navy sidebar CELLS show through; just a gold outline
    ' (filling a custom-geometry freeform directly is unreliable)
    shp.Fill.Visible = msoFalse
    With shp.Line: .Visible = msoTrue: .ForeColor.RGB = cGold: .Weight = 1.25: End With
    On Error GoTo 0
End Sub

Private Sub StylePanelLabel(ByVal shp As Shape)
    On Error Resume Next
    With shp.Fill: .Visible = msoFalse: End With
    With shp.Line: .Visible = msoFalse: End With
    SetText shp, cGold, True
    On Error GoTo 0
End Sub

Private Sub StyleBanner(ByVal shp As Shape)
    On Error Resume Next
    With shp.Fill
        .Visible = msoTrue
        .TwoColorGradient msoGradientHorizontal, 1
        .ForeColor.RGB = cNavy
        .BackColor.RGB = cSoftNavy
    End With
    With shp.Line: .Visible = msoTrue: .ForeColor.RGB = cGold: .Weight = 1.25: End With
    SetText shp, cWhite, True
    On Error GoTo 0
End Sub

Private Sub SetText(ByVal shp As Shape, ByVal clr As Long, ByVal center As Boolean)
    On Error Resume Next
    If shp.TextFrame.HasText Then
        shp.TextFrame.Characters.Font.Color = clr
        If center Then shp.TextFrame2.TextRange.ParagraphFormat.Alignment = msoAlignCenter
    End If
    On Error GoTo 0
End Sub

'---- FOLDS: gold corner triangle on each banner --------------------
Private Sub AddFolds(ByVal ws As Worksheet)
    Dim shp As Shape, n As Long
    For Each shp In ws.Shapes
        If ShapeKind(shp) = "banner" Then
            n = n + 1
            Dim sz As Single: sz = 26
            Dim t As Shape
            Set t = ws.Shapes.AddShape(msoShapeRightTriangle, shp.Left + shp.Width - sz, shp.Top, sz, sz)
            t.Name = ADD_PFX & "FOLD_" & n
            t.Rotation = 90
            With t.Fill: .Visible = msoTrue: .Solid: .ForeColor.RGB = cGold: End With
            t.Line.Visible = msoFalse
        End If
    Next shp
End Sub

'---- ICONS: Segoe MDL2 glyphs on banners + buttons -----------------
' Each icon is a separate borderless textbox (tagged NGADD_) so it's
' removed cleanly and never alters the banner/button text.
' Glyph codes are Segoe MDL2 Assets; change any hex here if one renders
' as the wrong picture on your build.
Private Sub AddIcons(ByVal ws As Worksheet)
    AddIconAt ws, FindByText(ws, "Alert Related"), ChrW(&HE7BA), cGold, True    ' warning
    AddIconAt ws, FindByText(ws, "Customer Inf"), ChrW(&HE77B), cGold, True     ' contact
    AddIconAt ws, FindByText(ws, "Counterparty Inf"), ChrW(&HE716), cGold, True ' people
    AddIconAt ws, FindByText(ws, "Country Risk"), ChrW(&HE774), cGold, True     ' globe
    AddIconAt ws, FindButton(ws, "Start"), ChrW(&HE768), cNavy, False           ' play
    AddIconAt ws, FindButton(ws, "Export Trx"), ChrW(&HE898), cCream, False     ' upload
    AddIconAt ws, FindButton(ws, "OSDD"), ChrW(&HE721), cCream, False           ' search
    AddIconAt ws, FindButton(ws, "Generate Narr"), ChrW(&HE70F), cCream, False  ' edit
    AddIconAt ws, FindButton(ws, "Rename"), ChrW(&HE8AC), cCream, False         ' rename
    AddIconAt ws, FindButton(ws, "PDF Merge"), ChrW(&HE8A5), cCream, False      ' document
    AddIconAt ws, FindButton(ws, "Profile Warm"), ChrW(&HE945), cCream, False   ' lightbulb
    AddIconAt ws, FindButton(ws, "Reset"), ChrW(&HE72C), cCream, False          ' refresh
End Sub

Private Sub AddIconAt(ByVal ws As Worksheet, ByVal host As Shape, ByVal glyph As String, _
                      ByVal clr As Long, ByVal isBanner As Boolean)
    On Error Resume Next
    If host Is Nothing Then Exit Sub
    Static cnt As Long: cnt = cnt + 1
    Dim w As Single, h As Single, leftPad As Single, fsz As Single
    If isBanner Then
        w = 26: h = 26: leftPad = 12: fsz = 15
    Else
        w = 22: h = 18: leftPad = 12: fsz = 11
    End If
    Dim ic As Shape
    Set ic = ws.Shapes.AddTextbox(msoTextOrientationHorizontal, _
             host.Left + leftPad, host.Top + (host.Height - h) / 2, w, h)
    ic.Name = ADD_PFX & "ICON_" & cnt
    ic.Fill.Visible = msoFalse
    ic.Line.Visible = msoFalse
    With ic.TextFrame
        .Characters.Text = glyph
        .Characters.Font.Name = "Segoe MDL2 Assets"
        .Characters.Font.Size = fsz
        .Characters.Font.Color = clr
        .HorizontalAlignment = xlHAlignCenter
        .VerticalAlignment = xlVAlignCenter
        .MarginLeft = 0: .MarginRight = 0: .MarginTop = 0: .MarginBottom = 0
    End With
    On Error GoTo 0
End Sub

'---- CONTENT FRAMES: rounded gold outline around each section ------
Private Sub AddContentFrames(ByVal ws As Worksheet)
    ' ONE rounded gold outline around the whole content area (left edge at
    ' column F) - no messy per-section internal lines.
    AddFrame2 ws, "F3:T28"
End Sub

Private Sub AddFrame2(ByVal ws As Worksheet, ByVal addr As String)
    On Error Resume Next
    Static c As Long: c = c + 1
    Dim r As Range: Set r = ws.Range(addr)
    Dim shp As Shape
    Set shp = ws.Shapes.AddShape(msoShapeRoundedRectangle, r.Left - 2, r.Top - 2, r.Width + 4, r.Height + 4)
    shp.Name = ADD_PFX & "CFRAME_" & c
    shp.Fill.Visible = msoFalse                 ' no fill -> click-through, cells show
    With shp.Line: .Visible = msoTrue: .ForeColor.RGB = cGold: .Weight = 1.25: End With
    shp.Adjustments(1) = 0.05                    ' rounded corners
    shp.ZOrder msoSendToBack                     ' behind banners; wraps the section
    On Error GoTo 0
End Sub

'---- PANEL DECOR: gold underline + flanking lines (from the demo) ---
Private Sub AddPanelDecor(ByVal ws As Worksheet)
    On Error Resume Next
    Dim beta As Shape, actLbl As Shape, utlLbl As Shape
    Set beta = FindByText(ws, "Beta")
    Set actLbl = FindByText(ws, "Action Panel")
    Set utlLbl = FindByText(ws, "Utl")           ' "Utlity Panel"
    ' gold underline just under Beta 3.5
    If Not beta Is Nothing Then
        Dim by As Single: by = beta.Top + beta.Height - 2
        AddGoldLine ws, beta.Left + beta.Width * 0.2, by, beta.Left + beta.Width * 0.8, by, 1.25
    End If
    AddFlank ws, actLbl
    AddFlank ws, utlLbl
    On Error GoTo 0
End Sub

Private Sub AddFlank(ByVal ws As Worksheet, ByVal lbl As Shape)
    On Error Resume Next
    If lbl Is Nothing Then Exit Sub
    Dim y As Single, sbL As Single, sbR As Single
    y = lbl.Top + lbl.Height / 2
    sbL = ws.Range("A1").Left + 14
    sbR = ws.Range("E1").Left + ws.Range("E1").Width - 14
    If lbl.Left - 8 > sbL Then AddGoldLine ws, sbL, y, lbl.Left - 8, y, 0.75
    If sbR > lbl.Left + lbl.Width + 8 Then AddGoldLine ws, lbl.Left + lbl.Width + 8, y, sbR, y, 0.75
    On Error GoTo 0
End Sub

Private Sub AddGoldLine(ByVal ws As Worksheet, ByVal x1 As Single, ByVal y1 As Single, _
                        ByVal x2 As Single, ByVal y2 As Single, ByVal wt As Single)
    On Error Resume Next
    Static c As Long: c = c + 1
    Dim ln As Shape
    Set ln = ws.Shapes.AddLine(x1, y1, x2, y2)
    ln.Name = ADD_PFX & "LINE_" & c
    ln.Line.ForeColor.RGB = cGold
    ln.Line.Weight = wt
    On Error GoTo 0
End Sub

Private Sub RemoveAdded(ByVal ws As Worksheet)
    Dim i As Long
    For i = ws.Shapes.count To 1 Step -1
        If Left$(ws.Shapes(i).Name, Len(ADD_PFX)) = ADD_PFX Then ws.Shapes(i).Delete
    Next i
End Sub

'---- REPOSITION: Beta 3.5 parallel with Alert ----------------------
Private Sub RepositionBeta(ByVal ws As Worksheet)
    On Error Resume Next
    Dim bak As Worksheet: Set bak = GetBak(False)
    If Not bak Is Nothing Then
        If Len(CStr(bak.Cells(1, "G").Value)) > 0 Then Exit Sub    ' already repositioned
    End If

    Dim beta As Shape, actLbl As Shape, shield As Shape, actCard As Shape
    Set beta = FindByText(ws, "Beta")
    Set actLbl = FindByText(ws, "Action Panel")
    Set shield = FindPicture(ws)
    Set actCard = FindUpperBlank(ws)
    If beta Is Nothing Or actLbl Is Nothing Then Exit Sub

    Dim sbLeft As Single, sbWidth As Single, newTop As Single, newH As Single, shift As Single
    sbLeft = ws.Range("A1").Left
    sbWidth = ws.Range("A1:E1").Width
    newTop = ws.Range("A3").Top          ' deterministic: row 3 (level with the Alert banner)
    newH = ws.Range("A3:A4").Height       ' ~2 rows tall
    shift = newH + 16                     ' extra gap so Action Panel isn't cramped under Beta

    ' shift the Action group DOWN to make room for Beta on top
    PosBackup ws, actLbl
    actLbl.Top = actLbl.Top + shift

    Dim nm As Variant, s As Shape
    For Each nm In Array("Start", "Export Trx File", "OSDD Search", "Generate Narrative")
        Set s = FindButton(ws, CStr(nm))
        If Not s Is Nothing Then
            PosBackup ws, s
            s.Top = s.Top + shift
        End If
    Next nm

    If Not actCard Is Nothing Then
        PosBackup ws, actCard
        actCard.Top = actCard.Top + shift
    End If

    ' move Beta into the sidebar top, level with the Alert banner
    PosBackup ws, beta
    beta.Left = sbLeft: beta.Top = newTop: beta.Width = sbWidth: beta.Height = newH

    ' bring the shield with it
    If Not shield Is Nothing Then
        PosBackup ws, shield
        shield.Top = newTop + (newH - shield.Height) / 2
        shield.Left = sbLeft + 8
    End If
    On Error GoTo 0
End Sub

Private Sub PosBackup(ByVal ws As Worksheet, ByVal shp As Shape)
    On Error Resume Next
    Dim bak As Worksheet: Set bak = GetBak(True)
    Dim r As Long
    r = bak.Cells(bak.Rows.count, "G").End(xlUp).Row
    If Len(CStr(bak.Cells(1, "G").Value)) = 0 Then r = 0
    r = r + 1
    bak.Cells(r, "G").Value = shp.Name
    bak.Cells(r, "H").Value = shp.Left
    bak.Cells(r, "I").Value = shp.Top
    bak.Cells(r, "J").Value = shp.Width
    bak.Cells(r, "K").Value = shp.Height
    On Error GoTo 0
End Sub

Private Sub RestorePositions(ByVal ws As Worksheet)
    On Error Resume Next
    Dim bak As Worksheet: Set bak = GetBak(False)
    If bak Is Nothing Then Exit Sub
    Dim r As Long, nm As String, shp As Shape
    r = 0
    Do
        r = r + 1
        nm = CStr(bak.Cells(r, "G").Value)
        If Len(nm) = 0 Then Exit Do
        Set shp = Nothing
        Set shp = ws.Shapes(nm)
        If Not shp Is Nothing Then
            shp.Left = bak.Cells(r, "H").Value
            shp.Top = bak.Cells(r, "I").Value
            shp.Width = bak.Cells(r, "J").Value
            shp.Height = bak.Cells(r, "K").Value
        End If
    Loop
    bak.Range("G:K").Clear
    On Error GoTo 0
End Sub

Private Function FindByText(ByVal ws As Worksheet, ByVal t As String) As Shape
    Dim shp As Shape
    For Each shp In ws.Shapes
        On Error Resume Next
        If shp.TextFrame.HasText Then
            If InStr(1, shp.TextFrame.Characters.Text, t, vbTextCompare) > 0 Then
                Set FindByText = shp: Exit Function
            End If
        End If
        On Error GoTo 0
    Next shp
End Function

Private Function FindButton(ByVal ws As Worksheet, ByVal t As String) As Shape
    Dim shp As Shape
    For Each shp In ws.Shapes
        On Error Resume Next
        If Len(shp.OnAction) > 0 And shp.TextFrame.HasText Then
            If InStr(1, shp.TextFrame.Characters.Text, t, vbTextCompare) > 0 Then
                Set FindButton = shp: Exit Function
            End If
        End If
        On Error GoTo 0
    Next shp
End Function

Private Function FindPicture(ByVal ws As Worksheet) As Shape
    Dim shp As Shape
    For Each shp In ws.Shapes
        If shp.Type = msoPicture Then Set FindPicture = shp: Exit Function
    Next shp
End Function

Private Function FindUpperBlank(ByVal ws As Worksheet) As Shape
    Dim shp As Shape, best As Shape, blank As Boolean
    For Each shp In ws.Shapes
        On Error Resume Next
        blank = True
        If shp.TextFrame.HasText Then
            If Len(Trim$(shp.TextFrame.Characters.Text)) > 0 Then blank = False
        End If
        If Len(shp.OnAction) > 0 Then blank = False
        If (shp.Type = msoAutoShape Or shp.Type = msoFreeform) And blank _
           And Left$(shp.Name, Len(ADD_PFX)) <> ADD_PFX Then
            If best Is Nothing Then
                Set best = shp
            ElseIf shp.Top < best.Top Then
                Set best = shp
            End If
        End If
        On Error GoTo 0
    Next shp
    Set FindUpperBlank = best
End Function

'---- stash / restore (shape AltText) -------------------------------
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
    s = TAG & shp.AutoShapeType & "|" & shp.Fill.Visible & "|" & shp.Fill.ForeColor.RGB & "|" & _
        CLng(shp.Fill.Transparency * 1000) & "|" & shp.Line.Visible & "|" & shp.Line.ForeColor.RGB & "|" & _
        CLng(shp.Line.Weight * 100) & "|" & ht & "|" & fn & "|" & CLng(fs * 10) & "|" & fb & "|" & fi & "|" & fc
    SetAlt shp, s
    On Error GoTo 0
End Sub

Private Sub RestoreOriginal(ByVal shp As Shape)
    On Error Resume Next
    Dim p() As String: p = Split(GetAlt(shp), "|")
    If UBound(p) < 7 Then Exit Sub
    If CLng(p(1)) > 0 Then shp.AutoShapeType = CLng(p(1))
    If CLng(p(2)) = msoTrue Then
        shp.Fill.Visible = msoTrue: shp.Fill.Solid
        shp.Fill.ForeColor.RGB = CLng(p(3))
        shp.Fill.Transparency = CLng(p(4)) / 1000
    Else
        shp.Fill.Visible = msoFalse
    End If
    If CLng(p(5)) = msoTrue Then
        shp.Line.Visible = msoTrue: shp.Line.ForeColor.RGB = CLng(p(6)): shp.Line.Weight = CLng(p(7)) / 100
    Else
        shp.Line.Visible = msoFalse
    End If
    If UBound(p) >= 13 Then
        If CLng(p(8)) <> 0 And Len(p(9)) > 0 Then
            With shp.TextFrame.Characters.Font
                .Name = p(9): .Size = CLng(p(10)) / 10
                .Bold = (CLng(p(11)) <> 0): .Italic = (CLng(p(12)) <> 0): .Color = CLng(p(13))
            End With
        End If
    End If
    SetAlt shp, ""
    On Error GoTo 0
End Sub

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
        Set GetBak = ThisWorkbook.Sheets.Add
        GetBak.Name = BAK
    End If
End Function
