Attribute VB_Name = "modNavyGold"
Option Explicit
'=====================================================================
' modNavyGold - premium NAVY & GOLD theme for the Sheet1 dashboard.
'
'   ApplyNavyGold   - pearl canvas, full-height slate navy sidebar (#18243E),
'                     dual rounded gold-bordered panel cards, gold-bordered
'                     buttons (Start = gold, Reset = crimson), navy+gold
'                     banners with gold corner folds, solid white data cards
'   RemoveNavyGold  - restores everything (shapes, cells, removes folds/frames)
'=====================================================================
Private Const SHEET_NAME As String = "Sheet1"
Private Const TAG As String = "NGORIG|"
Private Const ADD_PFX As String = "NGADD_"
Private Const BAK As String = "_NGBak"
Private Const CANVAS As String = "A1:AC30"       ' pearl canvas

' palette (set by InitPalette; RGB() is used so no hand-computed Longs)
Private cNavy As Long, cSoftNavy As Long, cPearl As Long, cWarmWhite As Long
Private cGold As Long, cSoftGold As Long, cWarmBorder As Long
Private cText As Long, cText2 As Long, cMutedNavy As Long
Private cCrimson As Long, cCream As Long, cWhite As Long
Private cBannerNavy As Long, cBannerSoftNavy As Long
Private cSidebarBg As Long, cBtnNavy As Long

Private Sub InitPalette()
    cNavy = RGB(24, 36, 62)              ' #18243E Rich Slate Navy (sidebar background)
    cSoftNavy = RGB(34, 52, 86)          ' #223456 Button & Badge Navy
    cSidebarBg = cNavy                   ' #18243E Full-Height Sidebar Background
    cBtnNavy = cSoftNavy                 ' #223456 Elevated Button Tone

    cBannerNavy = RGB(28, 48, 86)        ' #1C3056 Rich Executive Slate Navy (headers)
    cBannerSoftNavy = RGB(42, 68, 115)   ' #2A4473 Soft Cobalt-Slate gradient highlight
    cPearl = RGB(243, 241, 236)          ' #F3F1EC Pearl Base Canvas
    cWarmWhite = RGB(255, 255, 255)      ' #FFFFFF Solid Pure White (consistent data cells)
    cGold = RGB(217, 164, 65)            ' #D9A441 Champagne Gold
    cSoftGold = RGB(229, 194, 122)       ' #E5C27A Soft Gold
    cWarmBorder = RGB(225, 222, 214)     ' #E1DED6 Subtle Warm Border
    cText = RGB(37, 37, 37)              ' #252525 Primary Text
    cText2 = RGB(107, 107, 107)          ' #6B6B6B Secondary Text
    cMutedNavy = RGB(58, 66, 97)         ' #3A4261 Muted Navy
    cCrimson = RGB(158, 69, 60)          ' #9E453C Reset Danger
    cCream = RGB(251, 237, 234)          ' #FBEDEA Light text on dark buttons
    cWhite = RGB(255, 255, 255)          ' #FFFFFF Pure White
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
    Application.CommandBars.ExecuteMso "SheetBackgroundDelete"   ' clear leftover background image
    On Error GoTo 0

    RemoveAdded ws                                             ' clear existing decorators
    ApplyCells ws                                              ' solid white cards + full-height slate sidebar
    StyleShapes ws
    RepositionBeta ws                                          ' position buttons and panel cards with precision
    AddPanelCards ws                                           ' perfectly fitted upper & lower rounded gold boxes
    RemoveStraySidebarIcons ws                                 ' permanently delete stray icons in sidebar
    AddFolds ws
    AddIconsInline ws                                          ' glyphs prepended into shape text

    On Error Resume Next
    ActiveWindow.DisplayGridlines = False
    On Error GoTo 0

    If wasProt Then ws.Protect Password:="p7ss"
    If wbProt Then ThisWorkbook.Protect Password:="p7ss", Structure:=True
    Application.ScreenUpdating = True
    MsgBox "Navy & Gold theme applied.", vbInformation, "Navy & Gold"
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
    RestoreText ws
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
    MsgBox "Removed leftover background pictures from all sheets.", vbInformation, "Backgrounds cleared"
End Sub

'---- CELLS: solid white cards (NO zebra) + full-height sidebar ----
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

    ' 1. Paint entire canvas in Pearl base
    ws.Range(CANVAS).Interior.Color = cPearl
    ws.Range(CANVAS).Borders(xlEdgeBottom).LineStyle = xlNone
    ws.Range(CANVAS).Borders(xlInsideHorizontal).LineStyle = xlNone

    ' 2. Full-height sidebar strip (cols A-E, Rows 1 to 29) - aligns flush with Country Risk
    ws.Range("A1:E29").Interior.Color = cSidebarBg

    ' 3. Paint data rows in ONE single consistent pure white color (NO zebra striping):
    Dim dataRowRanges As Variant, rngAddr As Variant
    dataRowRanges = Array("G5:T10", "G13:T14", "G17:T23", "G26:T27")

    Dim secRng As Range, r As Long, rr As Range
    For Each rngAddr In dataRowRanges
        Set secRng = ws.Range(CStr(rngAddr))
        For r = secRng.Row To secRng.Row + secRng.Rows.count - 1
            Set rr = ws.Range("G" & r & ":T" & r)
            rr.Interior.Color = cWarmWhite              ' Solid pure white
            rr.Font.Color = cText
            rr.Font.Name = "Segoe UI"
            rr.Borders(xlEdgeBottom).Color = cWarmBorder
            rr.Borders(xlEdgeBottom).Weight = xlThin
            rr.Borders(xlEdgeLeft).Color = cWarmBorder
            rr.Borders(xlEdgeLeft).Weight = xlThin
            rr.Borders(xlEdgeRight).Color = cWarmBorder
            rr.Borders(xlEdgeRight).Weight = xlThin
        Next r
    Next rngAddr
End Sub

Private Sub RestoreCells(ByVal ws As Worksheet)
    Dim bak As Worksheet: Set bak = GetBak(False)
    If bak Is Nothing Then Exit Sub
    Dim i As Long, addr As String
    ws.Range(CANVAS).Borders(xlEdgeBottom).LineStyle = xlNone
    ws.Range(CANVAS).Borders(xlInsideHorizontal).LineStyle = xlNone
    ws.Range(CANVAS).Borders(xlEdgeLeft).LineStyle = xlNone
    ws.Range(CANVAS).Borders(xlEdgeRight).LineStyle = xlNone
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

'---- PANEL CARDS: perfectly fitted rounded gold boxes enclosing button groups
Private Sub AddPanelCards(ByVal ws As Worksheet)
    On Error Resume Next
    Dim sbLeft As Single, sbWidth As Single
    sbLeft = ws.Range("A1").Left + 10
    sbWidth = ws.Range("A1:E1").Width - 20

    ' 1. Upper Rounded Box enclosing Action Group (Rows 3 to 12)
    Dim y1 As Single, h1 As Single, c1 As Shape
    y1 = ws.Range("A3").Top
    h1 = (ws.Range("A11").Top + ws.Range("A11").Height) - y1 + 8
    Set c1 = ws.Shapes.AddShape(msoShapeRoundedRectangle, sbLeft, y1, sbWidth, h1)
    c1.Name = ADD_PFX & "PANELCARD_ACT"
    c1.Adjustments(1) = 0.08
    c1.Fill.Visible = msoFalse                 ' transparent -> deep navy sidebar shows through
    With c1.Line
        .Visible = msoTrue
        .ForeColor.RGB = cGold                  ' Champagne Gold border
        .Weight = 1.25                          ' 1.25pt clean border
    End With
    c1.ZOrder msoSendToBack

    ' 2. Lower Rounded Box enclosing Utility Group (Rows 14 to 23)
    Dim y2 As Single, h2 As Single, c2 As Shape
    y2 = ws.Range("A14").Top
    h2 = (ws.Range("A22").Top + ws.Range("A22").Height) - y2 + 8
    Set c2 = ws.Shapes.AddShape(msoShapeRoundedRectangle, sbLeft, y2, sbWidth, h2)
    c2.Name = ADD_PFX & "PANELCARD_UTL"
    c2.Adjustments(1) = 0.08
    c2.Fill.Visible = msoFalse                 ' transparent -> deep navy sidebar shows through
    With c2.Line
        .Visible = msoTrue
        .ForeColor.RGB = cGold                  ' Champagne Gold border
        .Weight = 1.25                          ' 1.25pt clean border
    End With
    c2.ZOrder msoSendToBack
    On Error GoTo 0
End Sub

'---- CONTENT FRAMES: removed (no golden outer wireframe) -----------
Private Sub AddContentFrames(ByVal ws As Worksheet)
    ' Intentionally empty: removed outer golden ring
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
        shp.Visible = msoTrue
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
    ElseIf InStr(1, txt, "Beta", vbTextCompare) > 0 Then
        ShapeKind = "title"
    ElseIf Len(txt) = 0 Then
        ShapeKind = "panelcard"
    ElseIf InStr(1, txt, "Panel", vbTextCompare) > 0 Or InStr(1, txt, "ACTION", vbTextCompare) > 0 Or InStr(1, txt, "UTILITY", vbTextCompare) > 0 Then
        ShapeKind = "panellabel"
    Else
        ShapeKind = "banner"
    End If
    On Error GoTo 0
End Function

Private Sub StyleButton(ByVal shp As Shape)
    On Error Resume Next
    shp.AutoShapeType = msoShapeRoundedRectangle
    shp.Adjustments(1) = 0.28                   ' soft pill-shaped modern rounded rectangle
    Dim txt As String: txt = ""
    If shp.TextFrame.HasText Then txt = shp.TextFrame.Characters.Text
    With shp.Fill: .Visible = msoTrue: .Solid: End With
    If InStr(1, txt, "Start", vbTextCompare) > 0 Then
        shp.Fill.ForeColor.RGB = cGold
        shp.Line.ForeColor.RGB = cSoftGold
        shp.Line.Weight = 1.25
        shp.Line.Transparency = 0
        SetText shp, cNavy, True
    ElseIf InStr(1, txt, "Reset", vbTextCompare) > 0 Then
        shp.Fill.ForeColor.RGB = cCrimson
        shp.Line.ForeColor.RGB = RGB(196, 106, 96)
        shp.Line.Weight = 1
        shp.Line.Transparency = 0.2
        SetText shp, cCream, True
    Else
        shp.Fill.ForeColor.RGB = cBtnNavy       ' #223456 Button Navy
        shp.Line.ForeColor.RGB = cGold
        shp.Line.Weight = 1
        shp.Line.Transparency = 0.3
        SetText shp, cCream, True
    End If
    If shp.TextFrame.HasText Then
        With shp.TextFrame.Characters.Font
            .Bold = True
            .Size = 9.5
            .Name = "Segoe UI"
        End With
    End If
    On Error GoTo 0
End Sub

Private Sub StyleTitle(ByVal shp As Shape)
    On Error Resume Next
    shp.Visible = msoFalse                      ' completely hide Beta 3.5 shape
    On Error GoTo 0
End Sub

Private Sub StylePanelCard(ByVal shp As Shape)
    On Error Resume Next
    shp.Visible = msoFalse                      ' superseded by AddPanelCards
    On Error GoTo 0
End Sub

Private Sub StylePanelLabel(ByVal shp As Shape)
    On Error Resume Next
    shp.Visible = msoTrue
    shp.AutoShapeType = msoShapeRoundedRectangle
    shp.Adjustments(1) = 0.3
    With shp.Fill
        .Visible = msoTrue
        .Solid
        .ForeColor.RGB = cBtnNavy               ' #223456 Button Navy fill
    End With
    With shp.Line
        .Visible = msoTrue
        .ForeColor.RGB = cGold                  ' Champagne Gold border
        .Weight = 1
        .Transparency = 0.15
    End With
    SetText shp, cGold, True
    If shp.TextFrame.HasText Then
        With shp.TextFrame.Characters.Font
            .Bold = True
            .Size = 9.5
            .Name = "Segoe UI"
        End With
    End If
    On Error GoTo 0
End Sub

Private Sub StyleBanner(ByVal shp As Shape)
    On Error Resume Next
    shp.AutoShapeType = msoShapeRoundedRectangle
    shp.Adjustments(1) = 0.15                   ' soft curved banner corners
    With shp.Fill
        .Visible = msoTrue
        .TwoColorGradient msoGradientHorizontal, 1
        .ForeColor.RGB = cBannerNavy            ' #1C3056 Rich Slate Navy (softer & lighter)
        .BackColor.RGB = cBannerSoftNavy        ' #2A4473 Soft Cobalt-Slate
    End With
    With shp.Line
        .Visible = msoTrue
        .ForeColor.RGB = cGold                  ' Champagne Gold border
        .Weight = 1.25
    End With
    SetText shp, cWhite, True
    If shp.TextFrame.HasText Then
        With shp.TextFrame.Characters.Font
            .Bold = True
            .Size = 10.5
            .Name = "Segoe UI"
        End With
    End If
    On Error GoTo 0
End Sub

Private Sub SetText(ByVal shp As Shape, ByVal clr As Long, ByVal center As Boolean)
    On Error Resume Next
    If shp.TextFrame.HasText Then
        shp.TextFrame.Characters.Font.Color = clr
        shp.TextFrame.Characters.Font.Name = "Segoe UI"
        If center Then shp.TextFrame2.TextRange.ParagraphFormat.Alignment = msoAlignCenter
        shp.TextFrame2.VerticalAnchor = msoAnchorMiddle
    End If
    On Error GoTo 0
End Sub

'---- FOLDS: gold corner triangle on each banner --------------------
Private Sub AddFolds(ByVal ws As Worksheet)
    Dim shp As Shape, n As Long
    For Each shp In ws.Shapes
        If ShapeKind(shp) = "banner" Then
            n = n + 1
            Dim sz As Single: sz = 24
            Dim t As Shape
            Set t = ws.Shapes.AddShape(msoShapeRightTriangle, shp.Left + shp.Width - sz, shp.Top, sz, sz)
            t.Name = ADD_PFX & "FOLD_" & n
            t.Rotation = 90
            With t.Fill: .Visible = msoTrue: .Solid: .ForeColor.RGB = cGold: End With
            t.Line.Visible = msoFalse
        End If
    Next shp
End Sub

'---- ICONS: Segoe MDL2 glyphs on buttons ---------------------------
Private Sub AddIconsInline(ByVal ws As Worksheet)
    PrependIcon ws, FindByText(ws, "Alert Related"), ChrW(&HE7BA&), cGold     ' warning
    PrependIcon ws, FindByText(ws, "Customer Inf"), ChrW(&HE77B&), cGold      ' contact
    PrependIcon ws, FindByText(ws, "Counterparty Inf"), ChrW(&HE716&), cGold  ' people
    PrependIcon ws, FindByText(ws, "Country Risk"), ChrW(&HE774&), cGold      ' globe
    PrependIcon ws, FindButton(ws, "Start"), ChrW(&HE768&), cNavy            ' play
    PrependIcon ws, FindButton(ws, "Export Trx"), ChrW(&HE898&), cGold       ' upload
    PrependIcon ws, FindButton(ws, "OSDD"), ChrW(&HE721&), cGold             ' search
    PrependIcon ws, FindButton(ws, "Generate Narr"), ChrW(&HE70F&), cGold    ' edit
    PrependIcon ws, FindButton(ws, "Rename"), ChrW(&HE8AC&), cGold           ' rename
    PrependIcon ws, FindButton(ws, "PDF Merge"), ChrW(&HEA90&), cGold        ' PDF
    PrependIcon ws, FindButton(ws, "Profile Warm"), ChrW(&HE945&), cGold     ' lightning
    PrependIcon ws, FindButton(ws, "Reset"), ChrW(&HE72C&), cCream           ' refresh
End Sub

Private Sub PrependIcon(ByVal ws As Worksheet, ByVal shp As Shape, ByVal glyph As String, ByVal clr As Long)
    On Error Resume Next
    If shp Is Nothing Then Exit Sub
    If Not shp.TextFrame.HasText Then Exit Sub
    Dim orig As String: orig = Trim$(shp.TextFrame.Characters.Text)
    If AscW(Left$(orig, 1)) < 0 Then Exit Sub          ' already carries a glyph
    TextBackup ws, shp.Name, orig
    shp.TextFrame.Characters.Text = glyph & "  " & orig
    With shp.TextFrame.Characters(1, 1).Font
        .Name = "Segoe MDL2 Assets"
        .Color = clr
    End With
    On Error GoTo 0
End Sub

Private Sub TextBackup(ByVal ws As Worksheet, ByVal nm As String, ByVal txt As String)
    On Error Resume Next
    Dim bak As Worksheet: Set bak = GetBak(True)
    Dim r As Long
    r = bak.Cells(bak.Rows.count, "M").End(xlUp).Row
    If Len(CStr(bak.Cells(1, "M").Value)) = 0 And r = 1 Then r = 0
    bak.Cells(r + 1, "M").Value = nm
    bak.Cells(r + 1, "N").Value = txt
    On Error GoTo 0
End Sub

Private Sub RestoreText(ByVal ws As Worksheet)
    On Error Resume Next
    Dim bak As Worksheet: Set bak = GetBak(False)
    If bak Is Nothing Then Exit Sub
    Dim r As Long, nm As String, shp As Shape
    r = 0
    Do
        r = r + 1
        nm = CStr(bak.Cells(r, "M").Value)
        If Len(nm) = 0 Then Exit Do
        Set shp = Nothing
        Set shp = ws.Shapes(nm)
        If Not shp Is Nothing Then
            If shp.TextFrame.HasText Then shp.TextFrame.Characters.Text = CStr(bak.Cells(r, "N").Value)
        End If
    Loop
    bak.Range("M:N").Clear
    On Error GoTo 0
End Sub

Private Sub RemoveAdded(ByVal ws As Worksheet)
    Dim i As Long
    For i = ws.Shapes.count To 1 Step -1
        If Left$(ws.Shapes(i).Name, Len(ADD_PFX)) = ADD_PFX Then ws.Shapes(i).Delete
    Next i
End Sub

'---- REPOSITION: Align Header & Buttons inside Sidebar -------------
Private Sub RepositionBeta(ByVal ws As Worksheet)
    On Error Resume Next
    Dim bak As Worksheet: Set bak = GetBak(False)

    Dim beta As Shape, actLbl As Shape, utlLbl As Shape
    Set beta = FindByText(ws, "Beta")
    Set actLbl = FindByText(ws, "Action Panel")
    If actLbl Is Nothing Then Set actLbl = FindByText(ws, "ACTION")
    Set utlLbl = FindByText(ws, "Utility Panel")
    If utlLbl Is Nothing Then Set utlLbl = FindByText(ws, "UTILITY")

    ' Hide Beta 3.5 completely
    If Not beta Is Nothing Then
        PosBackup ws, beta
        beta.Visible = msoFalse
    End If

    Dim sbLeft As Single, sbWidth As Single
    sbLeft = ws.Range("A1").Left + 14
    sbWidth = ws.Range("A1:E1").Width - 28

    ' 1. ACTION PANEL badge at Row 3 (inside the upper box)
    If Not actLbl Is Nothing Then
        PosBackup ws, actLbl
        actLbl.Visible = msoTrue
        actLbl.Left = sbLeft + 8
        actLbl.Top = ws.Range("A3").Top + 2
        actLbl.Width = sbWidth - 16
        actLbl.Height = 22
    End If

    ' 2. Action Buttons (Rows 5, 7, 9, 11)
    Dim actBtns As Variant, nm As Variant, s As Shape, rBtn As Long
    actBtns = Array("Start", "Export Trx File", "OSDD Search", "Generate Narrative")
    rBtn = 5
    For Each nm In actBtns
        Set s = FindButton(ws, CStr(nm))
        If Not s Is Nothing Then
            PosBackup ws, s
            s.Left = sbLeft
            s.Top = ws.Range("A" & rBtn).Top + 2
            s.Width = sbWidth
            s.Height = 24
        End If
        rBtn = rBtn + 2
    Next nm

    ' 3. UTILITY PANEL badge at Row 14 (inside the lower box)
    If Not utlLbl Is Nothing Then
        PosBackup ws, utlLbl
        utlLbl.Visible = msoTrue
        utlLbl.Left = sbLeft + 8
        utlLbl.Top = ws.Range("A14").Top + 2
        utlLbl.Width = sbWidth - 16
        utlLbl.Height = 22
    End If

    ' 4. Utility Buttons (Rows 16, 18, 20, 22)
    Dim utlBtns As Variant
    utlBtns = Array("Rename", "PDF Merge", "Profile Warm-Up", "Reset")
    rBtn = 16
    For Each nm In utlBtns
        Set s = FindButton(ws, CStr(nm))
        If Not s Is Nothing Then
            PosBackup ws, s
            s.Left = sbLeft
            s.Top = ws.Range("A" & rBtn).Top + 2
            s.Width = sbWidth
            s.Height = 24
        End If
        rBtn = rBtn + 2
    Next nm
    On Error GoTo 0
End Sub

'---- REMOVE STRAY ICONS / PICTURES IN SIDEBAR ---------------------
Private Sub RemoveStraySidebarIcons(ByVal ws As Worksheet)
    On Error Resume Next
    Dim i As Long, shp As Shape, sbRight As Single
    sbRight = ws.Range("F1").Left
    For i = ws.Shapes.count To 1 Step -1
        Set shp = ws.Shapes(i)
        ' Check if shape is inside the sidebar area (Cols A-E) and not dynamically generated
        If shp.Left < sbRight And Left$(shp.Name, Len(ADD_PFX)) <> ADD_PFX Then
            ' Delete pictures, vector graphics (msoGraphic), OLE objects, or orphan shapes without macros
            If shp.Type = msoPicture Or shp.Type = 28 Or shp.Type = msoLinkedPicture _
               Or shp.Type = msoOLEControlObject Then
                shp.Delete
            ElseIf Len(shp.OnAction) = 0 Then
                Dim txt As String: txt = ""
                If shp.TextFrame.HasText Then txt = Trim$(shp.TextFrame.Characters.Text)
                ' If it's not a panel label, delete it
                If InStr(1, txt, "ACTION", vbTextCompare) = 0 And _
                   InStr(1, txt, "UTILITY", vbTextCompare) = 0 And _
                   InStr(1, txt, "Beta", vbTextCompare) = 0 And _
                   Len(txt) > 0 Then
                    shp.Delete
                End If
            End If
        End If
    Next i
    On Error GoTo 0
End Sub

Private Sub PosBackup(ByVal ws As Worksheet, ByVal shp As Shape)
    On Error Resume Next
    Dim bak As Worksheet: Set bak = GetBak(True)
    Dim r As Long
    For r = 1 To bak.Cells(bak.Rows.count, "G").End(xlUp).Row
        If CStr(bak.Cells(r, "G").Value) = shp.Name Then Exit Sub
    Next r
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
