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
    AddFolds ws
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
    With shp.Fill: .Visible = msoTrue: .Solid: .ForeColor.RGB = cNavy: End With
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

Private Sub RemoveAdded(ByVal ws As Worksheet)
    Dim i As Long
    For i = ws.Shapes.count To 1 Step -1
        If Left$(ws.Shapes(i).Name, Len(ADD_PFX)) = ADD_PFX Then ws.Shapes(i).Delete
    Next i
End Sub

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
