Attribute VB_Name = "modThemeDebug"
Option Explicit
'=====================================================================
' modThemeDebug - tells you WHY a theme looks wrong.
'
' Run ThemeDiagnostic and it builds a visible "_ThemeDebug" sheet with
' two parts:
'
'   FINDINGS  - the actual problems, worst first. Each one names the
'               shape and says what to do about it.
'   INVENTORY - every shape on Sheet1: how the theme classified it,
'               whose stash tag it carries, its fill/line colours, and
'               whether those colours belong to the active palette.
'
' It is read-only. It never changes a shape, a cell, or a theme.
'
' What it catches - all four of these are real failures seen in the
' wild, and none of them announce themselves as an error:
'
'   1. Foreign decorations. modNavyGold generates NGADD_* shapes (corner
'      folds, panel cards, icon glyphs). If Tron is active and these are
'      still present, they were restyled instead of removed - gold folds
'      on a black canvas.
'   2. Crossed stashes. A shape tagged TRONORIG| while Navy & Gold is
'      active (or the reverse) means one theme was applied over another
'      and the "original" it saved is really the other theme's colours.
'      Neither theme can be cleanly removed from that point on.
'   3. Off-palette colours. Any fill, line or font colour that is not in
'      the active theme's palette, listed with its hex so you can see at
'      a glance whose colour it is.
'   4. Unclassified shapes. A shape the theme skipped entirely, usually
'      because it is a Picture or a Group rather than an AutoShape - so
'      it keeps its old look while everything around it changes.
'=====================================================================
Private Const SHEET_NAME As String = "Sheet1"
Private Const OUT_SHEET As String = "_ThemeDebug"
Private Const CANVAS As String = "A1:AC30"

Private rowOut As Long
Private nCrit As Long, nWarn As Long

'--------------------------------------------------------------------
Sub ThemeDiagnostic()
    Dim ws As Worksheet, dbg As Worksheet, shp As Shape
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets(SHEET_NAME)
    On Error GoTo 0
    If ws Is Nothing Then MsgBox SHEET_NAME & " not found.", vbCritical: Exit Sub

    Dim active As String
    active = SafeActiveTheme()

    Application.ScreenUpdating = False
    Set dbg = FreshOutputSheet()
    rowOut = 1: nCrit = 0: nWarn = 0

    ' ---------- header ----------
    Title dbg, "THEME DIAGNOSTIC"
    KeyVal dbg, "Run at", Format$(Now, "yyyy-mm-dd hh:nn:ss")
    KeyVal dbg, "Workbook", ThisWorkbook.Name
    KeyVal dbg, "Active theme (recorded)", active
    KeyVal dbg, "Shapes on " & SHEET_NAME, CStr(ws.Shapes.count)
    rowOut = rowOut + 1

    ' ---------- findings ----------
    Dim findRow As Long
    Title dbg, "FINDINGS"
    findRow = rowOut

    CheckForeignDecor dbg, ws, active
    CheckCrossedStashes dbg, ws, active
    CheckHiddenShapes dbg, ws
    CheckUnclassified dbg, ws
    CheckOffPalette dbg, ws, active
    CheckCanvas dbg, ws, active
    CheckBackupSheets dbg, active

    If rowOut = findRow Then
        Line1 dbg, "OK", "No problems found. Every shape matches the active palette.", RGB(0, 140, 90)
    End If
    rowOut = rowOut + 1

    ' ---------- inventory ----------
    Title dbg, "SHAPE INVENTORY"
    dbg.Cells(rowOut, 1).Resize(1, 8).Value = _
        Array("Shape name", "Type", "Role", "Stash", "Caption", "Fill", "Line", "Font")
    With dbg.Cells(rowOut, 1).Resize(1, 8)
        .Font.Bold = True
        .Interior.Color = RGB(238, 238, 242)
        .Borders(xlEdgeBottom).LineStyle = xlContinuous
    End With
    rowOut = rowOut + 1

    For Each shp In ws.Shapes
        dbg.Cells(rowOut, 1).Value = shp.Name
        dbg.Cells(rowOut, 2).Value = ShapeTypeName(shp)
        dbg.Cells(rowOut, 3).Value = RoleOf(shp)
        dbg.Cells(rowOut, 4).Value = StashOf(shp)
        dbg.Cells(rowOut, 5).Value = Left$(CaptionOf(shp), 40)
        dbg.Cells(rowOut, 6).Value = HexOf(FillColour(shp))
        dbg.Cells(rowOut, 7).Value = HexOf(LineColour(shp))
        dbg.Cells(rowOut, 8).Value = HexOf(FontColour(shp))
        Swatch dbg, rowOut, 6, FillColour(shp)
        Swatch dbg, rowOut, 7, LineColour(shp)
        Swatch dbg, rowOut, 8, FontColour(shp)
        rowOut = rowOut + 1
    Next shp

    dbg.Columns("A:H").AutoFit
    If dbg.Columns(5).ColumnWidth > 32 Then dbg.Columns(5).ColumnWidth = 32
    dbg.Rows(1).RowHeight = 22
    dbg.Activate
    dbg.Range("A1").Select
    Application.ScreenUpdating = True

    MsgBox "Diagnostic complete." & vbCrLf & vbCrLf & _
           nCrit & " critical, " & nWarn & " warning(s)." & vbCrLf & _
           "See the " & OUT_SHEET & " sheet.", _
           IIf(nCrit > 0, vbExclamation, vbInformation), "Theme Diagnostic"
End Sub

' Removes the debug sheet when you are done looking at it.
Sub ThemeDiagnosticClear()
    On Error Resume Next
    Application.DisplayAlerts = False
    ThisWorkbook.Sheets(OUT_SHEET).Delete
    Application.DisplayAlerts = True
    On Error GoTo 0
End Sub

'---- individual checks ----------------------------------------------
Private Sub CheckForeignDecor(ByVal dbg As Worksheet, ByVal ws As Worksheet, ByVal active As String)
    Dim shp As Shape, n As Long, names As String
    For Each shp In ws.Shapes
        If Left$(shp.Name, 6) = "NGADD_" And active <> "NAVYGOLD" Then
            n = n + 1
            If n <= 6 Then names = names & IIf(Len(names) > 0, ", ", "") & shp.Name
        End If
    Next shp
    If n > 0 Then
        Line1 dbg, "CRITICAL", n & " Navy & Gold decoration shape(s) still present while '" & active & _
            "' is active (" & names & IIf(n > 6, ", ...", "") & "). These are the corner folds and " & _
            "icon glyphs. Fix: run modThemeManager.ApplyTron, which removes them before styling.", _
            RGB(190, 40, 40)
        nCrit = nCrit + 1
    End If
End Sub

Private Sub CheckCrossedStashes(ByVal dbg As Worksheet, ByVal ws As Worksheet, ByVal active As String)
    Dim shp As Shape, nTron As Long, nNavy As Long
    For Each shp In ws.Shapes
        Select Case StashOf(shp)
            Case "TRONORIG": nTron = nTron + 1
            Case "NGORIG":   nNavy = nNavy + 1
        End Select
    Next shp

    If nTron > 0 And nNavy > 0 Then
        Line1 dbg, "CRITICAL", "Both theme stashes are live at once (" & nTron & " Tron, " & nNavy & _
            " Navy & Gold). One theme was applied on top of the other, so at least one set of " & _
            "'original' colours is really the other theme's. Fix: modThemeManager.RemoveActiveTheme, " & _
            "then apply a single theme.", RGB(190, 40, 40)
        nCrit = nCrit + 1
    ElseIf active = "TRON" And nNavy > 0 Then
        Line1 dbg, "CRITICAL", nNavy & " shape(s) still carry a Navy & Gold stash while Tron is active. " & _
            "Removing Tron will return them to Navy & Gold, not to baseline.", RGB(190, 40, 40)
        nCrit = nCrit + 1
    ElseIf active = "NAVYGOLD" And nTron > 0 Then
        Line1 dbg, "CRITICAL", nTron & " shape(s) still carry a Tron stash while Navy & Gold is active.", _
            RGB(190, 40, 40)
        nCrit = nCrit + 1
    End If
End Sub

' A theme that hides a real shape and draws its own replacement leaves no
' visible trace once the replacement is deleted - the shape is simply
' absent, with nothing to click on and no error. modNavyGold hides the
' Beta title this way (RepositionBeta) and never un-hides it on removal.
Private Sub CheckHiddenShapes(ByVal dbg As Worksheet, ByVal ws As Worksheet)
    Dim shp As Shape, n As Long, detail As String
    For Each shp In ws.Shapes
        If Left$(shp.Name, 8) <> "TRONADD_" And Left$(shp.Name, 6) <> "NGADD_" Then
            If shp.Visible = msoFalse Then
                n = n + 1
                If n <= 6 Then detail = detail & IIf(Len(detail) > 0, ", ", "") & _
                    shp.Name & IIf(Len(CaptionOf(shp)) > 0, " (""" & Left$(CaptionOf(shp), 20) & """)", "")
            End If
        End If
    Next shp
    If n > 0 Then
        Line1 dbg, "CRITICAL", n & " real dashboard shape(s) are hidden: " & detail & _
            IIf(n > 6, ", ...", "") & ". A theme hid them and drew its own replacement, then did " & _
            "not un-hide them on removal - so the shape is gone from the sheet with no error. " & _
            "This is how the nav bar loses its title. Fix: modThemeManager.RemoveActiveTheme, " & _
            "which un-hides them.", RGB(190, 40, 40)
        nCrit = nCrit + 1
    End If
End Sub

Private Sub CheckUnclassified(ByVal dbg As Worksheet, ByVal ws As Worksheet)
    Dim shp As Shape, n As Long, names As String
    For Each shp In ws.Shapes
        If Left$(shp.Name, 8) <> "TRONADD_" And Left$(shp.Name, 6) <> "NGADD_" Then
            If Not (shp.Type = msoAutoShape Or shp.Type = msoFreeform) Then
                n = n + 1
                If n <= 6 Then names = names & IIf(Len(names) > 0, ", ", "") & _
                    shp.Name & " (" & ShapeTypeName(shp) & ")"
            End If
        End If
    Next shp
    If n > 0 Then
        Line1 dbg, "WARNING", n & " shape(s) are not AutoShapes so no theme touches them: " & names & _
            IIf(n > 6, ", ...", "") & ". They keep their old look while everything around them " & _
            "changes. Usually a Picture or a Group - ungroup it, or accept it as-is.", RGB(190, 120, 0)
        nWarn = nWarn + 1
    End If
End Sub

Private Sub CheckOffPalette(ByVal dbg As Worksheet, ByVal ws As Worksheet, ByVal active As String)
    If active <> "TRON" Then Exit Sub
    Dim shp As Shape, n As Long, detail As String, c As Long
    For Each shp In ws.Shapes
        If shp.Type = msoAutoShape Or shp.Type = msoFreeform Then
            c = LineColour(shp)
            If c <> -1 Then
                If Not InTronPalette(c) Then
                    n = n + 1
                    If n <= 6 Then detail = detail & IIf(Len(detail) > 0, "; ", "") & _
                        shp.Name & " line " & HexOf(c) & NameThatColour(c)
                End If
            End If
            c = FillColour(shp)
            If c <> -1 Then
                If Not InTronPalette(c) Then
                    n = n + 1
                    If n <= 6 Then detail = detail & IIf(Len(detail) > 0, "; ", "") & _
                        shp.Name & " fill " & HexOf(c) & NameThatColour(c)
                End If
            End If
        End If
    Next shp
    If n > 0 Then
        Line1 dbg, "WARNING", n & " off-palette colour(s) on themed shapes: " & detail & _
            IIf(n > 6, "; ...", "") & ".", RGB(190, 120, 0)
        nWarn = nWarn + 1
    End If
End Sub

Private Sub CheckCanvas(ByVal dbg As Worksheet, ByVal ws As Worksheet, ByVal active As String)
    If active <> "TRON" Then Exit Sub
    Dim c As Range, n As Long, sample As String
    Dim void As Long, panel As Long, slab As Long
    void = RGB(4, 7, 10): panel = RGB(13, 22, 30): slab = RGB(8, 14, 20)

    For Each c In ws.Range(CANVAS).Cells
        If c.Interior.ColorIndex <> xlNone Then
            If c.Interior.Color <> void And c.Interior.Color <> panel _
               And c.Interior.Color <> slab Then
                n = n + 1
                If n <= 5 Then sample = sample & IIf(Len(sample) > 0, ", ", "") & _
                    c.Address(False, False) & " " & HexOf(c.Interior.Color) & NameThatColour(c.Interior.Color)
            End If
        End If
    Next c

    If n > 0 Then
        Line1 dbg, "WARNING", n & " cell(s) in " & CANVAS & " are filled with a colour outside the " & _
            "Tron palette: " & sample & IIf(n > 5, ", ...", "") & ". If these look like a mottled " & _
            "wash rather than solid blocks, the sheet still has a BACKGROUND IMAGE showing through - " & _
            "cell fills paint over it only where cells are opaque. Fix: re-run ApplyTron, which " & _
            "deletes sheet backgrounds first.", RGB(190, 120, 0)
        nWarn = nWarn + 1
    End If
End Sub

Private Sub CheckBackupSheets(ByVal dbg As Worksheet, ByVal active As String)
    Dim hasTron As Boolean, hasNavy As Boolean
    hasTron = SheetExists("_TronBak")
    hasNavy = SheetExists("_NGBak")

    If hasTron And hasNavy Then
        Line1 dbg, "WARNING", "Both _TronBak and _NGBak backup sheets hold data. Whichever theme is " & _
            "removed second will restore from a baseline the other theme already changed.", RGB(190, 120, 0)
        nWarn = nWarn + 1
    End If
    If active = "NONE" And (hasTron Or hasNavy) Then
        Line1 dbg, "WARNING", "No theme is recorded as active, but a backup sheet still holds cell " & _
            "data - a theme was applied without the manager, or a removal did not finish.", RGB(190, 120, 0)
        nWarn = nWarn + 1
    End If
End Sub

'---- palette knowledge ----------------------------------------------
' Must stay in step with modTronLegacy.InitPalette, or every themed shape
' gets reported as off-palette.
Private Function InTronPalette(ByVal c As Long) As Boolean
    Select Case c
        Case RGB(4, 7, 10), RGB(8, 14, 20), RGB(13, 22, 30), RGB(22, 40, 52), _
             RGB(38, 130, 156), RGB(122, 214, 245), RGB(180, 240, 255), _
             RGB(240, 253, 255), RGB(118, 162, 178), _
             RGB(255, 78, 26), RGB(255, 160, 51)
            InTronPalette = True
    End Select
End Function

' Names the colour when it belongs to a palette we recognise, so a
' finding reads "#D9A441 (Navy & Gold: gold)" instead of a bare hex.
Private Function NameThatColour(ByVal c As Long) As String
    Select Case c
        Case RGB(217, 164, 65):   NameThatColour = " (Navy & Gold: gold)"
        Case RGB(229, 194, 122):  NameThatColour = " (Navy & Gold: soft gold)"
        Case RGB(24, 36, 62):     NameThatColour = " (Navy & Gold: sidebar navy)"
        Case RGB(34, 52, 86):     NameThatColour = " (Navy & Gold: button navy)"
        Case RGB(28, 48, 86):     NameThatColour = " (Navy & Gold: banner navy)"
        Case RGB(158, 69, 60):    NameThatColour = " (Navy & Gold: crimson)"
        Case RGB(243, 241, 236):  NameThatColour = " (Navy & Gold: pearl)"
        Case RGB(255, 255, 255):  NameThatColour = " (white)"
        Case Else:                NameThatColour = ""
    End Select
End Function

'---- shape readers (all failure-tolerant) ---------------------------
Private Function RoleOf(ByVal shp As Shape) As String
    On Error Resume Next
    Dim t As String, hasMac As Boolean
    If Left$(shp.Name, 8) = "TRONADD_" Then RoleOf = "trace": Exit Function
    If Left$(shp.Name, 6) = "NGADD_" Then RoleOf = "NG decor": Exit Function
    If Not (shp.Type = msoAutoShape Or shp.Type = msoFreeform) Then
        RoleOf = "(skipped)": Exit Function
    End If
    hasMac = (Len(shp.OnAction) > 0)
    If shp.TextFrame.HasText Then t = Trim$(shp.TextFrame.Characters.Text)
    If InStr(1, t, "Beta", vbTextCompare) > 0 And Not hasMac Then RoleOf = "TITLE": Exit Function
    If hasMac Then
        If InStr(1, t, "Reset", vbTextCompare) > 0 Then
            RoleOf = "DANGER"
        ElseIf InStr(1, t, "Start", vbTextCompare) > 0 Then
            RoleOf = "PRIMARY"
        Else
            RoleOf = "BUTTON"
        End If
    Else
        RoleOf = "PANEL"
    End If
    On Error GoTo 0
End Function

Private Function StashOf(ByVal shp As Shape) As String
    On Error Resume Next
    Dim a As String
    a = shp.AlternativeText
    If Left$(a, 9) = "TRONORIG|" Then
        StashOf = "TRONORIG"
    ElseIf Left$(a, 7) = "NGORIG|" Then
        StashOf = "NGORIG"
    ElseIf Left$(a, 10) = "GLASSORIG|" Then
        StashOf = "GLASSORIG"
    Else
        StashOf = "-"
    End If
    On Error GoTo 0
End Function

Private Function CaptionOf(ByVal shp As Shape) As String
    On Error Resume Next
    If shp.TextFrame.HasText Then CaptionOf = Trim$(shp.TextFrame.Characters.Text)
    On Error GoTo 0
End Function

Private Function FillColour(ByVal shp As Shape) As Long
    FillColour = -1
    On Error Resume Next
    If shp.Fill.Visible = msoTrue Then FillColour = shp.Fill.ForeColor.RGB
    On Error GoTo 0
End Function

Private Function LineColour(ByVal shp As Shape) As Long
    LineColour = -1
    On Error Resume Next
    If shp.Line.Visible = msoTrue Then LineColour = shp.Line.ForeColor.RGB
    On Error GoTo 0
End Function

Private Function FontColour(ByVal shp As Shape) As Long
    FontColour = -1
    On Error Resume Next
    If shp.TextFrame.HasText Then FontColour = shp.TextFrame.Characters.Font.Color
    On Error GoTo 0
End Function

Private Function ShapeTypeName(ByVal shp As Shape) As String
    On Error Resume Next
    Select Case shp.Type
        Case msoAutoShape:   ShapeTypeName = "AutoShape"
        Case msoFreeform:    ShapeTypeName = "Freeform"
        Case msoPicture:     ShapeTypeName = "Picture"
        Case msoGroup:       ShapeTypeName = "Group"
        Case msoTextBox:     ShapeTypeName = "TextBox"
        Case msoLine:        ShapeTypeName = "Line"
        Case msoFormControl: ShapeTypeName = "FormControl"
        Case Else:           ShapeTypeName = "Type " & shp.Type
    End Select
    On Error GoTo 0
End Function

'---- output helpers -------------------------------------------------
Private Function FreshOutputSheet() As Worksheet
    On Error Resume Next
    ThisWorkbook.Unprotect Password:="p7ss"
    Application.DisplayAlerts = False
    ThisWorkbook.Sheets(OUT_SHEET).Delete
    Application.DisplayAlerts = True
    Set FreshOutputSheet = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.count))
    FreshOutputSheet.Name = OUT_SHEET
    FreshOutputSheet.Cells.Font.Name = "Consolas"
    FreshOutputSheet.Cells.Font.Size = 10
    On Error GoTo 0
End Function

Private Sub Title(ByVal dbg As Worksheet, ByVal t As String)
    dbg.Cells(rowOut, 1).Value = t
    With dbg.Cells(rowOut, 1)
        .Font.Bold = True
        .Font.Size = 12
    End With
    rowOut = rowOut + 1
End Sub

Private Sub KeyVal(ByVal dbg As Worksheet, ByVal k As String, ByVal v As String)
    dbg.Cells(rowOut, 1).Value = k
    dbg.Cells(rowOut, 2).Value = v
    dbg.Cells(rowOut, 2).Font.Bold = True
    rowOut = rowOut + 1
End Sub

Private Sub Line1(ByVal dbg As Worksheet, ByVal sev As String, ByVal msg As String, ByVal clr As Long)
    dbg.Cells(rowOut, 1).Value = sev
    dbg.Cells(rowOut, 1).Font.Bold = True
    dbg.Cells(rowOut, 1).Font.Color = clr
    dbg.Cells(rowOut, 2).Value = msg
    dbg.Cells(rowOut, 2).WrapText = False
    rowOut = rowOut + 1
End Sub

Private Sub Swatch(ByVal dbg As Worksheet, ByVal r As Long, ByVal c As Long, ByVal clr As Long)
    On Error Resume Next
    If clr >= 0 Then dbg.Cells(r, c).Interior.Color = clr
    If clr >= 0 Then dbg.Cells(r, c).Font.Color = IIf(Brightness(clr) > 128, vbBlack, vbWhite)
    On Error GoTo 0
End Sub

Private Function Brightness(ByVal c As Long) As Long
    Brightness = ((c Mod 256) * 30 + ((c \ 256) Mod 256) * 59 + ((c \ 65536) Mod 256) * 11) \ 100
End Function

Private Function HexOf(ByVal c As Long) As String
    If c < 0 Then HexOf = "-": Exit Function
    ' VBA colours are &HBBGGRR; flip to the #RRGGBB people actually read.
    HexOf = "#" & Right$("0" & Hex$(c Mod 256), 2) & _
                  Right$("0" & Hex$((c \ 256) Mod 256), 2) & _
                  Right$("0" & Hex$((c \ 65536) Mod 256), 2)
End Function

Private Function SheetExists(ByVal nm As String) As Boolean
    Dim sh As Object
    On Error Resume Next
    Set sh = ThisWorkbook.Sheets(nm)
    On Error GoTo 0
    SheetExists = Not sh Is Nothing
End Function

' Application.Run, not modThemeManager.ActiveTheme: a qualified call is
' resolved at compile time and needs a module of that exact name, which
' a pasted (rather than imported) .bas will not have. This way the
' diagnostic still runs even when the manager is missing entirely -
' which is precisely when you most want to be able to run it.
Private Function SafeActiveTheme() As String
    SafeActiveTheme = "NONE"
    On Error Resume Next
    SafeActiveTheme = CStr(Application.Run("ActiveTheme"))
    If Len(SafeActiveTheme) = 0 Then SafeActiveTheme = "NONE"
    On Error GoTo 0
End Function
