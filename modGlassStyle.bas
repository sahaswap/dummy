Option Explicit
'=====================================================================
' modGlassStyle - rounded "glass" restyle for the Sheet1 dashboard.
'
' Works on the dashboard SHAPES (your buttons and section banners are
' real shapes, not Form/ActiveX controls), so it can restyle them in
' place - no converting, no re-wiring macros.
'
'   ApplyGlassStyle  - buttons -> rounded, frosted-glass, soft-edged;
'                      section panels/banners -> lightly frosted.
'   RemoveGlassStyle - restores every shape to exactly how it was.
'
' HOW REVERSAL WORKS: before changing a shape, its original style
' (auto-shape type, fill, line) is stashed in the shape's AlternativeText
' as "GLASSORIG|...". RemoveGlassStyle reads that back and restores it,
' then clears the tag. Re-running ApplyGlassStyle is safe (idempotent) -
' it never overwrites an existing stash.
'
' Touches ONLY shape formatting - no cells, values, text, or macros.
' Pairs with the frosted background (ApplyPearlBackground).
'=====================================================================
Private Const SHEET_NAME As String = "Sheet1"
Private Const TAG As String = "GLASSORIG|"
Private Const CORNER As Single = 0.35          ' button corner roundness (0..0.5)
Private Const BTN_TRANS As Single = 0.18       ' button fill transparency
Private Const PANEL_TRANS As Single = 0.5      ' banner/panel fill transparency
Private Const GLASS_WHITE As Long = 16777215   ' RGB(255,255,255)

Sub ApplyGlassStyle()
    Dim ws As Worksheet, shp As Shape, n As Long
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets(SHEET_NAME)
    On Error GoTo 0
    If ws Is Nothing Then MsgBox SHEET_NAME & " not found.", vbCritical: Exit Sub

    Dim wasProt As Boolean
    On Error Resume Next
    wasProt = ws.ProtectContents
    ws.Unprotect Password:="p7ss"
    On Error GoTo 0

    For Each shp In ws.Shapes
        If IsStylableShape(shp) Then
            StashOriginal shp
            If HasMacro(shp) Then
                StyleButton shp
            Else
                StylePanel shp
            End If
            n = n + 1
        End If
    Next shp

    On Error Resume Next
    ActiveWindow.DisplayGridlines = False
    On Error GoTo 0
    If wasProt Then ws.Protect Password:="p7ss"

    MsgBox "Glass style applied to " & n & " dashboard shape(s)." & vbCrLf & vbCrLf & _
           "Buttons are now rounded frosted glass; panels are lightly frosted." & vbCrLf & _
           "Run RemoveGlassStyle to undo (restores every shape exactly).", _
           vbInformation, "Glass Style"
End Sub

Sub RemoveGlassStyle()
    Dim ws As Worksheet, shp As Shape, n As Long
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets(SHEET_NAME)
    On Error GoTo 0
    If ws Is Nothing Then Exit Sub

    Dim wasProt As Boolean
    On Error Resume Next
    wasProt = ws.ProtectContents
    ws.Unprotect Password:="p7ss"
    On Error GoTo 0

    For Each shp In ws.Shapes
        If Left$(GetAltText(shp), Len(TAG)) = TAG Then
            RestoreOriginal shp
            n = n + 1
        End If
    Next shp

    On Error Resume Next
    ActiveWindow.DisplayGridlines = True
    On Error GoTo 0
    If wasProt Then ws.Protect Password:="p7ss"
    MsgBox "Glass style removed; " & n & " shape(s) restored.", vbInformation, "Glass Style"
End Sub

' ---- helpers -------------------------------------------------------

Private Function IsStylableShape(ByVal shp As Shape) As Boolean
    On Error Resume Next
    ' Only autoshape rectangles / rounded rectangles (the dashboard
    ' buttons and banners). Skips the background, pictures, etc.
    If shp.Type = msoAutoShape Then
        Dim t As Long: t = shp.AutoShapeType
        IsStylableShape = (t = msoShapeRectangle Or t = msoShapeRoundedRectangle)
    End If
    On Error GoTo 0
End Function

Private Function HasMacro(ByVal shp As Shape) As Boolean
    On Error Resume Next
    HasMacro = (Len(shp.OnAction) > 0)
    On Error GoTo 0
End Function

Private Sub StyleButton(ByVal shp As Shape)
    On Error Resume Next
    shp.AutoShapeType = msoShapeRoundedRectangle
    shp.Adjustments(1) = CORNER
    With shp.Fill
        .Visible = msoTrue
        .Solid
        .ForeColor.RGB = GLASS_WHITE
        .Transparency = BTN_TRANS
    End With
    With shp.Line
        .Visible = msoTrue
        .ForeColor.RGB = RGB(200, 212, 230)
        .Transparency = 0.2
        .Weight = 1
    End With
    shp.SoftEdge.Type = 2
    With shp.Shadow
        .Type = msoShadow25
        .Visible = msoTrue
        .Transparency = 0.72
        .Blur = 8
        .OffsetX = 0: .OffsetY = 2
    End With
    On Error GoTo 0
End Sub

Private Sub StylePanel(ByVal shp As Shape)
    On Error Resume Next
    With shp.Fill
        .Visible = msoTrue
        .Solid
        .ForeColor.RGB = GLASS_WHITE
        .Transparency = PANEL_TRANS
    End With
    With shp.Line
        .Visible = msoTrue
        .ForeColor.RGB = RGB(210, 220, 235)
        .Transparency = 0.3
        .Weight = 0.75
    End With
    shp.SoftEdge.Type = 1
    On Error GoTo 0
End Sub

Private Sub StashOriginal(ByVal shp As Shape)
    On Error Resume Next
    If Left$(GetAltText(shp), Len(TAG)) = TAG Then Exit Sub   ' already stashed
    Dim s As String
    s = TAG & shp.AutoShapeType & "|" & shp.Fill.Visible & "|" & _
        shp.Fill.ForeColor.RGB & "|" & CLng(shp.Fill.Transparency * 1000) & "|" & _
        shp.Line.Visible & "|" & shp.Line.ForeColor.RGB & "|" & CLng(shp.Line.Weight * 100)
    SetAltText shp, s
    On Error GoTo 0
End Sub

Private Sub RestoreOriginal(ByVal shp As Shape)
    On Error Resume Next
    Dim p() As String
    p = Split(GetAltText(shp), "|")
    If UBound(p) < 7 Then Exit Sub

    ' clear glass extras first
    shp.SoftEdge.Type = 0        ' msoSoftEdgeTypeNone
    shp.Shadow.Visible = msoFalse

    shp.AutoShapeType = CLng(p(1))
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
    SetAltText shp, ""
    On Error GoTo 0
End Sub

Private Function GetAltText(ByVal shp As Shape) As String
    On Error Resume Next
    GetAltText = shp.AlternativeText
    On Error GoTo 0
End Function

Private Sub SetAltText(ByVal shp As Shape, ByVal s As String)
    On Error Resume Next
    shp.AlternativeText = s
    On Error GoTo 0
End Sub
