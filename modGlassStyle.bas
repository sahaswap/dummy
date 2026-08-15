Option Explicit
'=====================================================================
' modGlassStyle - DARK frosted-glass restyle for the Sheet1 dashboard,
' to sit on the dark aurora background (ApplyPearlBackground).
'
' Your buttons and banners are real shapes, so it restyles them in
' place - dark translucent fill (aurora shimmers through), soft edge,
' and LIGHT text so it reads on the dark theme. Rounds the buttons.
'
'   ApplyGlassStyle  - buttons -> rounded dark frosted glass, light text;
'                      banners/title -> dark frosted, light text;
'                      "Reset" keeps a soft danger-red label.
'   RemoveGlassStyle - restores every shape EXACTLY (type, fill, line AND
'                      font/colour), from a stash in each shape's AltText.
'
' Idempotent (never re-stashes). Touches ONLY shape formatting - no
' cells, values, data or macros.
'=====================================================================
Private Const SHEET_NAME As String = "Sheet1"
Private Const TAG As String = "GLASSORIG|"
Private Const CORNER As Single = 0.35          ' button corner roundness (0..0.5)
Private Const BTN_TRANS As Single = 0.28       ' button fill transparency
Private Const PANEL_TRANS As Single = 0.34     ' banner fill transparency
Private Const TITLE_TRANS As Single = 0.5      ' title box transparency

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

    On Error Resume Next
    ActiveWindow.DisplayGridlines = False
    On Error GoTo 0
    If wasProt Then ws.Protect Password:="p7ss"

    MsgBox "Dark glass style applied to " & n & " shape(s)." & vbCrLf & vbCrLf & _
           "Run RemoveGlassStyle to undo (restores type, fill, line and font).", _
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

' ---- classification ------------------------------------------------

Private Function IsStylableShape(ByVal shp As Shape) As Boolean
    On Error Resume Next
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

Private Function IsTitle(ByVal shp As Shape) As Boolean
    On Error Resume Next
    If Len(shp.OnAction) = 0 And shp.TextFrame.HasText Then
        IsTitle = (InStr(1, Trim$(shp.TextFrame.Characters.Text), "Beta", vbTextCompare) = 1)
    End If
    On Error GoTo 0
End Function

' ---- styling (dark frosted glass, light text) ----------------------

Private Sub StyleButton(ByVal shp As Shape)
    On Error Resume Next
    shp.AutoShapeType = msoShapeRoundedRectangle
    shp.Adjustments(1) = CORNER
    With shp.Fill
        .Visible = msoTrue: .Solid
        .ForeColor.RGB = RGB(18, 28, 42)       ' dark slate
        .Transparency = BTN_TRANS
    End With
    With shp.Line
        .Visible = msoTrue
        .ForeColor.RGB = RGB(120, 196, 205)    ' soft teal edge
        .Transparency = 0.55
        .Weight = 1
    End With
    shp.SoftEdge.Type = 2
    With shp.Shadow
        .Type = msoShadow25: .Visible = msoTrue
        .Transparency = 0.65: .Blur = 9: .OffsetX = 0: .OffsetY = 2
    End With
    If shp.TextFrame.HasText Then
        With shp.TextFrame.Characters.Font
            If InStr(1, shp.TextFrame.Characters.Text, "Reset", vbTextCompare) > 0 Then
                .Color = RGB(255, 150, 150)    ' soft danger red
            Else
                .Color = RGB(232, 240, 246)    ' light
            End If
        End With
    End If
    On Error GoTo 0
End Sub

Private Sub StylePanel(ByVal shp As Shape)
    On Error Resume Next
    With shp.Fill
        .Visible = msoTrue: .Solid
        .ForeColor.RGB = RGB(14, 22, 34)
        .Transparency = PANEL_TRANS
    End With
    With shp.Line
        .Visible = msoTrue
        .ForeColor.RGB = RGB(90, 150, 165)
        .Transparency = 0.6
        .Weight = 0.75
    End With
    shp.SoftEdge.Type = 1
    If shp.TextFrame.HasText Then
        shp.TextFrame.Characters.Font.Color = RGB(232, 240, 246)
    End If
    On Error GoTo 0
End Sub

Private Sub StyleTitle(ByVal shp As Shape)
    On Error Resume Next
    With shp.Fill
        .Visible = msoTrue: .Solid
        .ForeColor.RGB = RGB(14, 22, 34)
        .Transparency = TITLE_TRANS
    End With
    shp.Line.Visible = msoFalse
    shp.SoftEdge.Type = 1
    If shp.TextFrame.HasText Then
        shp.TextFrame.Characters.Font.Color = RGB(236, 243, 249)
    End If
    On Error GoTo 0
End Sub

' ---- stash / restore (self-contained via shape AltText) ------------

Private Sub StashOriginal(ByVal shp As Shape)
    On Error Resume Next
    If Left$(GetAltText(shp), Len(TAG)) = TAG Then Exit Sub   ' already stashed

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
    SetAltText shp, s
    On Error GoTo 0
End Sub

Private Sub RestoreOriginal(ByVal shp As Shape)
    On Error Resume Next
    Dim p() As String
    p = Split(GetAltText(shp), "|")
    If UBound(p) < 7 Then Exit Sub

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
