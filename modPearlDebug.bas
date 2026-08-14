Option Explicit
'=====================================================================
' modPearlDebug - instruments the pearl-background pipeline step by step
' so we can see EXACTLY where/why it isn't showing.
'
' Import this module, run PearlDiagnostic, and send me the message box
' text (it's also printed to the Immediate window, Ctrl+G).
'
' It reports, in order:
'   - is Sheet1 found, is it protected
'   - do the data cells have a SOLID FILL (a filled cell hides the
'     background - this is one of the two likely culprits)
'   - shape created?  gradient stops applied?
'   - CopyPicture ok?  chart created?  paste ok? (how many shapes landed)
'   - Export ok?  PNG written?  FILE SIZE  (a blank/white image is tiny;
'     a real gradient is much bigger - this tells us if the image is white)
'   - SetBackgroundPicture ok?
'=====================================================================
Private Const SHEET_NAME As String = "Sheet1"

Sub PearlDiagnostic()
    Dim log As String, ws As Worksheet, er As Long
    Dim shp As Shape, co As ChartObject
    Dim wasProt As Boolean

    log = "PEARL BACKGROUND DIAGNOSTIC" & vbCrLf & String(46, "=") & vbCrLf

    On Error Resume Next
    Set ws = ThisWorkbook.Sheets(SHEET_NAME)
    On Error GoTo 0
    If ws Is Nothing Then MsgBox log & "FAIL: " & SHEET_NAME & " not found.": Exit Sub
    log = log & "[OK] Sheet found: tab='" & ws.Name & "'  codename='" & ws.CodeName & "'" & vbCrLf

    ' --- are the cells covering the background with a solid fill? ---
    Dim addrs As Variant, a As Variant, c As Range, filled As Long, n As Long
    addrs = Array("A1", "G4", "J9", "J13", "M20", "P10", "H26", "V5", "H32")
    log = log & "--- cell fills (a SOLID fill hides the sheet background) ---" & vbCrLf
    For Each a In addrs
        Set c = ws.Range(CStr(a)): n = n + 1
        If c.Interior.Pattern = xlNone Then
            log = log & "   " & a & ": no fill (background shows through)" & vbCrLf
        Else
            filled = filled + 1
            log = log & "   " & a & ": FILLED  color=" & c.Interior.Color & vbCrLf
        End If
    Next a
    log = log & "   => " & filled & " of " & n & " sampled cells are filled." & vbCrLf
    If filled >= n - 1 Then _
        log = log & "   !! Almost everything is filled - THAT is why it looks white." & vbCrLf

    On Error Resume Next
    wasProt = ws.ProtectContents
    ws.Unprotect "p7ss"
    On Error GoTo 0
    ws.Activate

    Dim tmp As String
    tmp = Environ$("TEMP") & "\pearl_dbg.png"
    On Error Resume Next: Kill tmp: On Error GoTo 0

    ' --- build gradient shape ---
    On Error Resume Next
    Set shp = ws.Shapes.AddShape(msoShapeRectangle, 0, 0, 1600, 1000)
    er = Err.Number: On Error GoTo 0
    If shp Is Nothing Then MsgBox log & "FAIL at AddShape (err " & er & ")": GoTo Done
    log = log & "[OK] gradient shape created" & vbCrLf

    On Error Resume Next
    With shp.Fill
        .Visible = msoTrue
        .ForeColor.RGB = RGB(233, 120, 176)
        .OneColorGradient msoGradientDiagonalUp, 1, 1
    End With
    With shp.Fill.GradientStops
        .Insert RGB(233, 120, 176), 0#
        .Insert RGB(240, 176, 96), 0.18
        .Insert RGB(96, 200, 150), 0.36
        .Insert RGB(80, 182, 214), 0.54
        .Insert RGB(150, 110, 228), 0.72
        .Insert RGB(224, 110, 178), 0.86
        .Insert RGB(96, 160, 228), 1#
    End With
    shp.Line.Visible = msoFalse
    er = Err.Number: On Error GoTo 0
    log = log & "[i] gradient stops = " & shp.Fill.GradientStops.count & " (err " & er & ")" & vbCrLf

    ' --- copy picture ---
    On Error Resume Next: Err.Clear
    shp.CopyPicture Appearance:=xlScreen, Format:=xlBitmap
    er = Err.Number: On Error GoTo 0
    log = log & IIf(er = 0, "[OK]", "[FAIL]") & " CopyPicture (err " & er & ")" & vbCrLf

    ' --- throwaway chart + paste ---
    On Error Resume Next
    Set co = ws.ChartObjects.Add(0, 0, shp.Width, shp.Height)
    er = Err.Number: On Error GoTo 0
    If co Is Nothing Then MsgBox log & "FAIL at ChartObjects.Add (err " & er & ")": GoTo Done
    log = log & "[OK] temp chart created" & vbCrLf

    On Error Resume Next: Err.Clear
    DoEvents
    co.Chart.Paste
    er = Err.Number: On Error GoTo 0
    Dim cnt As Long: cnt = -1
    On Error Resume Next: cnt = co.Chart.Shapes.count: On Error GoTo 0
    log = log & IIf(er = 0, "[OK]", "[FAIL]") & " Chart.Paste (err " & er & _
          ") - shapes landed in chart: " & cnt & vbCrLf
    If cnt = 0 Then log = log & "   !! Nothing pasted -> exported image will be WHITE." & vbCrLf

    ' --- export ---
    On Error Resume Next: Err.Clear
    co.Chart.Export fileName:=tmp, FilterName:="PNG"
    er = Err.Number: On Error GoTo 0
    log = log & IIf(er = 0, "[OK]", "[FAIL]") & " Chart.Export (err " & er & ")" & vbCrLf

    co.Delete: Set co = Nothing
    shp.Delete: Set shp = Nothing

    If Dir(tmp) = "" Then
        log = log & "[FAIL] no PNG written to " & tmp & vbCrLf
    Else
        Dim fl As Long: fl = FileLen(tmp)
        log = log & "[i] PNG size = " & fl & " bytes  " & _
              IIf(fl < 4000, "(TINY -> almost certainly a blank/white image)", _
              "(looks substantial -> image probably has the gradient)") & vbCrLf
        On Error Resume Next: Err.Clear
        ws.SetBackgroundPicture fileName:=tmp
        er = Err.Number: On Error GoTo 0
        log = log & IIf(er = 0, "[OK]", "[FAIL]") & " SetBackgroundPicture (err " & er & ")" & vbCrLf
    End If

Done:
    On Error Resume Next
    If Not co Is Nothing Then co.Delete
    If Not shp Is Nothing Then shp.Delete
    If wasProt Then ws.Protect "p7ss"
    On Error GoTo 0
    log = log & String(46, "=") & vbCrLf & "Send me these lines (esp. any [FAIL], the fill count, and the PNG size)."
    Debug.Print log
    MsgBox log, vbInformation, "Pearl Diagnostic"
End Sub
