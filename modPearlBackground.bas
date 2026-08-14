Option Explicit
'=====================================================================
' modPearlBackground - mother-of-pearl gradient as Sheet1's BACKGROUND.
'
' Generates a soft pearlescent gradient image and sets it as the sheet's
' background picture, so it sits BEHIND all your cells/data/buttons. This
' is the only way to get a full-sheet backdrop in Excel - worksheet shapes
' always render above cells, so a shape can't go behind your data.
'
' It touches NO cells, rows, values, formulas, dropdowns or buttons.
'
' Caveats (both fine for a working dashboard):
'   - Background pictures do NOT print.
'   - It's a static wash: a flat sheet has no viewing angle, so it can't
'     actually shift color like real nacre - it just looks pearly.
'
'   ApplyPearlBackground  - build the image + set it as the background
'   RemovePearlBackground - clear it and turn gridlines back on
'
' >>> Run on a COPY first. Image export can be finicky on locked-down
'     setups - if the result looks off, tell me and I'll adjust. <<<
'=====================================================================
Private Const SHEET_NAME As String = "Sheet1"

Sub ApplyPearlBackground()
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets(SHEET_NAME)
    On Error GoTo 0
    If ws Is Nothing Then MsgBox SHEET_NAME & " not found.", vbCritical: Exit Sub

    Dim wasProt As Boolean
    On Error Resume Next
    wasProt = ws.ProtectContents
    ws.Unprotect Password:="p7ss"
    On Error GoTo 0

    ws.Activate
    Application.ScreenUpdating = False

    Dim tmp As String
    tmp = Environ$("TEMP") & "\pearl_bg_" & Format(Now, "hhmmss") & ".png"

    Dim shp As Shape, co As ChartObject
    On Error GoTo Fail

    ' 1. Large pearlescent gradient rectangle (soft nacre pastels).
    Set shp = ws.Shapes.AddShape(msoShapeRectangle, 0, 0, 1600, 1000)
    With shp.Fill
        .Visible = msoTrue
        .ForeColor.RGB = RGB(250, 226, 236)
        .OneColorGradient msoGradientDiagonalUp, 1, 1
    End With
    On Error Resume Next          ' GradientStops.Insert is version-sensitive
    With shp.Fill.GradientStops
        .Insert RGB(250, 226, 236), 0#      ' pink
        .Insert RGB(252, 243, 222), 0.25    ' cream
        .Insert RGB(222, 245, 230), 0.5     ' mint
        .Insert RGB(230, 222, 248), 0.75    ' lilac
        .Insert RGB(220, 236, 250), 1#      ' powder blue
    End With
    On Error GoTo Fail
    shp.Line.Visible = msoFalse

    ' 2. Export the shape to PNG via a throwaway chart.
    shp.CopyPicture Appearance:=xlScreen, Format:=xlBitmap
    Set co = ws.ChartObjects.Add(0, 0, shp.Width, shp.Height)
    co.Chart.ChartArea.Format.Line.Visible = msoFalse
    co.Chart.Paste
    co.Chart.Export fileName:=tmp, FilterName:="PNG"
    co.Delete: Set co = Nothing
    shp.Delete: Set shp = Nothing

    ' 3. Set it as the sheet background (sits behind every cell).
    ws.SetBackgroundPicture fileName:=tmp

    On Error Resume Next
    ActiveWindow.DisplayGridlines = False
    On Error GoTo 0

    If wasProt Then ws.Protect Password:="p7ss"
    Application.ScreenUpdating = True
    MsgBox "Mother-of-pearl background applied to " & SHEET_NAME & _
           " (behind your data)." & vbCrLf & vbCrLf & _
           "Run RemovePearlBackground to clear it. It won't print, by design.", _
           vbInformation, "Pearl Background"
    Exit Sub

Fail:
    Dim d As String: d = Err.Description
    On Error Resume Next
    If Not co Is Nothing Then co.Delete
    If Not shp Is Nothing Then shp.Delete
    If wasProt Then ws.Protect Password:="p7ss"
    Application.ScreenUpdating = True
    On Error GoTo 0
    MsgBox "Couldn't apply the pearl background:" & vbCrLf & d & vbCrLf & vbCrLf & _
           "(Image export can be blocked on some locked-down machines - " & _
           "tell me and I'll switch to an embedded-image approach instead.)", _
           vbExclamation, "Pearl Background"
End Sub

Sub RemovePearlBackground()
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets(SHEET_NAME)
    On Error GoTo 0
    If ws Is Nothing Then Exit Sub

    Dim wasProt As Boolean
    On Error Resume Next
    wasProt = ws.ProtectContents
    ws.Unprotect Password:="p7ss"
    ws.Activate
    Application.CommandBars.ExecuteMso "SheetBackgroundDelete"   ' clears bg
    ActiveWindow.DisplayGridlines = True
    If wasProt Then ws.Protect Password:="p7ss"
    On Error GoTo 0
    MsgBox "Pearl background removed; gridlines restored.", vbInformation, "Pearl Background"
End Sub
