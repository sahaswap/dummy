Option Explicit
'=====================================================================
' modSectionFrames - rounded "card" outlines around each dashboard
' section (plus an outer box), to match the demo's rounded boxes.
'
' Each frame is a rounded rectangle with NO FILL and a soft border, so
' its interior is CLICK-THROUGH - you can still select/edit the cells
' inside. It only draws the rounded outline; the dark card colour comes
' from the cell fills underneath (ApplyDarkCells).
'
'   AddSectionFrames    - draw the frames
'   RemoveSectionFrames - delete them (idempotent, safe to re-run)
'=====================================================================
Private Const SHEET_NAME As String = "Sheet1"
Private Const PFX As String = "SECFRAME_"

Sub AddSectionFrames()
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

    RemoveFrames ws                       ' idempotent

    AddFrame ws, "B2:U28", PFX & "Outer", 0.02, 0.3      ' outer box
    AddFrame ws, "G4:T10", PFX & "Alert", 0.06, 0.12
    AddFrame ws, "G12:T14", PFX & "Customer", 0.09, 0.12
    AddFrame ws, "G16:T23", PFX & "Counterparty", 0.05, 0.12
    AddFrame ws, "G25:T27", PFX & "Country", 0.09, 0.12

    If wasProt Then ws.Protect Password:="p7ss"
    MsgBox "Section frames added." & vbCrLf & vbCrLf & _
           "They're no-fill, so you can still click the cells inside. " & _
           "Run RemoveSectionFrames to undo.", vbInformation, "Frames"
End Sub

Sub RemoveSectionFrames()
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets(SHEET_NAME)
    On Error GoTo 0
    If ws Is Nothing Then Exit Sub

    Dim wasProt As Boolean
    On Error Resume Next
    wasProt = ws.ProtectContents
    ws.Unprotect Password:="p7ss"
    On Error GoTo 0

    RemoveFrames ws
    If wasProt Then ws.Protect Password:="p7ss"
    MsgBox "Section frames removed.", vbInformation, "Frames"
End Sub

Private Sub AddFrame(ByVal ws As Worksheet, ByVal addr As String, ByVal nm As String, _
                     ByVal rnd As Single, ByVal lineTrans As Single)
    Dim r As Range, shp As Shape, pad As Single
    pad = 3
    Set r = ws.Range(addr)
    Set shp = ws.Shapes.AddShape(msoShapeRoundedRectangle, _
              r.Left - pad, r.Top - pad, r.Width + 2 * pad, r.Height + 2 * pad)
    shp.Name = nm
    shp.Fill.Visible = msoFalse            ' NO fill -> interior is click-through
    With shp.Line
        .Visible = msoTrue
        .ForeColor.RGB = RGB(120, 145, 180) ' blue-grey edge (reads on light & dark)
        .Transparency = lineTrans
        .Weight = 1.5
    End With
    On Error Resume Next
    shp.Adjustments(1) = rnd               ' corner roundness
    On Error GoTo 0
    shp.ZOrder msoSendToBack               ' behind banners/buttons, above cells
End Sub

Private Sub RemoveFrames(ByVal ws As Worksheet)
    Dim i As Long
    For i = ws.Shapes.count To 1 Step -1
        If Left$(ws.Shapes(i).Name, Len(PFX)) = PFX Then ws.Shapes(i).Delete
    Next i
End Sub
