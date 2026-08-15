Option Explicit
'=====================================================================
' modDarkTheme - flips the Sheet1 data area to dark mode so it reads on
' the dark aurora background.
'
' For every cell in the dashboard range it:
'   - removes the fill (No Fill) so the aurora shows through, and
'   - sets the font to a light off-white.
'
' FULLY REVERSIBLE: before touching anything it copies each cell's
' original fill + font colour to a very-hidden backup sheet (_DarkBak).
' RemoveDarkCells reads that back and restores every cell exactly, then
' deletes the backup.
'
'   ApplyDarkCells   - back up + go dark
'   RemoveDarkCells  - restore from backup
'
' >>> RUN ON A COPY first. Only the range below is touched; the hidden
'     master table (row 37+) and other sheets are never affected. <<<
'=====================================================================
Private Const SHEET_NAME As String = "Sheet1"
Private Const RANGE_ADDR As String = "A1:U29"     ' the visible dashboard block
Private Const BAK_SHEET As String = "_DarkBak"
Private Const LIGHT_FONT As Long = 16051428        ' RGB(228, 236, 244) off-white

Sub ApplyDarkCells()
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets(SHEET_NAME)
    On Error GoTo 0
    If ws Is Nothing Then MsgBox SHEET_NAME & " not found.", vbCritical: Exit Sub

    ' guard: don't double-apply (would back up the already-dark state)
    Dim bak As Worksheet
    Set bak = GetBak(False)
    If Not bak Is Nothing Then
        If Len(CStr(bak.Cells(1, 1).Value)) > 0 Then
            MsgBox "Dark cells already applied (a backup exists)." & vbCrLf & _
                   "Run RemoveDarkCells first if you want to re-apply.", vbExclamation, "Dark Theme"
            Exit Sub
        End If
    End If

    Dim wasProt As Boolean
    On Error Resume Next
    wasProt = ws.ProtectContents
    ws.Unprotect Password:="p7ss"
    On Error GoTo 0

    Application.ScreenUpdating = False
    Set bak = GetBak(True)
    bak.Cells.Clear

    Dim rng As Range, c As Range, i As Long
    Set rng = ws.Range(RANGE_ADDR)
    i = 0
    For Each c In rng.Cells
        i = i + 1
        bak.Cells(i, 1).Value = c.Address
        bak.Cells(i, 2).Value = c.Interior.ColorIndex     ' xlNone => was no-fill
        bak.Cells(i, 3).Value = c.Interior.Color
        bak.Cells(i, 4).Value = c.Font.Color
    Next c

    ' go dark: clear fills (aurora shows through) + light text
    rng.Interior.ColorIndex = xlNone
    rng.Font.Color = LIGHT_FONT

    bak.Visible = xlSheetVeryHidden
    If wasProt Then ws.Protect Password:="p7ss"
    Application.ScreenUpdating = True

    MsgBox "Dark mode applied to " & i & " cells (" & RANGE_ADDR & ")." & vbCrLf & vbCrLf & _
           "Run RemoveDarkCells to restore the originals exactly.", vbInformation, "Dark Theme"
End Sub

Sub RemoveDarkCells()
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets(SHEET_NAME)
    On Error GoTo 0
    If ws Is Nothing Then Exit Sub

    Dim bak As Worksheet
    Set bak = GetBak(False)
    If bak Is Nothing Then
        MsgBox "No backup found - nothing to restore.", vbExclamation, "Dark Theme"
        Exit Sub
    End If

    Dim wasProt As Boolean
    On Error Resume Next
    wasProt = ws.ProtectContents
    ws.Unprotect Password:="p7ss"
    On Error GoTo 0

    Application.ScreenUpdating = False
    Dim i As Long, addr As String, ci As Variant
    i = 0
    Do
        i = i + 1
        addr = CStr(bak.Cells(i, 1).Value)
        If Len(addr) = 0 Then Exit Do
        ci = bak.Cells(i, 2).Value
        If ci = xlNone Then
            ws.Range(addr).Interior.ColorIndex = xlNone
        Else
            ws.Range(addr).Interior.Color = bak.Cells(i, 3).Value
        End If
        ws.Range(addr).Font.Color = bak.Cells(i, 4).Value
    Loop

    Application.DisplayAlerts = False
    bak.Delete
    Application.DisplayAlerts = True

    If wasProt Then ws.Protect Password:="p7ss"
    Application.ScreenUpdating = True
    MsgBox "Dark mode removed; " & (i - 1) & " cells restored to their originals.", _
           vbInformation, "Dark Theme"
End Sub

Private Function GetBak(ByVal createIfMissing As Boolean) As Worksheet
    On Error Resume Next
    Set GetBak = ThisWorkbook.Sheets(BAK_SHEET)
    On Error GoTo 0
    If GetBak Is Nothing And createIfMissing Then
        Set GetBak = ThisWorkbook.Sheets.Add
        GetBak.Name = BAK_SHEET
    End If
End Function
