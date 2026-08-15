Option Explicit
'=====================================================================
' modDarkTheme - flips the Sheet1 data area to DARK or LIGHT mode.
'
' For every cell in the dashboard range it sets the font colour and
' clears the fill (so the background shows in the gaps), then paints a
' card colour behind each section so the background only peeks through
' the GAPS between sections. The stray "Utilities" label in B16 is
' blended out (its value is kept).
'
' FULLY REVERSIBLE: originals are copied to a very-hidden sheet
' (_DarkBak) first; RemoveDarkCells restores every cell exactly.
'
'   ApplyDarkCells    - dark cards, light text
'   ApplyLightCells   - light cards, dark text
'   RemoveDarkCells   - restore originals
'   ClearStrayDropdowns - delete the stray row-28 dropdowns
'
' >>> RUN ON A COPY first. Only the range below is touched. <<<
'=====================================================================
Private Const SHEET_NAME As String = "Sheet1"
Private Const RANGE_ADDR As String = "A1:U29"
Private Const BAK_SHEET As String = "_DarkBak"

Sub ApplyDarkCells()
    ApplyThemeCells RGB(18, 29, 45), RGB(228, 236, 244), RGB(14, 22, 34), "Dark"
End Sub

Sub ApplyLightCells()
    ApplyThemeCells RGB(255, 255, 255), RGB(40, 54, 78), RGB(255, 255, 255), "Light"
End Sub

Private Sub ApplyThemeCells(ByVal cardColor As Long, ByVal fontColor As Long, _
                            ByVal hideColor As Long, ByVal modeName As String)
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets(SHEET_NAME)
    On Error GoTo 0
    If ws Is Nothing Then MsgBox SHEET_NAME & " not found.", vbCritical: Exit Sub

    ' guard: don't double-apply (would back up the already-themed state)
    Dim bak As Worksheet
    Set bak = GetBak(False)
    If Not bak Is Nothing Then
        If Len(CStr(bak.Cells(1, 1).Value)) > 0 Then
            MsgBox "A theme is already applied (a backup exists)." & vbCrLf & _
                   "Run RemoveDarkCells first, then apply the one you want.", vbExclamation, "Theme"
            Exit Sub
        End If
    End If

    Dim wasProt As Boolean, wbProt As Boolean
    On Error Resume Next
    wasProt = ws.ProtectContents
    ws.Unprotect Password:="p7ss"
    wbProt = ThisWorkbook.ProtectStructure
    ThisWorkbook.Unprotect Password:="p7ss"
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
        bak.Cells(i, 2).Value = c.Interior.ColorIndex
        bak.Cells(i, 3).Value = c.Interior.Color
        bak.Cells(i, 4).Value = c.Font.Color
    Next c

    ' baseline: theme font everywhere, no fill
    rng.Font.Color = fontColor
    rng.Interior.ColorIndex = xlNone

    ' card colour behind each section (background peeks through the gaps)
    ws.Range("G4:T10").Interior.Color = cardColor     ' Alert Related Information
    ws.Range("G12:T14").Interior.Color = cardColor    ' Customer Information
    ws.Range("G16:T23").Interior.Color = cardColor    ' Counterparty Information
    ws.Range("G25:T27").Interior.Color = cardColor    ' Country Risk Rating

    ' blend out the stray "Utilities" label (B16) without deleting it
    ws.Range("B16").Font.Color = hideColor

    bak.Visible = xlSheetVeryHidden
    If wasProt Then ws.Protect Password:="p7ss"
    If wbProt Then ThisWorkbook.Protect Password:="p7ss", Structure:=True
    Application.ScreenUpdating = True

    MsgBox modeName & " mode applied to " & i & " cells." & vbCrLf & vbCrLf & _
           "Run RemoveDarkCells to restore the originals.", vbInformation, "Theme"
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
        MsgBox "No backup found - nothing to restore.", vbExclamation, "Theme"
        Exit Sub
    End If

    Dim wasProt As Boolean, wbProt As Boolean
    On Error Resume Next
    wasProt = ws.ProtectContents
    ws.Unprotect Password:="p7ss"
    wbProt = ThisWorkbook.ProtectStructure
    ThisWorkbook.Unprotect Password:="p7ss"
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

    bak.Cells.Clear
    On Error Resume Next
    bak.Visible = xlSheetVeryHidden
    On Error GoTo 0

    If wasProt Then ws.Protect Password:="p7ss"
    If wbProt Then ThisWorkbook.Protect Password:="p7ss", Structure:=True
    Application.ScreenUpdating = True
    MsgBox "Theme removed; " & (i - 1) & " cells restored to their originals.", _
           vbInformation, "Theme"
End Sub

' Removes the stray dropdowns on row 28 (the Country dropdowns' data
' validation over-extended from row 27 into row 28). Values untouched.
Sub ClearStrayDropdowns()
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets(SHEET_NAME)
    On Error GoTo 0
    If ws Is Nothing Then MsgBox SHEET_NAME & " not found.", vbCritical: Exit Sub

    Dim wasProt As Boolean
    On Error Resume Next
    wasProt = ws.ProtectContents
    ws.Unprotect Password:="p7ss"
    ws.Range("G28:S28").Validation.Delete
    If wasProt Then ws.Protect Password:="p7ss"
    On Error GoTo 0
    MsgBox "Stray dropdowns on row 28 (G28:S28) removed.", vbInformation, "Theme"
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
