'==================================================================
' modBuildLegendsSheet  -  ONE-TIME setup macro.
'
' Run BuildLegendsAndNotesSheet once to create a "Legends & Notes"
' sheet positioned just before Sheet1. The sheet holds:
'   - a button that opens the two worker-profile Edge windows
'     VISIBLY against a Google search (manual cookie/consent
'     warm-up, the same idea as WarmEdgeProfileCookies but on
'     demand - a human accepts consent / handles any check in a
'     real window, so the headless workers start warm)
'   - a large empty Notes block to fill in later
'
' Unlike the UserForm builders, this needs NO Trust Center setting -
' it only uses the normal Excel object model (Sheets.Add + a Form
' Control button), not the VBA project model.
'
' Safe to re-run: deletes and rebuilds the sheet from scratch.
'
' Delete this builder after running it if you like - but keep
' WarmUpWorkerProfiles below, that's the sub the button calls.
'==================================================================
Option Explicit

Private Const LEGENDS_SHEET_NAME As String = "Legends & Notes"
Private Const WORKER_COUNT As Long = 2   ' OSINT_EdgeWorker_0 .. _(n-1)
Private Const WARMUP_QUERY As String = "test"
Private Const WB_PASSWORD As String = "p7ss"

Sub BuildLegendsAndNotesSheet()
On Error GoTo Fail

' The workbook normally runs structure-protected (see Module9) - drop
' it so we can add/delete a sheet, restore it at the end.
On Error Resume Next
ThisWorkbook.Unprotect Password:=WB_PASSWORD
On Error GoTo Fail

Dim ws As Worksheet

' Remove any prior copy so re-running is clean.
On Error Resume Next
Application.DisplayAlerts = False
ThisWorkbook.Sheets(LEGENDS_SHEET_NAME).Delete
Application.DisplayAlerts = True
On Error GoTo Fail

' Insert just before Sheet1 (fall back to first position if the tab
' named "Sheet1" isn't found).
Dim beforeSheet As Object
On Error Resume Next
Set beforeSheet = ThisWorkbook.Sheets("Sheet1")
On Error GoTo Fail
If beforeSheet Is Nothing Then Set beforeSheet = ThisWorkbook.Sheets(1)

Set ws = ThisWorkbook.Sheets.Add(Before:=beforeSheet)
ws.Name = LEGENDS_SHEET_NAME

' ---- Layout ----
ws.Cells.Font.Name = "Calibri"
ws.Columns("A").ColumnWidth = 2.5
ws.Columns("B:H").ColumnWidth = 16

' Title
With ws.Range("B2")
.Value = "Legends & Notes"
.Font.Size = 20
.Font.Bold = True
.Font.Color = RGB(0, 70, 127)
End With

' Warm-up section header
With ws.Range("B4")
.Value = "Profile Warm-Up"
.Font.Size = 12
.Font.Bold = True
.Font.Color = RGB(255, 255, 255)
.Interior.Color = RGB(0, 70, 127)
End With
ws.Range("B4:H4").Merge

' Explanation
With ws.Range("B6")
.Value = "Opens " & WORKER_COUNT & " visible Edge window(s) - one per worker profile " & _
"(OSINT_EdgeWorker_0.." & (WORKER_COUNT - 1) & ") - on a Google search. " & _
"Accept the cookie consent and complete any on-screen check in each window, " & _
"then close them. The headless searches will reuse those warmed-up profiles."
.WrapText = True
.Font.Size = 10
.VerticalAlignment = xlTop
End With
ws.Range("B6:H8").Merge

' Button (Form Control - no Trust Center setting needed)
Dim btn As Object
Set btn = ws.Buttons.Add(ws.Range("B10").Left, ws.Range("B10").Top, 230, 34)
btn.Caption = "Open Warm-Up Browser Windows"
btn.Name = "btnWarmUp"
btn.OnAction = "WarmUpWorkerProfiles"
btn.Font.Size = 10
btn.Font.Bold = True

' Notes section header
With ws.Range("B13")
.Value = "Notes"
.Font.Size = 12
.Font.Bold = True
.Font.Color = RGB(255, 255, 255)
.Interior.Color = RGB(0, 70, 127)
End With
ws.Range("B13:H13").Merge

' Empty, bordered notes block to fill in later.
Dim notesBlock As Range
Set notesBlock = ws.Range("B14:H40")
With notesBlock
.Interior.Color = RGB(255, 255, 240)   ' faint cream so the area reads as "writeable"
.VerticalAlignment = xlTop
.WrapText = True
.Font.Size = 10
End With
With notesBlock.Borders
.LineStyle = xlContinuous
.Color = RGB(180, 180, 180)
.Weight = xlThin
End With

ws.Range("B1").Select

' Restore the workbook's normal structure protection.
On Error Resume Next
ThisWorkbook.Protect Password:=WB_PASSWORD, Structure:=True, Windows:=False
On Error GoTo 0

MsgBox "'" & LEGENDS_SHEET_NAME & "' sheet built (just before Sheet1), with the " & _
"warm-up button and an empty Notes area." & vbCrLf & vbCrLf & _
"You can delete the modBuildLegendsSheet builder now - just keep the " & _
"WarmUpWorkerProfiles sub, which is what the button runs.", _
vbInformation, "Build Complete"
Exit Sub

Fail:
On Error Resume Next
Application.DisplayAlerts = True
ThisWorkbook.Protect Password:=WB_PASSWORD, Structure:=True, Windows:=False
On Error GoTo 0
MsgBox "Could not build the sheet." & vbCrLf & "Error " & Err.Number & ": " & Err.Description, _
vbCritical, "Build Failed"
End Sub

'------------------------------------------------------------------
' Button target. Opens one visible Edge window per worker profile on
' a plain Google search so the analyst can accept consent / clear
' any on-screen check by hand, warming those profiles' cookies for
' the headless batch. Launches Edge directly (not via cmd/start) so
' there's no console flash and no shell-escaping of the URL's '&'.
'------------------------------------------------------------------
Sub WarmUpWorkerProfiles()
On Error Resume Next

Dim edgeExe As String, firstToken As String
edgeExe = ResolveEdgeExe()
If Len(edgeExe) > 0 Then
firstToken = """" & edgeExe & """"
Else
firstToken = "msedge"   ' fall back to PATH / App Paths resolution
End If

Dim wsh As Object: Set wsh = CreateObject("WScript.Shell")
Dim url As String
url = "https://www.google.com/search?q=" & WARMUP_QUERY & "&num=100&hl=en"

Dim w As Long, profilePath As String, cmd As String, opened As Long
For w = 0 To WORKER_COUNT - 1
profilePath = Environ$("LOCALAPPDATA") & "\OSINT_EdgeWorker_" & w
cmd = firstToken & _
" --user-data-dir=""" & profilePath & """" & _
" --window-size=1024,768" & _
" --window-position=" & (w * 80) & ",60" & _
" """ & url & """"
wsh.Run cmd, 1, False   ' 1 = SW_SHOWNORMAL - a real, visible window
opened = opened + 1
Next w

MsgBox "Opened " & opened & " warm-up window(s)." & vbCrLf & vbCrLf & _
"In each one: accept the cookie/consent prompt and complete any on-screen " & _
"check, then close the window. That leaves the worker profiles warmed up " & _
"for the next search run.", vbInformation, "Warm-Up Windows Opened"

On Error GoTo 0
End Sub

' Full path to msedge.exe, or "" if none of the usual spots have it.
Private Function ResolveEdgeExe() As String
Dim candidates As Variant, p As Variant
candidates = Array( _
"C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe", _
"C:\Program Files\Microsoft\Edge\Application\msedge.exe", _
Environ$("LOCALAPPDATA") & "\Microsoft\Edge\Application\msedge.exe")
For Each p In candidates
If Dir(CStr(p)) <> "" Then
ResolveEdgeExe = CStr(p)
Exit Function
End If
Next p
ResolveEdgeExe = ""
End Function

'==================================================================
' AddWarmUpButtonToSheet1  -  ONE-TIME.
'
' Ensures the "Profile Warm-Up" button exists, then lays out ALL the
' left-pane buttons uniformly inside the cell box C..E / top..row 27,
' and deletes the "Legends & Notes" tab.
'
' The layout is derived from the actual GRID (live Range().Left/.Top/
' .Width), not from the buttons' current positions - so it both fits
' everything neatly to E27 AND repairs any earlier displacement:
' every button is re-anchored to the same box, so none can stick out
' past column E onto/off the grey Action-Panel cells.
'
' Tunables below control the box edges and how tall the buttons are
' relative to their slots. Safe to re-run.
'==================================================================
Sub AddWarmUpButtonToSheet1()
Const CAPTION As String = "Profile Warm-Up"
Const LEFT_COL As String = "C"       ' button block's left column
Const RIGHT_COL As String = "E"      ' button block's right column (-> E27)
Const BOTTOM_ROW As Long = 27        ' block bottom (the "E27" you gave)
Const INSET As Single = 3            ' padding inside the box edges
Const FILL_RATIO As Single = 0.72    ' button height as a fraction of its slot

Dim ws As Worksheet
On Error Resume Next
Set ws = ThisWorkbook.Sheets("Sheet1")
On Error GoTo 0
If ws Is Nothing Then MsgBox "Sheet1 not found.", vbCritical: Exit Sub

On Error Resume Next
ThisWorkbook.Unprotect Password:=WB_PASSWORD
ws.Unprotect Password:=WB_PASSWORD
On Error GoTo Fail

' Start clean so a re-run doesn't stack duplicates.
On Error Resume Next
ws.Shapes("btnWarmUp").Delete
On Error GoTo Fail

' Collect the existing macro buttons; grab one as the style source
' and track the topmost, which anchors the top of the block.
Dim s As Shape, oa As String, ref As Shape
Dim btns() As Shape, nBtn As Long, minTop As Single
ReDim btns(1 To 50)
nBtn = 0: minTop = 1E+9
For Each s In ws.Shapes
oa = ""
On Error Resume Next
oa = s.OnAction
On Error GoTo Fail
If Len(oa) > 0 Then
nBtn = nBtn + 1
Set btns(nBtn) = s
If s.Top < minTop Then minTop = s.Top
If InStr(oa, "ClearForm") > 0 Then Set ref = s   ' Reset = good style source
End If
Next s

If nBtn = 0 Then
On Error Resume Next
ws.Protect Password:=WB_PASSWORD
ThisWorkbook.Protect Password:=WB_PASSWORD, Structure:=True, Windows:=False
On Error GoTo 0
MsgBox "No existing buttons found on Sheet1.", vbExclamation
Exit Sub
End If
If ref Is Nothing Then Set ref = btns(1)

' Font from the style source, so every button matches.
Dim fSize As Single, fBold As Boolean, fName As String, fColor As Long
On Error Resume Next
With ref.TextFrame.Characters.Font
fSize = .Size: fBold = .Bold: fName = .Name: fColor = .Color
End With
On Error GoTo Fail

' Create the warm-up button (exact-style copy) and add to the set.
Dim nb As Shape
Set nb = ref.Duplicate
nb.Name = "btnWarmUp"
On Error Resume Next
nb.TextFrame.Characters.Text = CAPTION
On Error GoTo Fail
nb.OnAction = "WarmUpWorkerProfiles"
nBtn = nBtn + 1
Set btns(nBtn) = nb

' Order top-to-bottom, forcing the warm-up button LAST.
Dim i As Long, j As Long, tmp As Shape, kt As Single
Dim keyTop() As Single
ReDim keyTop(1 To nBtn)
For i = 1 To nBtn
If btns(i) Is nb Then keyTop(i) = 1E+9 Else keyTop(i) = btns(i).Top
Next i
For i = 2 To nBtn
Set tmp = btns(i): kt = keyTop(i): j = i - 1
Do While j >= 1
If keyTop(j) > kt Then
Set btns(j + 1) = btns(j): keyTop(j + 1) = keyTop(j): j = j - 1
Else
Exit Do
End If
Loop
Set btns(j + 1) = tmp: keyTop(j + 1) = kt
Next i

' Box from the live grid: columns LEFT_COL..RIGHT_COL, top = current
' topmost button, bottom = bottom of BOTTOM_ROW. All measured in real
' points so it honours this machine's actual row heights/col widths.
Dim boxLeft As Single, boxRight As Single, boxTop As Single, boxBot As Single
boxLeft = ws.Range(LEFT_COL & "1").Left + INSET
boxRight = ws.Range(RIGHT_COL & "1").Left + ws.Range(RIGHT_COL & "1").Width - INSET
boxTop = minTop
boxBot = ws.Range("A" & BOTTOM_ROW).Top + ws.Range("A" & BOTTOM_ROW).Height - INSET

Dim boxW As Single, boxH As Single, pitch As Single, btnH As Single
boxW = boxRight - boxLeft
boxH = boxBot - boxTop
If boxH < nBtn * 8 Then boxH = nBtn * 8   ' guard against a silly-small box
pitch = boxH / nBtn
btnH = pitch * FILL_RATIO

' Apply uniform geometry + font to every button, centred in its slot.
For i = 1 To nBtn
With btns(i)
.Left = boxLeft
.Width = boxW
.Height = btnH
.Top = boxTop + (i - 1) * pitch + (pitch - btnH) / 2
On Error Resume Next
.TextFrame.Characters.Font.Size = fSize
.TextFrame.Characters.Font.Bold = fBold
.TextFrame.Characters.Font.Name = fName
.TextFrame.Characters.Font.Color = fColor
On Error GoTo Fail
End With
Next i

' Delete the now-redundant Legends & Notes sheet.
On Error Resume Next
Application.DisplayAlerts = False
ThisWorkbook.Sheets(LEGENDS_SHEET_NAME).Delete
Application.DisplayAlerts = True
On Error GoTo Fail

Application.CutCopyMode = False

On Error Resume Next
ws.Protect Password:=WB_PASSWORD
ThisWorkbook.Protect Password:=WB_PASSWORD, Structure:=True, Windows:=False
On Error GoTo 0

MsgBox "Laid out " & nBtn & " buttons uniformly inside " & LEFT_COL & ".." & _
RIGHT_COL & "/row " & BOTTOM_ROW & ", added '" & CAPTION & "', and removed the " & _
"Legends & Notes tab.", vbInformation, "Done"
Exit Sub

Fail:
On Error Resume Next
Application.DisplayAlerts = True
ws.Protect Password:=WB_PASSWORD
ThisWorkbook.Protect Password:=WB_PASSWORD, Structure:=True, Windows:=False
On Error GoTo 0
MsgBox "Couldn't finish the layout." & vbCrLf & "Error " & Err.Number & ": " & _
Err.Description, vbCritical, "Failed"
End Sub
