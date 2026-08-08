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

' ---- Two grey panels (rows), from the Action Panel layout ----
' Group 1 buttons live in the TOP panel, group 2 in the BOTTOM panel.
' P1_TOP starts on the first GREY row (below the blue header), so the
' first button never lands on the dark-blue title bar.
Const P1_TOP As Long = 6         ' top grey panel: first row (below blue header)
Const P1_BOT As Long = 16        ' top grey panel: last row
Const P2_TOP As Long = 18        ' bottom grey panel: first row
Const P2_BOT As Long = 27        ' bottom grey panel: last row (grey extended to here)

' ---- Horizontal placement (centred) ----
Const BTN_LEFT_COL As String = "C"   ' buttons centred within these columns
Const BTN_RIGHT_COL As String = "E"
Const PANEL_LEFT_COL As String = "B" ' grey panel's left column (for background fill)
Const H_MARGIN As Single = 8         ' left+right margin inside the B..E panel (centres the button)
Const FILL_RATIO As Single = 0.62    ' button height as a fraction of its vertical slot

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

' Map each button to its INTENDED slot by the macro it runs, so the
' order is exactly Start/Export/OSDD/Narrative | Rename/PDFMerge/
' Reset/Warm-Up regardless of where the shapes currently sit.
Dim keyMacro(1 To 8) As String
keyMacro(1) = "Start_Button_Create_Folders"   ' Start
keyMacro(2) = "Consolidated_AML_Workflow"      ' Export Trx File
keyMacro(3) = "SearchAndSavePDF_Direct"        ' OSDD Search
keyMacro(4) = "ExportToWord"                   ' Generate Narrative
keyMacro(5) = "Run_Mass_Rename"                ' Rename
keyMacro(6) = "Trigger_PAD_Merge_Flow"         ' PDF Merge
keyMacro(7) = "ClearForm"                      ' Reset
keyMacro(8) = "WarmUpWorkerProfiles"           ' Profile Warm-Up (created below)

Dim ordered(1 To 8) As Shape
Dim s As Shape, oa As String, k As Long, resetBtn As Shape
For Each s In ws.Shapes
oa = ""
On Error Resume Next
oa = s.OnAction
On Error GoTo Fail
If Len(oa) > 0 Then
For k = 1 To 8
If InStr(oa, keyMacro(k)) > 0 Then
Set ordered(k) = s
If k = 7 Then Set resetBtn = s
Exit For
End If
Next k
End If
Next s

If resetBtn Is Nothing Then
' fall back to any group-2 button as the style source
If Not ordered(6) Is Nothing Then Set resetBtn = ordered(6)
If resetBtn Is Nothing And Not ordered(5) Is Nothing Then Set resetBtn = ordered(5)
End If
If resetBtn Is Nothing Then
On Error Resume Next
ws.Protect Password:=WB_PASSWORD
ThisWorkbook.Protect Password:=WB_PASSWORD, Structure:=True, Windows:=False
On Error GoTo 0
MsgBox "Couldn't find the existing bottom-group buttons to match against.", vbExclamation
Exit Sub
End If

' Font from Reset, so every button matches.
Dim fSize As Single, fBold As Boolean, fName As String, fColor As Long
On Error Resume Next
With resetBtn.TextFrame.Characters.Font
fSize = .Size: fBold = .Bold: fName = .Name: fColor = .Color
End With
On Error GoTo Fail

' Create the warm-up button (exact-style copy of Reset).
Dim nb As Shape
Set nb = resetBtn.Duplicate
nb.Name = "btnWarmUp"
On Error Resume Next
nb.TextFrame.Characters.Text = CAPTION
On Error GoTo Fail
nb.OnAction = "WarmUpWorkerProfiles"
Set ordered(8) = nb

' Centred horizontal geometry: symmetric H_MARGIN inside the FULL
' grey panel (PANEL_LEFT_COL "B" .. BTN_RIGHT_COL "E") - so the
' button is centred across the whole panel, not just C..E (which
' left column B as extra margin and pushed the buttons right).
Dim bLeft As Single, bWidth As Single
bLeft = ws.Range(PANEL_LEFT_COL & "1").Left + H_MARGIN
bWidth = (ws.Range(BTN_RIGHT_COL & "1").Left + ws.Range(BTN_RIGHT_COL & "1").Width) _
- ws.Range(PANEL_LEFT_COL & "1").Left - 2 * H_MARGIN

' Clean the bottom panel: an earlier version copied a bordered row
' down this range, which drew a horizontal line at every row (the
' "gridlines" you saw). Clear the INTERIOR borders so it reads as a
' single clean panel like the top one. The panel's outer edge
' borders and its fill are left untouched.
On Error Resume Next
With ws.Range(PANEL_LEFT_COL & P2_TOP & ":" & BTN_RIGHT_COL & P2_BOT)
.Borders(xlInsideHorizontal).LineStyle = xlNone
.Borders(xlInsideVertical).LineStyle = xlNone
End With
On Error GoTo Fail

' Lay out the two groups: buttons 1-4 in the top panel, 5-8 in the
' bottom panel - each centred in its slot, evenly spaced.
PlaceGroup ws, ordered, 1, 4, P1_TOP, P1_BOT, bLeft, bWidth, FILL_RATIO, fSize, fBold, fName, fColor
PlaceGroup ws, ordered, 5, 8, P2_TOP, P2_BOT, bLeft, bWidth, FILL_RATIO, fSize, fBold, fName, fColor

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

MsgBox "Laid out 4 + 4 buttons, centred, inside the two grey panels " & _
"(rows " & P1_TOP & "-" & P1_BOT & " and " & P2_TOP & "-" & P2_BOT & "), " & _
"added '" & CAPTION & "', and removed the Legends & Notes tab.", _
vbInformation, "Done"
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

' Evenly places ordered(firstIdx..lastIdx) inside the row band
' topRow..botRow, each button centred in its vertical slot, at the
' given left/width, with a uniform height derived from fillRatio.
Private Sub PlaceGroup(ByVal ws As Worksheet, ByRef ordered() As Shape, _
ByVal firstIdx As Long, ByVal lastIdx As Long, _
ByVal topRow As Long, ByVal botRow As Long, _
ByVal bLeft As Single, ByVal bWidth As Single, ByVal fillRatio As Single, _
ByVal fSize As Single, ByVal fBold As Boolean, ByVal fName As String, ByVal fColor As Long)
On Error Resume Next

Dim boxTop As Single, boxBot As Single, boxH As Single
boxTop = ws.Range("A" & topRow).Top
boxBot = ws.Range("A" & botRow).Top + ws.Range("A" & botRow).Height
boxH = boxBot - boxTop

Dim count As Long
count = lastIdx - firstIdx + 1
If count < 1 Then Exit Sub

Dim pitch As Single, btnH As Single
pitch = boxH / count
btnH = pitch * fillRatio

Dim i As Long, slot As Long
slot = 0
For i = firstIdx To lastIdx
If Not ordered(i) Is Nothing Then
With ordered(i)
.Left = bLeft
.Width = bWidth
.Height = btnH
.Top = boxTop + slot * pitch + (pitch - btnH) / 2
.TextFrame.Characters.Font.Size = fSize
.TextFrame.Characters.Font.Bold = fBold
.TextFrame.Characters.Font.Name = fName
.TextFrame.Characters.Font.Color = fColor
End With
End If
slot = slot + 1
Next i
End Sub

'==================================================================
' UnhideColumnsAfterT  -  unhides the hidden "extension" columns
' immediately to the right of T (U, V) that you couldn't get rid of.
'
' The sheet is protected, which is why unhiding them was blocked -
' this unprotects, unhides, then re-protects. It only UNHIDES; it
' does NOT delete, because columns further right (W onward) hold
' backend/helper data that other formulas point at - deleting would
' shift those and break references. Once you can see U:V, if they're
' genuinely empty spacers you want gone, tell me and I'll delete a
' specific confirmed-empty column safely.
'
' UNHIDE_RANGE controls what gets unhidden - widen it if the column
' you mean is further out than V.
'==================================================================
Sub UnhideColumnsAfterT()
Const UNHIDE_RANGE As String = "U:V"

Dim ws As Worksheet
On Error Resume Next
Set ws = ThisWorkbook.Sheets("Sheet1")
On Error GoTo 0
If ws Is Nothing Then MsgBox "Sheet1 not found.", vbCritical: Exit Sub

On Error Resume Next
ThisWorkbook.Unprotect Password:=WB_PASSWORD
ws.Unprotect Password:=WB_PASSWORD
On Error GoTo 0

ws.Columns(UNHIDE_RANGE).EntireColumn.Hidden = False

' Report what's now visible so you can decide about deleting.
Dim uContent As String, vContent As String
uContent = Trim(CStr(ws.Range("U1").Value)) & Trim(CStr(ws.Range("U19").Value))
vContent = Trim(CStr(ws.Range("V1").Value)) & Trim(CStr(ws.Range("V19").Value))

On Error Resume Next
ws.Protect Password:=WB_PASSWORD
ThisWorkbook.Protect Password:=WB_PASSWORD, Structure:=True, Windows:=False
On Error GoTo 0

MsgBox "Unhid columns " & UNHIDE_RANGE & " (just right of T)." & vbCrLf & vbCrLf & _
"U appears " & IIf(Len(uContent) = 0, "empty", "to contain data") & ", " & _
"V appears " & IIf(Len(vContent) = 0, "empty", "to contain data") & "." & vbCrLf & vbCrLf & _
"I did NOT delete anything - columns further right hold backend data " & _
"that formulas reference. If U/V are empty and you want them deleted, " & _
"confirm and I'll remove the specific column safely.", _
vbInformation, "Columns Unhidden"
End Sub
