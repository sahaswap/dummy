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
' AddWarmUpButtonToSheet1  -  ONE-TIME. Re-lays-out the Sheet1
' left-pane button stack so the warm-up button fits in as the last
' one below Reset, and (per request) deletes the now-redundant
' "Legends & Notes" tab.
'
' What it does:
'   - DUPLICATES an existing Sheet1 button (OSDD) so the new one's
'     fill / corner style / font copy exactly, then re-captions it
'     and points it at WarmUpWorkerProfiles.
'   - Collects ALL the left-pane macro buttons, makes them a uniform
'     size (slightly smaller than now), aligns them to one column,
'     and re-stacks all of them - the 7 existing + the new one -
'     evenly from the top. This closes the wasted empty gap that's
'     currently between "Generate Narrative" and "Rename", which is
'     what frees the room for the 8th button without crowding the
'     Country Name field to the right.
'   - Caps the stack so it can never run past the row-34 divider
'     band (auto-shrinks height if a machine's row heights differ).
'   - Deletes the "Legends & Notes" sheet.
'
' Safe to re-run.
'==================================================================
Sub AddWarmUpButtonToSheet1()
Const CAPTION As String = "Profile Warm-Up"
Const BTN_HEIGHT As Single = 26      ' uniform height (was ~29 - "brought down")
Const BTN_GAP As Single = 14         ' uniform gap between buttons
Const BOTTOM_MARGIN As Single = 8    ' keep this clear of the row-34 divider

Dim ws As Worksheet
On Error Resume Next
Set ws = ThisWorkbook.Sheets("Sheet1")
On Error GoTo 0
If ws Is Nothing Then MsgBox "Sheet1 not found.", vbCritical: Exit Sub

' Drop workbook structure + sheet content protection so we can move
' shapes and delete a sheet; both restored at the end.
On Error Resume Next
ThisWorkbook.Unprotect Password:=WB_PASSWORD
ws.Unprotect Password:=WB_PASSWORD
On Error GoTo Fail

' Idempotent: clear any earlier copy first, so re-runs start clean.
On Error Resume Next
ws.Shapes("btnWarmUp").Delete
On Error GoTo Fail

' Collect the existing left-pane macro buttons (anything with a macro
' assigned), and grab the OSDD one as the style reference.
Dim s As Shape, ref As Shape, oa As String
Dim btns() As Shape, nBtn As Long
ReDim btns(1 To 50)
nBtn = 0
For Each s In ws.Shapes
oa = ""
On Error Resume Next
oa = s.OnAction
On Error GoTo Fail
If Len(oa) > 0 Then
nBtn = nBtn + 1
Set btns(nBtn) = s
If InStr(oa, "SearchAndSavePDF_Direct") > 0 Then Set ref = s
End If
Next s

If nBtn = 0 Then
On Error Resume Next
ws.Protect Password:=WB_PASSWORD
ThisWorkbook.Protect Password:=WB_PASSWORD, Structure:=True, Windows:=False
On Error GoTo 0
MsgBox "No existing macro buttons found on Sheet1 to match against.", vbExclamation
Exit Sub
End If
If ref Is Nothing Then Set ref = btns(1)   ' fallback if OSDD was renamed

' Uniform geometry taken from what's already there: align to the
' reference's left edge, use the widest existing button's width so no
' caption (e.g. "Generate Narrative") gets cramped.
Dim uLeft As Single, uWidth As Single, i As Long
uLeft = ref.Left
uWidth = 0
For i = 1 To nBtn
If btns(i).Width > uWidth Then uWidth = btns(i).Width
Next i

' Capture the reference font so we can re-apply it after re-texting.
Dim fSize As Single, fBold As Boolean, fName As String, fColor As Long
With ref.TextFrame.Characters.Font
fSize = .Size: fBold = .Bold: fName = .Name: fColor = .Color
End With

' Build the new button as an exact-style copy, then add it to the
' collection so it's laid out with the rest.
Dim nb As Shape
Set nb = ref.Duplicate
nb.Name = "btnWarmUp"
nb.TextFrame.Characters.Text = CAPTION
With nb.TextFrame.Characters.Font
.Size = fSize: .Bold = fBold: .Name = fName: .Color = fColor
End With
nb.OnAction = "WarmUpWorkerProfiles"
nBtn = nBtn + 1
Set btns(nBtn) = nb   ' new button appended -> ends up last (below Reset)

' Order the collection top-to-bottom by current Top (insertion sort).
' The new button is given a large key so it always sorts LAST, i.e.
' below Reset.
Dim j As Long, tmp As Shape, kt As Single
Dim keyTop() As Single
ReDim keyTop(1 To nBtn)
For i = 1 To nBtn
If btns(i) Is nb Then keyTop(i) = 1000000! Else keyTop(i) = btns(i).Top
Next i
For i = 2 To nBtn
Set tmp = btns(i): kt = keyTop(i)
j = i - 1
Do While j >= 1
If keyTop(j) > kt Then
Set btns(j + 1) = btns(j): keyTop(j + 1) = keyTop(j): j = j - 1
Else
Exit Do
End If
Loop
Set btns(j + 1) = tmp: keyTop(j + 1) = kt
Next i

' Vertical envelope: from the top of the current top button down to
' just above the row-34 divider band. Auto-shrink height if that
' many buttons at BTN_HEIGHT+BTN_GAP wouldn't fit.
Dim envTop As Single, envBottom As Single, availH As Single
envTop = btns(1).Top
envBottom = envTop + nBtn * (BTN_HEIGHT + BTN_GAP)   ' default if row 34 unknown
On Error Resume Next
envBottom = ws.Range("A34").Top - BOTTOM_MARGIN
On Error GoTo Fail
availH = envBottom - envTop

Dim useH As Single, useGap As Single
useH = BTN_HEIGHT: useGap = BTN_GAP
If nBtn * useH + (nBtn - 1) * useGap > availH Then
' Too tall for the space - shrink height and gap proportionally.
Dim scale As Single
scale = availH / (nBtn * BTN_HEIGHT + (nBtn - 1) * BTN_GAP)
If scale < 0.4 Then scale = 0.4   ' never collapse to nothing
useH = BTN_HEIGHT * scale
useGap = BTN_GAP * scale
End If

' Apply uniform size/position + font to every button in order.
Dim yPos As Single
yPos = envTop
For i = 1 To nBtn
With btns(i)
.Left = uLeft
.Width = uWidth
.Height = useH
.Top = yPos
On Error Resume Next
.TextFrame.Characters.Font.Size = fSize
.TextFrame.Characters.Font.Bold = fBold
.TextFrame.Characters.Font.Name = fName
.TextFrame.Characters.Font.Color = fColor
On Error GoTo Fail
End With
yPos = yPos + useH + useGap
Next i

' Delete the now-redundant Legends & Notes sheet.
On Error Resume Next
Application.DisplayAlerts = False
ThisWorkbook.Sheets(LEGENDS_SHEET_NAME).Delete
Application.DisplayAlerts = True
On Error GoTo Fail

Application.CutCopyMode = False

RestoreAndExit:
On Error Resume Next
ws.Protect Password:=WB_PASSWORD
ThisWorkbook.Protect Password:=WB_PASSWORD, Structure:=True, Windows:=False
On Error GoTo 0

MsgBox "Re-laid the Sheet1 button stack (" & nBtn & " buttons, uniform size) with " & _
"'" & CAPTION & "' added below Reset, and removed the Legends & Notes tab.", _
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
