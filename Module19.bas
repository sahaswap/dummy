Attribute VB_Name = "Module19"
'==================================================================
' modBuildFrmExportMode - ONE-TIME setup macro. Run BuildFrmExportMode
' once to programmatically build the frmExportMode UserForm (three
' clearly-labeled buttons + a Cancel) that replaces the old stacked
' Yes/No/Cancel MsgBox chain on the Export Trx File button - same idea
' as frmSearchMode's picker for OSDD Search.
'
' REQUIRES: File > Options > Trust Center > Trust Center Settings >
' Macro Settings > "Trust access to the VBA project object model"
' (same setting the other form builders, e.g. BuildFrmCalendar, need).
' If it errors immediately with "Could not access the VBA project",
' that setting isn't available and the form must be built manually.
'
' Safe to re-run: clears and rebuilds frmExportMode from scratch.
' Delete this module once frmExportMode is built.
'==================================================================
Option Explicit

Sub BuildFrmExportMode()
On Error GoTo TrustFail
Dim vbProj As Object
Set vbProj = Application.VBE.ActiveVBProject
On Error GoTo 0

Dim vbComp As Object
On Error Resume Next
Set vbComp = vbProj.VBComponents("frmExportMode")
On Error GoTo 0

If vbComp Is Nothing Then
Set vbComp = vbProj.VBComponents.Add(3) ' vbext_ct_MSForm
vbComp.Name = "frmExportMode"
End If

vbComp.Properties("Width").Value = 560
vbComp.Properties("Height").Value = 310
vbComp.Properties("Caption").Value = "Choose Export Type"

Dim frm As Object
Set frm = vbComp.Designer

' Clear any existing controls so re-runs start clean.
Dim i As Long
For i = frm.Controls.count - 1 To 0 Step -1
frm.Controls.Remove frm.Controls(i).Name
Next i

' All 4 buttons share ONE size (BTN_W x BTN_H) so nothing looks mismatched.
' Each description label is shorter than its button and vertically CENTERED
' against it (labelTop = buttonTop + (BTN_H - LBL_H)/2), so the button
' caption and its text sit "parallel in the middle" of the same row
' instead of the label text drifting toward the top of a tall box.
Const BTN_W As Single = 140
Const BTN_H As Single = 50
Const LBL_H As Single = 36
Const LBL_OFFSET As Single = (BTN_H - LBL_H) / 2   ' = 7

' Title line
Dim lblTitle As Object
Set lblTitle = frm.Controls.Add("Forms.Label.1", "lblTitle")
lblTitle.Left = 10: lblTitle.Top = 8: lblTitle.Width = 530: lblTitle.Height = 18
lblTitle.caption = "Pick ONE export to run:"
lblTitle.Font.Bold = True
lblTitle.TextAlign = 2 ' fmTextAlignCenter

' ---- Row 1: Legacy ----
Dim btnLegacy As Object, lblLegacy As Object
Set btnLegacy = frm.Controls.Add("Forms.CommandButton.1", "btnLegacy")
btnLegacy.caption = "Legacy": btnLegacy.Left = 10: btnLegacy.Top = 36: btnLegacy.Width = BTN_W: btnLegacy.Height = BTN_H

Set lblLegacy = frm.Controls.Add("Forms.Label.1", "lblLegacy")
lblLegacy.Left = 160: lblLegacy.Top = 36 + LBL_OFFSET: lblLegacy.Width = 380: lblLegacy.Height = LBL_H
lblLegacy.caption = "Dedupe + pivots - the original consolidated transaction export."
lblLegacy.WordWrap = True
lblLegacy.TextAlign = 2 ' fmTextAlignCenter

' ---- Row 2: EN Network ----
Dim btnEN As Object, lblEN As Object
Set btnEN = frm.Controls.Add("Forms.CommandButton.1", "btnEN")
btnEN.caption = "EN Network": btnEN.Left = 10: btnEN.Top = 96: btnEN.Width = BTN_W: btnEN.Height = BTN_H

Set lblEN = frm.Controls.Add("Forms.Label.1", "lblEN")
lblEN.Left = 160: lblEN.Top = 96 + LBL_OFFSET: lblEN.Width = 380: lblEN.Height = LBL_H
lblEN.caption = "Generates BOTH the Alerted/Non-Alerted file AND the Lookback Transactions file in one go."
lblEN.WordWrap = True
lblEN.TextAlign = 2 ' fmTextAlignCenter

' ---- Row 3: Pivot Analysis ----
Dim btnPivot As Object, lblPivot As Object
Set btnPivot = frm.Controls.Add("Forms.CommandButton.1", "btnPivot")
btnPivot.caption = "Pivot Analysis": btnPivot.Left = 10: btnPivot.Top = 156: btnPivot.Width = BTN_W: btnPivot.Height = BTN_H

Set lblPivot = frm.Controls.Add("Forms.Label.1", "lblPivot")
lblPivot.Left = 160: lblPivot.Top = 156 + LBL_OFFSET: lblPivot.Width = 380: lblPivot.Height = LBL_H
lblPivot.caption = "Builds pivot tables from the files in the \Pivot folder. Does not touch ConsolidatedData."
lblPivot.WordWrap = True
lblPivot.TextAlign = 2 ' fmTextAlignCenter

' ---- Cancel - same BTN_W x BTN_H as the other three ----
Dim btnCancel As Object
Set btnCancel = frm.Controls.Add("Forms.CommandButton.1", "btnCancel")
btnCancel.caption = "Cancel": btnCancel.Left = 10: btnCancel.Top = 216: btnCancel.Width = BTN_W: btnCancel.Height = BTN_H

' ---- code-behind ----
Dim codeMod As Object
Set codeMod = vbComp.CodeModule
If codeMod.CountOfLines > 0 Then codeMod.DeleteLines 1, codeMod.CountOfLines

Dim s As String
s = ""
s = s & "Option Explicit" & vbCrLf
s = s & "" & vbCrLf
s = s & "' Export-type picker (replaces the old stacked Yes/No/Cancel MsgBox" & vbCrLf
s = s & "' chain). Same pattern as frmSearchMode: SelectedMode / userCancelled" & vbCrLf
s = s & "' are read by the caller after Show vbModal." & vbCrLf
s = s & "" & vbCrLf
s = s & "Public SelectedMode As String   ' ""LEGACY"" / ""EN"" / ""PIVOT""" & vbCrLf
s = s & "Public userCancelled As Boolean" & vbCrLf
s = s & "" & vbCrLf
s = s & "Private Sub btnLegacy_Click()" & vbCrLf
s = s & "SelectedMode = ""LEGACY""" & vbCrLf
s = s & "userCancelled = False" & vbCrLf
s = s & "Me.Hide" & vbCrLf
s = s & "End Sub" & vbCrLf
s = s & "" & vbCrLf
s = s & "Private Sub btnEN_Click()" & vbCrLf
s = s & "SelectedMode = ""EN""" & vbCrLf
s = s & "userCancelled = False" & vbCrLf
s = s & "Me.Hide" & vbCrLf
s = s & "End Sub" & vbCrLf
s = s & "" & vbCrLf
s = s & "Private Sub btnPivot_Click()" & vbCrLf
s = s & "SelectedMode = ""PIVOT""" & vbCrLf
s = s & "userCancelled = False" & vbCrLf
s = s & "Me.Hide" & vbCrLf
s = s & "End Sub" & vbCrLf
s = s & "" & vbCrLf
s = s & "Private Sub btnCancel_Click()" & vbCrLf
s = s & "userCancelled = True" & vbCrLf
s = s & "Me.Hide" & vbCrLf
s = s & "End Sub" & vbCrLf
s = s & "" & vbCrLf
s = s & "Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)" & vbCrLf
s = s & "If CloseMode = vbFormControlMenu Then userCancelled = True" & vbCrLf
s = s & "End Sub" & vbCrLf
codeMod.AddFromString s

MsgBox "frmExportMode built: Legacy / EN Network / Pivot Analysis + Cancel." & vbCrLf & vbCrLf & _
"You can now delete this modBuildFrmExportMode module - it's done its job.", vbInformation, "Build Complete"
Exit Sub

TrustFail:
MsgBox "Could not access the VBA project object model." & vbCrLf & vbCrLf & _
"Enable 'Trust access to the VBA project object model' in:" & vbCrLf & _
"File > Options > Trust Center > Trust Center Settings > Macro Settings." & vbCrLf & vbCrLf & _
"If that's locked by IT policy, frmExportMode must be built manually.", vbCritical, "Trust Setting Required"
End Sub
